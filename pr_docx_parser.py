"""pr_docx_parser.py - Zero-dependency OpenXML (.docx) parser for PR requisitions.

Uses Python standard library only (zipfile + xml.etree.ElementTree).
"""

import datetime
import io
import re
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_DIACRITICS_RE = re.compile(r"[\u0640\u064b-\u065f\u0670]")

UNIT_MAP: Dict[str, str] = {
    "عدد": "No",
    "قطعة": "No",
    "قطعه": "No",
    "حبة": "No",
    "حبه": "No",
    "كيلوجرام": "Kg",
    "كيلو جرام": "Kg",
    "كجم": "Kg",
    "كج": "Kg",
    "متر": "m",
    "متر طولي": "m",
    "م.ط": "m",
    "متر مربع": "m2",
    "م2": "m2",
    "م²": "m2",
    "متر مكعب": "m3",
    "م3": "m3",
    "م³": "m3",
    "لتر": "L",
    "طقم": "Set",
    "مجموعة": "Set",
    "مجموعه": "Set",
}

_FOOTER_STOP_KEYWORDS = (
    "approved material submittal",
    "الميزانيه تسمح",
    "سلطات الاعتماد",
)

_FOOTER_SKIP_KEYWORDS = (
    "approved material submittal",
    "ms ref",
    "تاريخ الطلب",
    "الميزانيه تسمح",
    "سلطات الاعتماد",
    "يعتمد",
    "المدير العام",
    "مدير المشروع",
    "توقيع",
    "التوقيع",
    "prepared by",
    "approved by",
    "checked by",
)


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _to_ascii_digits(text: str) -> str:
    return (text or "").translate(_ARABIC_DIGITS)


def _norm_arabic(text: str) -> str:
    s = _DIACRITICS_RE.sub("", text or "")
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي").replace("ة", "ه")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _clean_text(text: str) -> str:
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in (text or "").splitlines()]
    return "\n".join(ln for ln in lines if ln)


def _extract_p_text(p_elem: ET.Element) -> str:
    parts: List[str] = []
    t_tag = _w("t")
    br_tag = _w("br")
    cr_tag = _w("cr")
    tab_tag = _w("tab")
    for elem in p_elem.iter():
        if elem.tag == t_tag and elem.text:
            parts.append(elem.text)
        elif elem.tag in (br_tag, cr_tag):
            parts.append("\n")
        elif elem.tag == tab_tag:
            parts.append(" ")
    return _clean_text("".join(parts))


def _extract_tc_text(tc_elem: ET.Element) -> str:
    p_tag = _w("p")
    paras = [_extract_p_text(p) for p in tc_elem.findall(p_tag)]
    return "\n".join(p for p in paras if p)


def _parse_table(tbl_elem: ET.Element) -> List[Dict[str, Any]]:
    tr_tag = _w("tr")
    tc_tag = _w("tc")
    tc_pr_tag = _w("tcPr")
    grid_span_tag = _w("gridSpan")
    v_merge_tag = _w("vMerge")
    val_attr = _w("val")

    parsed_rows: List[Dict[str, Any]] = []
    for tr in tbl_elem.findall(tr_tag):
        cells: List[str] = []
        spans: List[Tuple[int, int, str]] = []
        col_idx = 0
        for tc in tr.findall(tc_tag):
            tc_pr = tc.find(tc_pr_tag)
            span = 1
            is_vmerge_continue = False
            if tc_pr is not None:
                gs = tc_pr.find(grid_span_tag)
                if gs is not None:
                    raw_span = gs.get(val_attr) or gs.get("val") or "1"
                    try:
                        span = max(1, int(raw_span))
                    except ValueError:
                        span = 1
                vm = tc_pr.find(v_merge_tag)
                if vm is not None:
                    vm_val = vm.get(val_attr) or vm.get("val") or ""
                    if vm_val.lower() != "restart":
                        is_vmerge_continue = True

            text = "" if is_vmerge_continue else _extract_tc_text(tc)
            spans.append((col_idx, span, text))
            cells.append(text)
            for _ in range(span - 1):
                cells.append("")
            col_idx += span

        parsed_rows.append({"cells": cells, "spans": spans})
    return parsed_rows


def normalize_unit(unit_raw: str) -> str:
    cleaned = re.sub(r"\s+", " ", (unit_raw or "").strip())
    if not cleaned:
        return ""
    return UNIT_MAP.get(cleaned, cleaned)


def parse_date_to_iso(raw: str) -> str:
    s = _to_ascii_digits(raw or "").strip()
    if not s:
        return ""

    m_iso = re.search(r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b", s)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        try:
            return datetime.date(y, m, d).isoformat()
        except ValueError:
            pass

    m_dmy = re.search(r"\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\b", s)
    if m_dmy:
        d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        try:
            return datetime.date(y, m, d).isoformat()
        except ValueError:
            pass

    return ""


def _parse_quantity(raw: str) -> Tuple[Any, Optional[str]]:
    s = _to_ascii_digits(raw or "").strip()
    if not s:
        return "", None
    num_str = s.replace(",", "").strip()
    try:
        val = float(num_str)
        if val.is_integer():
            return int(val), None
        return val, None
    except ValueError:
        return raw.strip(), f"Unparseable quantity '{raw.strip()}'"


def _extract_after_label(
    text: str,
    label_patterns: Tuple[str, ...],
    single_line: bool = True,
) -> Optional[str]:
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    for line_idx, ln in enumerate(lines):
        norm_ln = _norm_arabic(ln)
        for pat in label_patterns:
            if re.search(pat, norm_ln):
                if ":" in ln:
                    after = ln.split(":", 1)[1].strip()
                    if after:
                        return after
                if line_idx + 1 < len(lines):
                    return lines[line_idx + 1] if single_line else "\n".join(lines[line_idx + 1 :])
                return ""
    return None


def _find_adjacent_value(cells: List[str], idx: int, single_line: bool = True) -> str:
    for offset in (1, -1):
        nbr = idx + offset
        if 0 <= nbr < len(cells) and cells[nbr].strip():
            val = cells[nbr].strip()
            if single_line:
                lines = [ln.strip() for ln in val.splitlines() if ln.strip()]
                return lines[0] if lines else ""
            return val
    return ""


def _extract_header_metadata(
    tables: List[List[Dict[str, Any]]],
    paragraphs: List[str],
) -> Dict[str, str]:
    header = {"docNo": "", "issuedDate": "", "msRef": "", "title": ""}

    doc_no_pats = (r"رقم\s*اشعار\s*الاحتياج", r"رقم\s*الاشعار", r"\bpr\s*no\.?\b", r"\bpr\s*number\b")
    date_pats = (r"تاريخ\s*ورود\s*الاشعار", r"تاريخ\s*الطلب", r"تاريخ\s*الاشعار", r"\bissued\s*date\b", r"\bpr\s*date\b", r"\bdate\b")
    ms_pats = (r"approved\s*material\s*submittal", r"\bms\s*ref\b", r"material\s*submittal", r"\brfi\s*ref\b")
    title_pats = (r"^title\b", r"^subject\b", r"^الموضوع\b")

    ms_refs: List[str] = []

    def _add_ms_ref(val: str):
        for ln in (val or "").splitlines():
            cleaned = ln.strip()
            if cleaned and cleaned not in ms_refs:
                ms_refs.append(cleaned)

    for tbl in tables:
        for row in tbl:
            cells = row["cells"]
            for idx, cell in enumerate(cells):
                if not cell:
                    continue
                if not header["docNo"]:
                    inline = _extract_after_label(cell, doc_no_pats, single_line=True)
                    if inline is not None:
                        header["docNo"] = inline or _find_adjacent_value(cells, idx, single_line=True)
                if not header["issuedDate"]:
                    inline = _extract_after_label(cell, date_pats, single_line=True)
                    if inline is not None:
                        candidate = inline or _find_adjacent_value(cells, idx, single_line=True)
                        parsed_date = parse_date_to_iso(candidate)
                        if parsed_date:
                            header["issuedDate"] = parsed_date
                if not header["title"]:
                    inline = _extract_after_label(cell, title_pats, single_line=True)
                    if inline is not None:
                        header["title"] = inline or _find_adjacent_value(cells, idx, single_line=True)

                inline_ms = _extract_after_label(cell, ms_pats, single_line=False)
                if inline_ms is not None:
                    if inline_ms:
                        _add_ms_ref(inline_ms)
                    else:
                        for j in range(idx + 1, len(cells)):
                            if cells[j].strip():
                                _add_ms_ref(cells[j])

    for p in paragraphs:
        if not p:
            continue
        if not header["docNo"]:
            inline = _extract_after_label(p, doc_no_pats, single_line=True)
            if inline:
                header["docNo"] = inline
        if not header["issuedDate"]:
            inline = _extract_after_label(p, date_pats, single_line=True)
            if inline:
                parsed_date = parse_date_to_iso(inline)
                if parsed_date:
                    header["issuedDate"] = parsed_date
        if not header["title"]:
            inline = _extract_after_label(p, title_pats, single_line=True)
            if inline:
                header["title"] = inline
        inline_ms = _extract_after_label(p, ms_pats, single_line=False)
        if inline_ms:
            _add_ms_ref(inline_ms)

    header["msRef"] = "\n".join(ms_refs)
    return header


def _match_header_columns(cells: List[str]) -> Optional[Dict[str, int]]:
    col_map: Dict[str, int] = {}
    for idx, raw in enumerate(cells):
        norm = _norm_arabic(raw)
        if not norm:
            continue
        if norm in ("م", "sr", "sr.", "s/n", "sn", "no", "no.", "#") and "sr" not in col_map:
            col_map["sr"] = idx
        elif any(k in norm for k in ("بيان المطلوب", "description", "item description", "material description")) or norm in ("item", "الوصف", "الصنف"):
            if "desc" not in col_map:
                col_map["desc"] = idx
        elif any(k in norm for k in ("الكميه", "quantity")) or norm in ("qty", "qty."):
            if "qty" not in col_map:
                col_map["qty"] = idx
        elif any(k in norm for k in ("الوحده", "uom")) or norm in ("unit", "units"):
            if "unit" not in col_map:
                col_map["unit"] = idx
        elif any(k in norm for k in ("ملاحظات", "remarks", "notes")) and "remarks" not in col_map:
            col_map["remarks"] = idx

    if "desc" in col_map and any(k in col_map for k in ("qty", "unit", "sr")):
        return col_map
    return None


def _is_footer_stop(row_norm: str) -> bool:
    return any(kw in row_norm for kw in _FOOTER_STOP_KEYWORDS)


def _is_footer_skip(row_norm: str) -> bool:
    return any(kw in row_norm for kw in _FOOTER_SKIP_KEYWORDS)


def _extract_items_from_table(
    rows: List[Dict[str, Any]],
    header_idx: int,
    col_map: Dict[str, int],
) -> Tuple[List[Dict[str, Any]], List[str]]:
    items: List[Dict[str, Any]] = []
    warnings: List[str] = []

    sr_col = col_map.get("sr")
    desc_col = col_map["desc"]
    qty_col = col_map.get("qty")
    unit_col = col_map.get("unit")
    remarks_col = col_map.get("remarks")

    for r_offset, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        cells = row["cells"]
        spans = row["spans"]

        def _get(col: Optional[int]) -> str:
            if col is None or col >= len(cells):
                return ""
            return cells[col].strip()

        non_empty = [c.strip() for c in cells if c.strip() and c.strip() not in ("-", "—", "_")]
        if not non_empty:
            continue

        row_norm = _norm_arabic(" ".join(non_empty))
        if _is_footer_stop(row_norm):
            break
        if _is_footer_skip(row_norm):
            continue

        sr_val = _get(sr_col)
        desc_val = _get(desc_col)
        qty_val = _get(qty_col)
        unit_val = _get(unit_col)
        remarks_val = _get(remarks_col)

        # Detect merged section header spanning across the description column (or whole row)
        merged_header_text = ""
        if not desc_val and not qty_val and not unit_val:
            for start_col, span_len, text in spans:
                t = text.strip()
                if t and span_len >= 2 and (start_col <= desc_col < start_col + span_len or len(non_empty) == 1):
                    if not re.fullmatch(r"\d+", _to_ascii_digits(t)):
                        merged_header_text = t
                        break

        if merged_header_text:
            items.append({"row_type": "header", "item_name": merged_header_text})
            continue

        if desc_val and not sr_val and not qty_val and not unit_val:
            items.append({"row_type": "header", "item_name": desc_val})
            continue

        if not desc_val and (qty_val or unit_val):
            warnings.append(
                f"Row {r_offset}: missing item description (qty={qty_val!r}, unit={unit_val!r})"
            )
            continue

        if desc_val and (sr_val or qty_val or unit_val):
            parsed_qty, qty_warn = _parse_quantity(qty_val)
            if qty_warn:
                warnings.append(f"Row {r_offset}: {qty_warn} for item '{desc_val}'")
            items.append({
                "row_type": "item",
                "item_name": desc_val,
                "unit": normalize_unit(unit_val),
                "quantity": parsed_qty,
                "remarks": remarks_val,
            })

    return items, warnings


def parse_pr_docx(raw_bytes: bytes) -> Dict[str, Any]:
    """Parse a Gas Chill PR .docx file into header metadata, items, and warnings."""
    if not raw_bytes:
        raise ValueError("Empty DOCX payload")

    MAX_XML_BYTES = 15 * 1024 * 1024  # 15 MB uncompressed cap for word/document.xml
    try:
        with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
            try:
                info = zf.getinfo("word/document.xml")
            except KeyError as exc:
                raise ValueError("Missing word/document.xml in DOCX archive") from exc
            if info.file_size > MAX_XML_BYTES:
                raise ValueError("word/document.xml exceeds maximum allowed size")
            with zf.open(info, "r") as xml_fp:
                xml_bytes = xml_fp.read(MAX_XML_BYTES + 1)
            if len(xml_bytes) > MAX_XML_BYTES:
                raise ValueError("word/document.xml exceeds maximum allowed size")
    except zipfile.BadZipFile as exc:
        raise ValueError("Invalid DOCX archive") from exc

    xml_lower = xml_bytes.lower()
    if b"<!doctype" in xml_lower or b"<!entity" in xml_lower:
        raise ValueError("XML DOCTYPE/ENTITY declarations are not allowed")

    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise ValueError("Malformed word/document.xml") from exc

    body = root.find(_w("body"))
    search_root = body if body is not None else root

    tables: List[List[Dict[str, Any]]] = [
        _parse_table(tbl) for tbl in search_root.findall(f".//{_w('tbl')}")
    ]
    paragraphs: List[str] = [
        _extract_p_text(p) for p in search_root.findall(_w("p"))
    ]

    header = _extract_header_metadata(tables, paragraphs)

    all_items: List[Dict[str, Any]] = []
    warnings: List[str] = []
    found_items_table = False

    for tbl_rows in tables:
        for idx, row in enumerate(tbl_rows):
            col_map = _match_header_columns(row["cells"])
            if col_map is not None:
                found_items_table = True
                tbl_items, tbl_warnings = _extract_items_from_table(tbl_rows, idx, col_map)
                all_items.extend(tbl_items)
                warnings.extend(tbl_warnings)
                break

    if not found_items_table:
        warnings.append("No requisition items table header found in document.")

    return {
        "header": header,
        "items": all_items,
        "warnings": warnings,
    }
