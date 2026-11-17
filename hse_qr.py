"""Pure-Python ISO/IEC 18004 QR Code SVG generator with optional segno fast-path."""

from html import escape
import io


# GF(256) with primitive polynomial 0x11D for QR Reed-Solomon ECC
_GF_EXP = [0] * 512
_GF_LOG = [0] * 256
_val = 1
for _i in range(255):
    _GF_EXP[_i] = _val
    _GF_LOG[_val] = _i
    _val <<= 1
    if _val & 0x100:
        _val ^= 0x11D
for _i in range(255, 512):
    _GF_EXP[_i] = _GF_EXP[_i - 255]

# ISO/IEC 18004 Level M parameters for Versions 1..8:
# (version, size, total_data_bytes, ec_per_block, group1_blocks, group1_data, group2_blocks, group2_data, align_centers)
_VERSION_SPECS_M = (
    (1, 21, 16, 10, 1, 16, 0, 0, ()),
    (2, 25, 28, 16, 1, 28, 0, 0, (6, 18)),
    (3, 29, 44, 26, 1, 44, 0, 0, (6, 22)),
    (4, 33, 64, 18, 2, 32, 0, 0, (6, 26)),
    (5, 37, 86, 24, 2, 43, 0, 0, (6, 30)),
    (6, 41, 108, 16, 4, 27, 0, 0, (6, 34)),
    (7, 45, 124, 18, 4, 31, 0, 0, (6, 22, 38)),
    (8, 49, 154, 22, 2, 38, 2, 39, (6, 24, 42)),
)


def _gf_mul(a, b):
    if a == 0 or b == 0:
        return 0
    return _GF_EXP[_GF_LOG[a] + _GF_LOG[b]]


def _rs_generator(degree):
    poly = [1]
    for i in range(degree):
        nxt = [0] * (len(poly) + 1)
        root = _GF_EXP[i]
        for j, coef in enumerate(poly):
            nxt[j] ^= coef
            nxt[j + 1] ^= _gf_mul(coef, root)
        poly = nxt
    return poly


def _rs_remainder(data, degree):
    gen = _rs_generator(degree)
    res = list(data) + [0] * degree
    for i in range(len(data)):
        coef = res[i]
        if coef != 0:
            for j in range(1, len(gen)):
                res[i + j] ^= _gf_mul(gen[j], coef)
    return res[len(data):]


def _encode_byte_payload(raw_bytes):
    spec = None
    needed_bits = 4 + 8 + len(raw_bytes) * 8
    for candidate in _VERSION_SPECS_M:
        if candidate[2] * 8 >= needed_bits:
            spec = candidate
            break
    if spec is None:
        raise ValueError("QR payload exceeds supported Version 8 Byte-mode capacity")

    total_data = spec[2]
    bits = [0, 1, 0, 0]  # Byte mode indicator
    length = len(raw_bytes)
    for shift in range(7, -1, -1):
        bits.append((length >> shift) & 1)
    for byte in raw_bytes:
        for shift in range(7, -1, -1):
            bits.append((byte >> shift) & 1)

    cap_bits = total_data * 8
    term_len = min(4, cap_bits - len(bits))
    bits.extend([0] * term_len)
    while len(bits) % 8 != 0:
        bits.append(0)

    data_bytes = []
    for i in range(0, len(bits), 8):
        b = 0
        for bit in bits[i:i + 8]:
            b = (b << 1) | bit
        data_bytes.append(b)

    pad_cycle = (0xEC, 0x11)
    idx = 0
    while len(data_bytes) < total_data:
        data_bytes.append(pad_cycle[idx & 1])
        idx += 1

    _, _, _, ec_len, g1_blk, g1_sz, g2_blk, g2_sz, _ = spec
    data_blocks = []
    ec_blocks = []
    pos = 0
    for count, bsize in ((g1_blk, g1_sz), (g2_blk, g2_sz)):
        for _ in range(count):
            blk = data_bytes[pos:pos + bsize]
            pos += bsize
            data_blocks.append(blk)
            ec_blocks.append(_rs_remainder(blk, ec_len))

    stream = []
    max_data = max(len(b) for b in data_blocks)
    for col in range(max_data):
        for blk in data_blocks:
            if col < len(blk):
                stream.append(blk[col])
    for col in range(ec_len):
        for blk in ec_blocks:
            stream.append(blk[col])

    bitstream = []
    for byte in stream:
        for shift in range(7, -1, -1):
            bitstream.append((byte >> shift) & 1)
    return spec, bitstream


def _format_bits_ecc_m(mask_id):
    data = (0 << 3) | mask_id  # ECC Level M = 00
    d = data << 10
    for i in range(4, -1, -1):
        if d & (1 << (i + 10)):
            d ^= 0x537 << i
    return ((data << 10) | d) ^ 0x5412


def _version_bits(version):
    if version < 7:
        return None
    d = version << 12
    for i in range(5, -1, -1):
        if d & (1 << (i + 12)):
            d ^= 0x1F25 << i
    return (version << 12) | d


def _build_base_matrix(spec):
    version, size, _, _, _, _, _, _, align = spec
    mat = [[None] * size for _ in range(size)]
    func = [[False] * size for _ in range(size)]

    def set_mod(r, c, dark):
        mat[r][c] = bool(dark)
        func[r][c] = True

    def place_finder(r0, c0):
        for dr in range(-1, 8):
            for dc in range(-1, 8):
                r, c = r0 + dr, c0 + dc
                if 0 <= r < size and 0 <= c < size:
                    in_outer = 0 <= dr <= 6 and 0 <= dc <= 6
                    on_border = dr in (0, 6) or dc in (0, 6)
                    in_center = 2 <= dr <= 4 and 2 <= dc <= 4
                    set_mod(r, c, in_outer and (on_border or in_center))

    place_finder(0, 0)
    place_finder(0, size - 7)
    place_finder(size - 7, 0)

    for i in range(8, size - 8):
        dark = (i % 2 == 0)
        set_mod(6, i, dark)
        set_mod(i, 6, dark)

    for r in align:
        for c in align:
            if func[r][c]:
                continue
            for dr in range(-2, 3):
                for dc in range(-2, 3):
                    dark = (abs(dr) == 2 or abs(dc) == 2 or (dr == 0 and dc == 0))
                    set_mod(r + dr, c + dc, dark)

    set_mod(4 * version + 9, 8, True)

    # Reserve format & version info areas
    for i in range(9):
        if not func[8][i]:
            set_mod(8, i, False)
        if not func[i][8]:
            set_mod(i, 8, False)
    for i in range(size - 8, size):
        set_mod(8, i, False)
    for i in range(size - 7, size):
        set_mod(i, 8, False)

    ver_info = _version_bits(version)
    if ver_info is not None:
        for i in range(18):
            dark = bool((ver_info >> i) & 1)
            r, c = i // 3, size - 11 + (i % 3)
            set_mod(r, c, dark)
            set_mod(c, r, dark)

    return mat, func


def _apply_data_and_mask(spec, mat_base, func, bitstream):
    size = spec[1]
    mat = [list(row) for row in mat_base]
    bit_idx = 0
    total_bits = len(bitstream)
    upward = True
    col = size - 1
    while col > 0:
        if col == 6:
            col -= 1
        rows = range(size - 1, -1, -1) if upward else range(size)
        for r in rows:
            for dc in (0, -1):
                c = col + dc
                if not func[r][c]:
                    bit = bitstream[bit_idx] if bit_idx < total_bits else 0
                    bit_idx += 1
                    mat[r][c] = bool(bit)
        upward = not upward
        col -= 2

    mask_preds = (
        lambda r, c: (r + c) % 2 == 0,
        lambda r, c: r % 2 == 0,
        lambda r, c: c % 3 == 0,
        lambda r, c: (r + c) % 3 == 0,
        lambda r, c: ((r // 2) + (c // 3)) % 2 == 0,
        lambda r, c: ((r * c) % 2 + (r * c) % 3) == 0,
        lambda r, c: (((r * c) % 2 + (r * c) % 3) % 2) == 0,
        lambda r, c: (((r + c) % 2 + (r * c) % 3) % 2) == 0,
    )

    best_mat = None
    best_penalty = None
    for mask_id, pred in enumerate(mask_preds):
        cand = [
            [
                (not mat[r][c] if (not func[r][c] and pred(r, c)) else bool(mat[r][c]))
                for c in range(size)
            ]
            for r in range(size)
        ]
        fmt = _format_bits_ecc_m(mask_id)
        for i in range(6):
            cand[8][i] = bool((fmt >> (14 - i)) & 1)
        cand[8][7] = bool((fmt >> 8) & 1)
        cand[8][8] = bool((fmt >> 7) & 1)
        cand[7][8] = bool((fmt >> 6) & 1)
        for i in range(6):
            cand[5 - i][8] = bool((fmt >> (5 - i)) & 1)
        for i in range(7):
            cand[size - 1 - i][8] = bool((fmt >> (14 - i)) & 1)
        for i in range(8):
            cand[8][size - 8 + i] = bool((fmt >> (7 - i)) & 1)

        penalty = _evaluate_penalty(cand)
        if best_penalty is None or penalty < best_penalty:
            best_penalty = penalty
            best_mat = cand
    return best_mat


def _evaluate_penalty(mat):
    size = len(mat)
    score = 0
    for r in range(size):
        run = 1
        for c in range(1, size):
            if mat[r][c] == mat[r][c - 1]:
                run += 1
            else:
                if run >= 5:
                    score += run - 2
                run = 1
        if run >= 5:
            score += run - 2
    for c in range(size):
        run = 1
        for r in range(1, size):
            if mat[r][c] == mat[r - 1][c]:
                run += 1
            else:
                if run >= 5:
                    score += run - 2
                run = 1
        if run >= 5:
            score += run - 2
    for r in range(size - 1):
        for c in range(size - 1):
            v = mat[r][c]
            if mat[r][c + 1] == v and mat[r + 1][c] == v and mat[r + 1][c + 1] == v:
                score += 3
    p1 = (True, False, True, True, True, False, True, False, False, False, False)
    p2 = (False, False, False, False, True, False, True, True, True, False, True)
    for r in range(size):
        row = mat[r]
        for c in range(size - 10):
            sl = tuple(row[c:c + 11])
            if sl == p1 or sl == p2:
                score += 40
    for c in range(size):
        for r in range(size - 10):
            sl = tuple(mat[r + k][c] for k in range(11))
            if sl == p1 or sl == p2:
                score += 40
    dark = sum(1 for r in mat for cell in r if cell)
    pct = (dark * 100.0) / (size * size)
    score += int(abs(pct - 50.0) // 5) * 10
    return score


def build_qr_matrix(text):
    """Return a boolean NxN QR code matrix for text."""
    raw = str(text or "").encode("utf-8")
    spec, bitstream = _encode_byte_payload(raw)
    mat_base, func = _build_base_matrix(spec)
    return _apply_data_and_mask(spec, mat_base, func, bitstream)


def generate_qr_svg(text, scale=6, border=4, dark_color="#0f172a", light_color="#ffffff"):
    """Generate a crisp SVG QR code using segno if installed, else the pure-Python encoder."""
    payload = str(text or "")
    try:
        import segno  # type: ignore

        qr = segno.make(payload, error="m")
        buf = io.BytesIO()
        qr.save(buf, kind="svg", scale=scale, border=border, dark=dark_color, light=light_color)
        return buf.getvalue().decode("utf-8")
    except ImportError:
        pass

    matrix = build_qr_matrix(payload)
    n = len(matrix)
    dim = (n + 2 * border) * scale
    paths = []
    for r, row in enumerate(matrix):
        c = 0
        while c < n:
            if row[c]:
                start = c
                while c < n and row[c]:
                    c += 1
                x = (start + border) * scale
                y = (r + border) * scale
                w = (c - start) * scale
                paths.append(f'<rect x="{x}" y="{y}" width="{w}" height="{scale}"/>')
            else:
                c += 1
    safe_dark = escape(dark_color, quote=True)
    safe_light = escape(light_color, quote=True)
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {dim} {dim}" '
        f'width="{dim}" height="{dim}" role="img" aria-label="HSE QR Code">'
        f'<rect width="100%" height="100%" fill="{safe_light}"/>'
        f'<g fill="{safe_dark}">{"".join(paths)}</g>'
        f"</svg>"
    )
