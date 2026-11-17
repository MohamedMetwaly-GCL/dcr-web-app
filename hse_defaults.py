"""Defaults for new HSE projects. See docs/hse/HSE_DECISION_LOG.md for provenance."""

HSE_REGISTERS = (
    ("PTW", "Permit to Work", "DOC_CATEGORIES.permits"),
    ("OBS", "Safety Observations", "DOC_CATEGORIES.OBS_KEYS"),
    ("INC", "Incidents and First Aid", "DOC_CATEGORIES.first_aid / injuries"),
    ("SVP", "Safety Violations and Penalties", "DOC_CATEGORIES.exclusion / financial_violation"),
    ("TBT", "Toolbox Talks", "DOC_CATEGORIES.toolbox / Tool Box Talk TBT.xlsx"),
    ("FEI", "Fire and Equipment Inspections", "TYPES.extinguisher / FORAM inspection forms"),
    ("JSA", "Job Safety Analysis and Risk Assessments", "risk_assessments"),
    ("CEV", "Contractor Evaluations", "contractor_eval_form"),
    ("EVAC", "Evacuation Drills", "evacuation_drills"),
    ("MHR", "Man-hours", "man_hours"),
    ("WKR", "Worker Register", "TYPES.person"),
)

# Ordered legacy category keys and the labels shown in the DCR subtype dropdown.
HSE_SUBTYPES = {
    "PTW": {
        "hot_permit": "Hot Work Permit",
        "cold_permit": "Cold Work Permit",
        "height_permit": "Work at Height Permit",
        "confined_permit": "Confined Space Permit",
        "excavation_permit": "Excavation Permit",
        "electrical_permit": "Electrical / Isolation Permit",
        "lifting_permit": "Lifting Operations Permit",
        "radiography_permit": "Radiography Permit",
        "maintenance": "Maintenance Work",
    },
    "OBS": {
        "unsafe_act": "Unsafe Act",
        "unsafe_condition": "Unsafe Condition",
        "near_miss": "Near Miss",
    },
    "INC": {
        "first_aid": "First Aid Report",
        "injuries": "Injury Register",
    },
    "SVP": {
        "exclusion": "Exclusion Report",
        "financial_violation": "Financial Violation Report",
        "HSE Violation Warning.docx": "HSE Violation Warning",
    },
    "FEI": {
        "extinguisher": "Fire Extinguisher Inspection",
        "hand_tools_inspection": "Hand Tools Inspection",
        "electrical_panel_inspection": "Electrical Panel Inspection",
        "safety_harness_inspection": "Safety Harness Inspection",
        "welding_machine_inspection": "Welding Machine Inspection",
    },
}

HSE_LISTS = {
    "hse_doc_status": ("Open", "Closed"),
    "hse_injury_class": ("FAC", "MTC", "LTI", "Near Miss", "Fatality"),
    "hse_body_part": (
        "Head / Face",
        "Neck",
        "Chest",
        "Abdomen",
        "Upper Back",
        "Lower Back",
        "Left Shoulder",
        "Right Shoulder",
        "Left Arm",
        "Right Arm",
        "Left Hand",
        "Right Hand",
        "Pelvis / Hip",
        "Left Thigh",
        "Right Thigh",
        "Left Knee",
        "Right Knee",
        "Left Leg",
        "Right Leg",
        "Left Foot",
        "Right Foot",
    ),
    "hse_extinguisher_type": (
        "CO2",
        "Dry Powder",
        "Foam",
        "Water",
        "Wet Chemical",
        "Fire Hose Reel",
    ),
    "hse_extinguisher_status": (
        "Valid",
        "Needs Maintenance",
        "Expired",
        "Under Refill",
        "Out of Service",
    ),
    "hse_jsa_status": ("Approved", "Under Review", "Draft", "Revised"),
    "hse_evac_scenario": (
        "Fire Emergency Drill",
        "Gas / Chemical Leak",
        "Structural Collapse / Earthquake",
        "General Site Evacuation",
        "Power Failure / Trapped Personnel",
    ),
    "hse_evac_grade": ("Excellent", "Very Good", "Satisfactory", "Needs Retraining"),
    "hse_alarm_check": ("Yes", "Partial", "No"),
    "hse_yes_no": ("Yes", "No"),
    "hse_worker_affiliation": ("Contractor", "Company Staff"),
    "hse_worker_role": (
        "Staff",
        "Worker",
        "Storekeeper",
        "Safety Officer",
        "Supervisor",
        "Engineer",
        "Visitor",
    ),
}
HSE_LISTS.update({f"hse_{code.lower()}_subtype": tuple(values.values()) for code, values in HSE_SUBTYPES.items()})


# Column key, label, DCR type, optional project-local dropdown name.
DOC_NO = ("docNo", "Register No.", "docno", None)
TITLE = ("title", "Subject", "text", None)
LOCATION = ("location", "Location", "text", None)
RESPONSIBLE = ("responsible", "Responsible Person", "text", None)
NOTES = ("remarks", "Notes", "long_text", None)
DOC_STATUS = ("status", "Status", "dropdown", "hse_doc_status")
YEAR = ("year", "Year", "number", None)
MONTH = ("month", "Month (1–12)", "number", None)
CORRECTIVE_ACTION = ("correctiveAction", "Corrective Action", "long_text", None)
FILE_LOCATION = ("fileLocation", "Drive Link", "url", None)

HSE_COLUMNS = {
    "PTW": (
        TITLE, ("issuedDate", "Permit Date", "date", None),
        ("validUntil", "Valid Until", "date", None), LOCATION, RESPONSIBLE, DOC_STATUS, NOTES,
    ),
    "OBS": (TITLE, ("issuedDate", "Observation Date", "date", None), LOCATION, RESPONSIBLE, DOC_STATUS, NOTES),
    "INC": (
        TITLE, ("issuedDate", "Incident Date", "date", None),
        ("incidentTime", "Incident Time", "text", None), LOCATION,
        ("injuryClass", "Incident Classification", "dropdown", "hse_injury_class"),
        ("bodyPart", "Injured Body Part", "dropdown", "hse_body_part"),
        ("lostDays", "Lost Days", "number", None),
        ("description", "Incident Description", "long_text", None),
        ("rootCause", "Root Cause", "long_text", None), CORRECTIVE_ACTION, NOTES,
    ),
    "SVP": (TITLE, ("issuedDate", "Report Date", "date", None), LOCATION, RESPONSIBLE, DOC_STATUS, NOTES),
    "TBT": (("title", "Talk Topic", "text", None), ("issuedDate", "Talk Date", "date", None), LOCATION, RESPONSIBLE, NOTES),
    "FEI": (
        TITLE, ("equipmentRef", "Equipment / Extinguisher No.", "text", None),
        ("issuedDate", "Inspection Date", "date", None), LOCATION,
        ("inspector", "Inspector", "text", None),
        ("extinguisherType", "Extinguisher Type", "dropdown", "hse_extinguisher_type"),
        ("capacity", "Extinguisher Capacity", "text", None),
        ("manufactureDate", "Manufacture Date", "date", None),
        ("expiryDate", "Expiry / Next Inspection Date", "date", None),
        ("extinguisherStatus", "Extinguisher Status", "dropdown", "hse_extinguisher_status"),
        ("conditionNotes", "Condition Notes", "long_text", None), NOTES,
    ),
    "JSA": (
        TITLE, ("issuedDate", "Assessment Date", "date", None),
        ("reviewDate", "Review Date", "date", None), LOCATION,
        ("activityName", "Activity", "text", None), ("department", "Department", "text", None),
        ("assessorName", "Assessor", "text", None), ("approverName", "Approver Name", "text", None),
        ("status", "Assessment Status", "dropdown", "hse_jsa_status"), NOTES,
    ),
    "CEV": (
        ("title", "Contractor", "text", None), YEAR, MONTH, ("issuedDate", "Evaluation Date", "date", None),
        ("evaluator", "Evaluator", "text", None), ("positiveNotes", "Strengths", "long_text", None),
        ("negativeNotes", "Observations and Deficiencies", "long_text", None),
        ("actionPlan", "Corrective Action Plan", "long_text", None),
    ),
    "EVAC": (
        TITLE, ("issuedDate", "Drill Date", "date", None),
        ("drillTime", "Drill Time", "text", None), ("scenario", "Scenario", "dropdown", "hse_evac_scenario"),
        LOCATION, ("assemblyPoint", "Assembly Point", "text", None), ("leaderName", "Drill Leader", "text", None),
        ("targetDurationSec", "Target Duration (seconds)", "number", None),
        ("actualDurationSec", "Actual Duration (seconds)", "number", None),
        ("evaluationGrade", "Recorded Evaluation", "dropdown", "hse_evac_grade"),
        ("alarmWorking", "Alarm Working", "dropdown", "hse_alarm_check"),
        ("exitRoutesClear", "Exit Routes Clear", "dropdown", "hse_yes_no"),
        ("assemblyAdequate", "Assembly Area Adequate", "dropdown", "hse_yes_no"),
        CORRECTIVE_ACTION, NOTES,
    ),
    "MHR": (YEAR, MONTH, ("hours", "Worked Hours", "number", None)),
    "WKR": (
        ("title", "Worker Name", "text", None),
        ("affiliation", "Affiliation", "dropdown", "hse_worker_affiliation"),
        ("company", "Company / Contractor", "text", None),
        ("personType", "Person Type", "dropdown", "hse_worker_role"),
        ("jobTitle", "Job Title", "text", None), ("hireDate", "Start Date", "date", None), NOTES,
    ),
}


def hse_columns(code):
    subtype = (
        (("legacySubtype", "Subtype", "dropdown", f"hse_{code.lower()}_subtype"),)
        if code in HSE_SUBTYPES else ()
    )
    return (DOC_NO, *subtype, *HSE_COLUMNS[code], FILE_LOCATION)

