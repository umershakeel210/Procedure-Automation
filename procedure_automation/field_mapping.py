DEFAULT_VALUES = {
    "defined_as": "Clinical Package",
    "rate_list": "Rate List-2024",
    "procedure_type": "IPD",
    "length_of_stay": 1,
    "notes": "",
    "active": True,
}

COMPONENT_MAPPING = {
    "Procedure Fee": {
        "component_name": "Procedure Fee",
        "billing_type": "Fixed"
    },
    "Assistant Surgeon Fee": {
        "component_name": "Assistant Surgeon Fee",
        "billing_type": "Fixed"
    },
    "Anaesthetist": {
        "component_name": "Anesthetist Fee",
        "billing_type": "Fixed"
    },
    "Procedure Room Time Charges": {
        "component_name": "Procedure Room Time",
        "billing_type": "Fixed"
    },
    "Stay Charges": {
        "component_name": "Stay Charges",
        "billing_type": "Fixed"
    },
    "Procedure Room Pharmacy": {
        "component_name": "Procedure Room Pharmacy",
        "billing_type": "Fixed"
    },
    "Ward Pharmacy": {
        "component_name": "Ward Pharmacy",
        "billing_type": "Fixed"
    },
    "Laboratory": {
        "component_name": "Laboratory",
        "billing_type": "Fixed"
    },
    "MO/Nursing charges": {
        "component_name": "MO/Nursing charges",
        "billing_type": "Fixed"
    },
}