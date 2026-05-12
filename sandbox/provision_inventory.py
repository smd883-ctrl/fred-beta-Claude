import re

PROVISION_PATTERNS = {
    "Communication & Interaction": [
        "visual timetable",
        "social stories",
        "speech and language therapy",
        "SALT",
        "communication support",
    ],

    "Sensory & Physical": [
        "sensory breaks",
        "movement breaks",
        "ear defenders",
        "wobble cushion",
        "OT",
        "occupational therapy",
    ],

    "SEMH / Regulation": [
        "emotional coaching",
        "Zones of Regulation",
        "safe space",
        "regulation support",
        "transition plan",
    ],

    "Cognition & Learning": [
        "chunked instructions",
        "1:1 support",
        "writing support",
        "literacy intervention",
    ],
}


def build_provision_inventory(full_text):
    inventory = {}

    lower_text = full_text.lower()

    for category, provisions in PROVISION_PATTERNS.items():

        found = []

        for item in provisions:
            if item.lower() in lower_text:
                found.append(item)

        if found:
            inventory[category] = found

    return inventory
