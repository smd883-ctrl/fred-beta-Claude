import re

PROVISION_PATTERNS = {
    "Communication & Interaction": [
        "visual timetable",
        "social stories",
        "speech and language therapy",
        "SALT",
        "communication support",
        "visual supports",
        "now and next board",
        "communication passport",
        "comic strip conversations",
        "language intervention",
        "speech and language input",
        "social communication support",
        "processing time",
        "reduced language load",
    ],

    "Sensory & Physical": [
        "sensory breaks",
        "movement breaks",
        "ear defenders",
        "wobble cushion",
        "OT",
        "occupational therapy",
        "sensory diet",
        "sensory circuit",
        "movement opportunity",
        "movement opportunities",
        "fidget tool",
        "weighted item",
        "chewelry",
        "quiet workstation",
        "low arousal environment",
        "sensory regulation",
    ],

    "SEMH / Regulation": [
        "emotional coaching",
        "Zones of Regulation",
        "safe space",
        "regulation support",
        "transition plan",
        "emotional regulation",
        "trusted adult",
        "check in",
        "safe adult",
        "time out card",
        "emotion coaching",
        "restorative conversation",
        "co-regulation",
        "regulation strategy",
        "anxiety support",
    ],

    "Cognition & Learning": [
        "chunked instructions",
        "1:1 support",
        "writing support",
        "literacy intervention",
        "task chunking",
        "scaffolded learning",
        "overlearning",
        "pre-teaching",
        "writing frame",
        "exam access arrangements",
        "processing support",
        "instruction breakdown",
        "differentiated resources",
    ],
}


def build_provision_inventory(full_text):
    inventory = {}

    lower_text = full_text.lower()

    for category, provisions in PROVISION_PATTERNS.items():

        found = set()

        for item in provisions:
            if item.lower() in lower_text:
                found.add(item)

        if found:
            inventory[category] = sorted(found)

    return inventory
