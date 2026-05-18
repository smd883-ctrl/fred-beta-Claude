"""
EHCP Parser — Section Registry
"""

import re
from typing import Optional


STATUTORY_SECTIONS = ["A", "B", "C", "D", "E", "F", "G", "H1", "H2", "I", "J", "K"]
PROVISION_SECTIONS = ["F", "G", "H1", "H2"]
HIGH_RISK_SECTIONS = ["B", "E", "F"]

FAMILY_PATTERNS = {
    "standard_lettered": [
        (r"^[Ss]ection\s+(F)\s*[-–—:]\s*(Cognition|Communication|Physical|Sensory|Social|SEMH|Independent|Other)?", "F"),
        (r"^[Ss]ection\s+(F[12])\s*[-–—:]", None),
        (r"^[Ss]ection\s+(F)\s*\(([ivxIVX]+)\)\s*[-–—:]", "F"),
        (r"^[Ss]ection\s+([A-K])\s*[-–—:]\s*(.+)?$", None),
        (r"^[Ss]ection\s+([A-K])\s+[-–—]\s*(.+)?$", None),
        (r"^[Ss]ection\s+(H[12])\s*[-–—:]", None),
        (r"^SECTION\s+([A-K][12]?)\b", None),
    ],
    "alpha_dash": [
        (r"^([A-K][12])\s*[-–—:]\s+(.{8,80})$", None),
        (r"^([A-K])\s+[-–—]\s+(.{8,80})$", None),
        (r"^([A-K])\s*:\s+(.{8,80})$", None),
    ],
    "numeric_part": [
        (r"^[Pp]art\s+(\d{1,2})\b", None),
        (r"^[Pp]art\s+(One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten)\b", None),
    ],
    "bold_heading_only": [
        (r"^(Views,?\s+Interests|Aspirations)\b", "A"),
        (r"^(Special Educational Needs)\b", "B"),
        (r"^(Health Needs?|Health and (?:Wellbeing|Care) Needs?)\b", "C"),
        (r"^(Social Care Needs?)\b", "D"),
        (r"^(Educational Outcomes?|Outcomes?)\b", "E"),
        (r"^(Special Educational Provision)\b", "F"),
        (r"^(Health (?:Care )?Provision)\b", "G"),
        (r"^(Social Care Provision)\b", "H1"),
        (r"^(Other Social Care Provision|Any Other Social)\b", "H2"),
        (r"^(Educational Placement|Placement|School(?:s?)|Setting)\b", "I"),
        (r"^(Personal Budget|Direct Payments?)\b", "J"),
        (r"^(Appendices|Appendix)\b", "K"),
    ],
    "numeric_dot_title": [
        (r"^(\d{1,2})\.\s+(Views,?\s+Interests|Aspirations)", "A"),
        (r"^(\d{1,2})\.\s+(Special Educational Needs)\b", "B"),
        (r"^(\d{1,2})\.\s+(Health Needs?)", "C"),
        (r"^(\d{1,2})\.\s+(Social Care Needs?)", "D"),
        (r"^(\d{1,2})\.\s+(Outcomes?|Educational Outcomes?)", "E"),
        (r"^(\d{1,2})\.\s+(Special Educational Provision)", "F"),
        (r"^(\d{1,2})\.\s+(Health (?:Care )?Provision)", "G"),
        (r"^(\d{1,2})\.\s+(Social Care Provision)\b", "H1"),
        (r"^(\d{1,2})\.\s+(Other Social|Any Other Social)", "H2"),
        (r"^(\d{1,2})\.\s+(Placement|School)", "I"),
        (r"^(\d{1,2})\.\s+(Personal Budget)", "J"),
        (r"^(\d{1,2})\.\s+(Appendi)", "K"),
    ],
}

_COMPILED: dict = {}


def _get_compiled(family: str):
    if family not in _COMPILED:
        _COMPILED[family] = [
            (re.compile(pat, re.IGNORECASE | re.MULTILINE), letter)
            for pat, letter in FAMILY_PATTERNS.get(family, [])
        ]
    return _COMPILED[family]


PART_NUMBER_TO_LETTER: dict = {
    "1": "A", "one": "A", "2": "B", "two": "B",
    "3": "C", "three": "C", "4": "D", "four": "D",
    "5": "E", "five": "E", "6": "F", "six": "F",
    "7": "G", "seven": "G", "8": "H", "eight": "H",
    "9": "I", "nine": "I", "10": "J", "ten": "J",
}

LA_FORMAT_MAP: dict = {
    "warwickshire": "standard_lettered",
    "kent": "standard_lettered",
    "surrey": "standard_lettered",
    "hampshire": "standard_lettered",
    "hertfordshire": "standard_lettered",
    "essex": "standard_lettered",
    "suffolk": "bold_heading_only",
    "norfolk": "standard_lettered",
    "birmingham": "standard_lettered",
    "manchester": "standard_lettered",
    "leeds": "standard_lettered",
    "bristol": "standard_lettered",
    "hackney": "bold_heading_only",
    "greenwich": "alpha_dash",
    "nottinghamshire": "alpha_dash",
    "cornwall": "alpha_dash",
}


def get_family_for_la(la_name: str) -> Optional[str]:
    key = la_name.lower().strip()
    for known, family in LA_FORMAT_MAP.items():
        if known in key or key in known:
            return family
    return None


_TOC_ENTRY_SUFFIX = re.compile(r'\s+\d{1,3}\s*$')

CONTENTS_PAGE_SIGNALS = [
    r"^\s*[Cc]ontents?\s*$",
    r"^\s*[Tt]able\s+of\s+[Cc]ontents?\s*$",
    r"^\s*[Ii]ndex\s*$",
]

CONTENTS_ENTRY_PATTERN = re.compile(
    r"([Ss]ection\s+[A-K]|[Pp]art\s+\d)\s*[-–—:.]*\s*\d+\s*$",
    re.IGNORECASE
)

NEED_AREA_GROUPS = [
    {
        "key": "communication_interaction",
        "label": "Communication and Interaction",
        "pattern": re.compile(r'Communication\s+and\s+Interaction', re.IGNORECASE),
        "maps_to_sections": ["B", "E", "F"],
    },
    {
        "key": "semh",
        "label": "Social, Emotional and Mental Health",
        "pattern": re.compile(r'Social,?\s+Emotional\s+and\s+Mental\s+Health', re.IGNORECASE),
        "maps_to_sections": ["B", "E", "F"],
    },
    {
        "key": "cognition_learning",
        "label": "Cognition and Learning",
        "pattern": re.compile(r'Cognition\s+and\s+Learning', re.IGNORECASE),
        "maps_to_sections": ["B", "E", "F"],
    },
    {
        "key": "physical_sensory",
        "label": "Physical and/or Sensory",
        "pattern": re.compile(r'Physical\s+(?:and/or|and|or)\s+Sensory', re.IGNORECASE),
        "maps_to_sections": ["B", "E", "F"],
    },
    {
        "key": "health_needs",
        "label": "Health Needs",
        "pattern": re.compile(r'Health\s+Needs?', re.IGNORECASE),
        "maps_to_sections": ["C", "E", "G"],
    },
    {
        "key": "social_care",
        "label": "Social Care",
        "pattern": re.compile(r'^Social\s+Care\s*$', re.IGNORECASE),
        "maps_to_sections": ["D", "E", "H1", "H2"],
    },
]


def detect_need_area_group(line: str) -> Optional[dict]:
    line = line.strip()
    if not line or len(line) > 80:
        return None
    for group in NEED_AREA_GROUPS:
        if group["pattern"].search(line):
            return group
    return None


def is_need_area_grouped_format(page_texts: list) -> bool:
    found_groups = set()
    combined = "\n".join(page_texts[:10])
    for group in NEED_AREA_GROUPS:
        if group["pattern"].search(combined):
            found_groups.add(group["key"])
    return len(found_groups) >= 2


def looks_like_contents_page(page_lines: list, threshold: int = 3) -> bool:
    has_contents_heading = any(
        re.match(sig, line)
        for line in page_lines
        for sig in CONTENTS_PAGE_SIGNALS
    )
    entry_count = sum(
        1 for line in page_lines
        if CONTENTS_ENTRY_PATTERN.search(line)
    )
    return has_contents_heading or (entry_count >= threshold)


def is_contents_entry(line: str) -> bool:
    return bool(CONTENTS_ENTRY_PATTERN.search(line))


def looks_like_two_column(page_text: str) -> bool:
    lines = [l for l in page_text.splitlines() if l.strip()]
    if not lines:
        return False
    mid_gap_count = sum(1 for l in lines if re.search(r'\S {8,}\S', l))
    return mid_gap_count >= max(3, len(lines) * 0.25)


PROVISION_TABLE_HEADER_SIGNALS = [
    r"\bprovision\b", r"\bwho\s+(will\s+)?deliver",
    r"\bhow\s+often\b", r"\bfrequency\b",
    r"\bresponsib", r"\bquantif", r"\bsupport\b",
]

_PROVISION_HEADER_RE = [
    re.compile(p, re.IGNORECASE) for p in PROVISION_TABLE_HEADER_SIGNALS
]


def score_as_provision_table_header(header_cells: list) -> float:
    if not header_cells:
        return 0.0
    combined = " ".join(header_cells).lower()
    hits = sum(1 for p in _PROVISION_HEADER_RE if p.search(combined))
    return min(hits / 3.0, 1.0)


def match_section_heading(line: str, families: Optional[list] = None):
    if families is None:
        families = list(FAMILY_PATTERNS.keys())

    line = line.strip()
    if not line or len(line) > 200:
        return None

    if _TOC_ENTRY_SUFFIX.search(line):
        return None

    confidence_map = {
        "standard_lettered": 0.90,
        "alpha_dash": 0.75,
        "numeric_dot_title": 0.70,
        "bold_heading_only": 0.60,
        "numeric_part": 0.45,
    }

    for family in families:
        compiled = _get_compiled(family)
        for pattern, fixed_letter in compiled:
            m = pattern.match(line)
            if m:
                if fixed_letter is not None:
                    letter = fixed_letter.upper()
                else:
                    raw = m.group(1)
                    if family == "numeric_part":
                        letter = PART_NUMBER_TO_LETTER.get(raw.lower(), raw.upper())
                    else:
                        letter = raw.upper()

                title = None
                try:
                    title = m.group(2).strip() if m.group(2) else None
                except IndexError:
                    title = None

                base_conf = confidence_map.get(family, 0.50)
                if re.search(r'\bsection\b', line, re.IGNORECASE):
                    base_conf = min(base_conf + 0.08, 1.0)

                return (letter, title, round(base_conf, 2))

    return None
