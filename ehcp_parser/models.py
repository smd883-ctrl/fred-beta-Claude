"""
EHCP Parser — Data Models
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Optional
from enum import Enum
import json


class ExtractionMethod(str, Enum):
    PDFPLUMBER   = "pdfplumber"
    PYMUPDF      = "pymupdf"
    DOCX         = "docx"
    PLAINTEXT    = "plaintext"
    OCR_FALLBACK = "ocr_fallback"
    UNKNOWN      = "unknown"


class ConfidenceLevel(str, Enum):
    HIGH   = "high"
    MEDIUM = "medium"
    LOW    = "low"
    FAILED = "failed"


class SectionStatus(str, Enum):
    FOUND         = "found"
    NOT_FOUND     = "not_found"
    AMBIGUOUS     = "ambiguous"
    CONTENTS_ONLY = "contents_only"
    DUPLICATE     = "duplicate"
    EMPTY         = "empty"


class DocumentFamily(str, Enum):
    STANDARD_LETTERED = "standard_lettered"
    ALPHA_DASH        = "alpha_dash"
    NUMERIC_PART      = "numeric_part"
    BOLD_HEADING_ONLY = "bold_heading_only"
    TWO_COLUMN_TABLE  = "two_column_table"
    HYBRID            = "hybrid"
    SCANNED_IMAGE     = "scanned_image"
    UNKNOWN           = "unknown"


@dataclass
class ProvisionRow:
    raw_cells: list[str]         = field(default_factory=list)
    provision_text: Optional[str] = None
    frequency: Optional[str]     = None
    who_delivers: Optional[str]  = None
    setting: Optional[str]       = None
    quantified: bool             = False
    confidence: float            = 0.0

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class ProvisionTable:
    page_number: Optional[int]          = None
    headers: list[str]                  = field(default_factory=list)
    rows: list[ProvisionRow]            = field(default_factory=list)
    column_count: int                   = 0
    extraction_method: ExtractionMethod = ExtractionMethod.UNKNOWN
    confidence: float                   = 0.0
    notes: list[str]                    = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class EHCPSection:
    letter: Optional[str]             = None
    label: Optional[str]              = None
    title: Optional[str]              = None
    body_text: Optional[str]          = None
    page_start: Optional[int]         = None
    page_end: Optional[int]           = None
    tables: list[ProvisionTable]      = field(default_factory=list)
    status: SectionStatus             = SectionStatus.NOT_FOUND
    confidence: float                 = 0.0
    confidence_level: ConfidenceLevel = ConfidenceLevel.FAILED
    extraction_notes: list[str]       = field(default_factory=list)
    from_contents_page: bool          = False

    def word_count(self) -> int:
        if not self.body_text:
            return 0
        return len(self.body_text.split())

    def has_quantified_provision(self) -> bool:
        if not self.body_text:
            return False
        import re
        pattern = r'\b\d+\s*(hour|minute|session|week|term|per|x\s*per|times)\b'
        return bool(re.search(pattern, self.body_text, re.IGNORECASE))

    def to_dict(self) -> dict:
        d = asdict(self)
        d["status"] = self.status.value
        d["confidence_level"] = self.confidence_level.value
        d["word_count"] = self.word_count()
        d["has_quantified_provision"] = self.has_quantified_provision()
        return d


@dataclass
class ParseWarning:
    code: str
    message: str
    section: Optional[str] = None
    page: Optional[int]    = None
    severity: str          = "warning"

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class EHCPDocument:
    source_path: Optional[str]             = None
    source_filename: Optional[str]         = None
    file_type: Optional[str]              = None
    extraction_method: ExtractionMethod    = ExtractionMethod.UNKNOWN
    page_count: Optional[int]             = None
    child_name: Optional[str]             = None
    child_dob: Optional[str]             = None
    la_name: Optional[str]              = None
    document_date: Optional[str]        = None
    document_family: DocumentFamily      = DocumentFamily.UNKNOWN
    has_contents_page: bool             = False
    contents_page_numbers: list[int]     = field(default_factory=list)
    is_scanned: bool                    = False
    is_two_column: bool                 = False
    sections: dict                       = field(default_factory=dict)
    overall_confidence: float           = 0.0
    overall_confidence_level: ConfidenceLevel = ConfidenceLevel.FAILED
    warnings: list[ParseWarning]        = field(default_factory=list)
    parse_errors: list[str]             = field(default_factory=list)
    sections_found: list[str]           = field(default_factory=list)
    sections_missing: list[str]         = field(default_factory=list)
    sections_ambiguous: list[str]       = field(default_factory=list)
    raw_text_sample: Optional[str]      = None

    def get_section(self, letter: str) -> list:
        return self.sections.get(letter.upper(), [])

    def get_section_f_blocks(self) -> list:
        return self.get_section("F")

    def add_warning(self, code: str, message: str,
                    section: Optional[str] = None,
                    page: Optional[int] = None,
                    severity: str = "warning") -> None:
        self.warnings.append(ParseWarning(
            code=code, message=message,
            section=section, page=page, severity=severity
        ))

    def to_dict(self) -> dict:
        return {
            "source_filename": self.source_filename,
            "file_type": self.file_type,
            "extraction_method": self.extraction_method.value,
            "page_count": self.page_count,
            "la_name": self.la_name,
            "document_family": self.document_family.value,
            "has_contents_page": self.has_contents_page,
            "is_scanned": self.is_scanned,
            "is_two_column": self.is_two_column,
            "overall_confidence": round(self.overall_confidence, 3),
            "overall_confidence_level": self.overall_confidence_level.value,
            "sections_found": self.sections_found,
            "sections_missing": self.sections_missing,
            "sections_ambiguous": self.sections_ambiguous,
            "sections": {
                letter: [s.to_dict() for s in blocks]
                for letter, blocks in self.sections.items()
            },
            "warnings": [w.to_dict() for w in self.warnings],
            "parse_errors": self.parse_errors,
        }

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), indent=indent, ensure_ascii=False)


def confidence_level_from_score(score: float) -> ConfidenceLevel:
    if score >= 0.80:
        return ConfidenceLevel.HIGH
    if score >= 0.50:
        return ConfidenceLevel.MEDIUM
    if score >= 0.25:
        return ConfidenceLevel.LOW
    return ConfidenceLevel.FAILED
