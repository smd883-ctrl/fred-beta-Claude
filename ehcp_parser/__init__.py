"""
EHCP Parser — Sandbox Package
"""

from .core import EHCPParser, parse_ehcp
from .models import (
    EHCPDocument, EHCPSection, ProvisionTable, ProvisionRow,
    ExtractionMethod, ConfidenceLevel, SectionStatus, DocumentFamily,
    ParseWarning,
)
from .section_registry import (
    STATUTORY_SECTIONS, PROVISION_SECTIONS, HIGH_RISK_SECTIONS,
    match_section_heading, looks_like_contents_page, LA_FORMAT_MAP,
)
from .classifier import (
    FormatClassifier, FormatClassification,
    KnowledgeBank, ClassifierAuditLog,
    classification_to_parser_args,
)
from .pipeline import EHCPPipeline, PipelineResult, process_ehcp

__version__ = "0.2.0"
