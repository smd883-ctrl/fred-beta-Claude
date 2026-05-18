"""
EHCP Processing Pipeline
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from .classifier import (
    FormatClassifier, FormatClassification, KnowledgeBank,
    ClassifierAuditLog, classification_to_parser_args,
)
from .core import EHCPParser
from .models import EHCPDocument

log = logging.getLogger(__name__)


@dataclass
class PipelineResult:
    document: EHCPDocument
    classification: FormatClassification
    from_knowledge_bank: bool = False
    pipeline_notes: list = field(default_factory=list)

    def summary(self) -> str:
        doc = self.document
        fc = self.classification
        lines = [
            f"File:              {doc.source_filename}",
            f"LA:                {doc.la_name or 'Unknown'}",
            f"Format family:     {fc.format_family}",
            f"Classifier source: {'knowledge bank' if self.from_knowledge_bank else 'Claude API'}",
            f"",
            f"Sections found:    {', '.join(doc.sections_found) or 'none'}",
            f"Sections missing:  {', '.join(doc.sections_missing) or 'none'}",
            f"Section F split:   {fc.section_f_split}",
            f"",
            f"Overall confidence:{doc.overall_confidence:.2f} ({doc.overall_confidence_level.value})",
            f"Scanned:           {doc.is_scanned}",
            f"Two-column:        {doc.is_two_column}",
            f"Has contents page: {doc.has_contents_page}",
            f"",
            f"Warnings:          {len(doc.warnings)}",
            f"Errors:            {len(doc.parse_errors)}",
        ]
        if doc.warnings:
            lines.append("")
            lines.append("Warnings:")
            for w in doc.warnings:
                lines.append(f"  [{w.severity.upper()}] {w.code}: {w.message}")
        if doc.parse_errors:
            lines.append("")
            lines.append("Errors:")
            for e in doc.parse_errors:
                lines.append(f"  {e}")
        return "\n".join(lines)


class EHCPPipeline:

    def __init__(
        self,
        api_key=None,
        knowledge_bank_path="ehcp_format_knowledge_bank.json",
        audit_log_path="classifier_audit.jsonl",
        model="claude-sonnet-4-20250514",
        debug=False,
    ):
        self.knowledge_bank = KnowledgeBank(knowledge_bank_path)
        self.classifier = FormatClassifier(
            knowledge_bank=self.knowledge_bank,
            api_key=api_key,
            model=model,
        )
        self.audit_log = ClassifierAuditLog(audit_log_path)
        self.debug = debug

    def process(self, path, la_hint=None) -> PipelineResult:
        path = Path(path)
        notes = []

        page_texts = self._extract_page_texts(path, notes)

        classification = self.classifier.classify(
            page_texts=page_texts,
            la_hint=la_hint,
        )
        from_kb = not classification.classifier_used and not classification.fallback_used

        self.audit_log.record(
            source_filename=path.name,
            la_hint=la_hint,
            classification=classification,
            from_knowledge_bank=from_kb,
        )

        if from_kb:
            notes.append(f"Format from knowledge bank: {classification.format_family}")
        elif classification.classifier_used:
            notes.append(f"Format classified by Claude: {classification.format_family}")
        else:
            notes.append(f"Classification fallback: {classification.format_family}")

        parser_args = classification_to_parser_args(classification)
        parser = EHCPParser(
            preferred_families=parser_args["preferred_families"],
            debug=self.debug,
        )

        la_for_parser = la_hint or classification.la_name_detected
        document = parser.parse(path, la_hint=la_for_parser)

        if classification.la_name_detected and not document.la_name:
            document.la_name = classification.la_name_detected

        if classification.is_scanned and not document.is_scanned:
            document.is_scanned = True
            document.add_warning(
                "SCANNED_PDF_CLASSIFIER",
                "Classifier detected scanned PDF",
                severity="error"
            )

        if classification.classifier_used and classification.la_name_detected:
            notes.append(
                f"LA '{classification.la_name_detected}' added to knowledge bank"
            )

        return PipelineResult(
            document=document,
            classification=classification,
            from_knowledge_bank=from_kb,
            pipeline_notes=notes,
        )

    def _extract_page_texts(self, path, notes):
        suffix = path.suffix.lower()
        texts = []

        if suffix == ".pdf":
            texts = self._pdf_page_texts(path, notes)
        elif suffix in (".docx", ".doc"):
            texts = self._docx_page_texts(path, notes)
        elif suffix in (".txt", ".text"):
            text = path.read_text(encoding="utf-8", errors="replace")
            chunks = text.split("\f") if "\f" in text else [text[:3000]]
            texts = chunks[:3]
        else:
            notes.append(f"Unsupported type for classifier: {suffix}")

        return texts[:3]

    def _pdf_page_texts(self, path, notes):
        texts = []
        try:
            import fitz
            pdf = fitz.open(str(path))
            for i, page in enumerate(pdf):
                if i >= 3:
                    break
                texts.append(page.get_text("text") or "")
            pdf.close()
            return texts
        except ImportError:
            pass
        except Exception as exc:
            notes.append(f"PyMuPDF classifier failed: {exc}")

        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages[:3]:
                    texts.append(page.extract_text() or "")
            return texts
        except ImportError:
            pass
        except Exception as exc:
            notes.append(f"pdfplumber classifier failed: {exc}")

        return []

    def _docx_page_texts(self, path, notes):
        try:
            from docx import Document
            doc = Document(str(path))
            paras = [p.text for p in doc.paragraphs[:60]]
            return ["\n".join(paras)]
        except ImportError:
            notes.append("python-docx not installed")
        except Exception as exc:
            notes.append(f"DOCX classifier failed: {exc}")
        return []

    def knowledge_bank_report(self) -> str:
        return self.knowledge_bank.coverage_report()


def process_ehcp(path, la_hint=None, api_key=None,
                 knowledge_bank_path="ehcp_format_knowledge_bank.json") -> PipelineResult:
    pipeline = EHCPPipeline(api_key=api_key, knowledge_bank_path=knowledge_bank_path)
    return pipeline.process(path, la_hint=la_hint)
