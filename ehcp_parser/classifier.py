"""
EHCP Format Classifier
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)


@dataclass
class FormatClassification:
    format_family: str            = "standard_lettered"
    section_f_split: bool         = False
    section_f_labels: list        = field(default_factory=list)
    has_contents_page: bool       = False
    is_two_column: bool           = False
    is_scanned: bool              = False
    heading_examples: list        = field(default_factory=list)
    la_name_detected: Optional[str] = None
    la_name_source: str           = "not_detected"
    classifier_used: bool         = False
    classifier_confidence: str    = "low"
    classifier_model: str         = ""
    pages_sampled: int            = 0
    fallback_used: bool           = False
    raw_response: Optional[str]   = None

    def to_dict(self) -> dict:
        return asdict(self)

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), indent=indent)


SYSTEM_PROMPT = """You are a document format analyser for UK Education, Health and Care Plans (EHCPs).

Your ONLY job is to identify the structural formatting of the document.

Return a JSON object and nothing else. No explanation, no preamble, no markdown fences.

{
  "format_family": string,
  "section_f_split": boolean,
  "section_f_labels": array of strings,
  "has_contents_page": boolean,
  "is_two_column": boolean,
  "is_scanned": boolean,
  "heading_examples": array of strings,
  "la_name_detected": string or null,
  "classifier_confidence": string
}

format_family must be exactly one of:
  "standard_lettered" - headings like "Section A:", "Section F -"
  "alpha_dash" - headings like "A -", "F:" without the word Section
  "numeric_part" - headings like "Part 1", "Part Six"
  "bold_heading_only" - headings are title text only e.g. "Special Educational Provision"
  "numeric_dot_title" - headings like "6. Special Educational Provision"
  "hybrid" - mix of above
  "unknown" - cannot determine

classifier_confidence: "high", "medium", or "low"

IMPORTANT: Return JSON only. Do not read provision content. Do not read child names or personal details."""

USER_PROMPT_TEMPLATE = """Here are the first {page_count} page(s) of an EHCP.
Identify the format only. Return JSON only.

--- DOCUMENT TEXT START ---
{text}
--- DOCUMENT TEXT END ---"""


class KnowledgeBank:

    def __init__(self, path="ehcp_format_knowledge_bank.json"):
        self.path = Path(path)
        self._data = {}
        self._load()

    def _load(self):
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    self._data = json.load(f)
            except Exception as exc:
                log.warning("Knowledge bank load failed: %s", exc)
                self._data = {}

    def _save(self):
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self._data, f, indent=2, ensure_ascii=False)
        except Exception as exc:
            log.warning("Knowledge bank save failed: %s", exc)

    def lookup(self, la_name: str) -> Optional[dict]:
        key = la_name.lower().strip()
        for stored_key, value in self._data.items():
            if stored_key in key or key in stored_key:
                return value
        return None

    def record(self, la_name: str, classification: FormatClassification) -> None:
        key = la_name.lower().strip()
        entry = {
            "la_name": la_name,
            "format_family": classification.format_family,
            "section_f_split": classification.section_f_split,
            "section_f_labels": classification.section_f_labels,
            "has_contents_page": classification.has_contents_page,
            "is_two_column": classification.is_two_column,
            "heading_examples": classification.heading_examples,
            "classifier_confidence": classification.classifier_confidence,
            "classifier_used": classification.classifier_used,
            "last_seen": time.strftime("%Y-%m-%d"),
            "document_count": self._data.get(key, {}).get("document_count", 0) + 1,
        }
        self._data[key] = entry
        self._save()

    def all_entries(self) -> dict:
        return dict(self._data)

    def coverage_report(self) -> str:
        if not self._data:
            return "Knowledge bank is empty."
        lines = [f"Knowledge bank: {len(self._data)} LA(s) recorded\n"]
        family_counts = {}
        for entry in self._data.values():
            fam = entry.get("format_family", "unknown")
            family_counts[fam] = family_counts.get(fam, 0) + 1
        for fam, count in sorted(family_counts.items(), key=lambda x: -x[1]):
            lines.append(f"  {fam}: {count} LA(s)")
        return "\n".join(lines)


class FormatClassifier:

    MAX_CHARS = 4000

    def __init__(self, knowledge_bank=None, api_key=None,
                 model="claude-sonnet-4-20250514"):
        self.bank = knowledge_bank or KnowledgeBank()
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        self.model = model

    def classify(self, page_texts: list, la_hint=None) -> FormatClassification:
        if la_hint:
            stored = self.bank.lookup(la_hint)
            if stored:
                return FormatClassification(
                    format_family=stored.get("format_family", "standard_lettered"),
                    section_f_split=stored.get("section_f_split", False),
                    section_f_labels=stored.get("section_f_labels", []),
                    has_contents_page=stored.get("has_contents_page", False),
                    is_two_column=stored.get("is_two_column", False),
                    heading_examples=stored.get("heading_examples", []),
                    la_name_detected=la_hint,
                    la_name_source="user_provided",
                    classifier_used=False,
                    classifier_confidence=stored.get("classifier_confidence", "medium"),
                )

        sample_pages = page_texts[:3]
        sample_text = "\n\n--- PAGE BREAK ---\n\n".join(sample_pages)
        sample_text = sample_text[:self.MAX_CHARS]
        pages_sampled = len(sample_pages)

        if not self.api_key:
            return self._default_classification(la_hint=la_hint,
                                                fallback_reason="no_api_key")

        try:
            fc = self._call_claude(sample_text, pages_sampled, la_hint)
        except Exception as exc:
            log.warning("Classifier failed: %s", exc)
            fc = self._default_classification(la_hint=la_hint, fallback_reason=str(exc))

        la_name = la_hint or fc.la_name_detected
        if la_name:
            self.bank.record(la_name, fc)

        return fc

    def _call_claude(self, text, pages_sampled, la_hint):
        import urllib.request

        prompt = USER_PROMPT_TEMPLATE.format(
            page_count=pages_sampled,
            text=text,
        )

        payload = json.dumps({
            "model": self.model,
            "max_tokens": 500,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": prompt}],
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": self.api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST",
        )

        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))

        raw_text = body["content"][0]["text"].strip()
        return self._parse_response(raw_text, pages_sampled, la_hint)

    def _parse_response(self, raw_text, pages_sampled, la_hint):
        clean = re.sub(r"```(?:json)?|```", "", raw_text).strip()
        try:
            data = json.loads(clean)
        except json.JSONDecodeError:
            return self._default_classification(la_hint=la_hint,
                                                fallback_reason="json_parse_failed")

        valid_families = {
            "standard_lettered", "alpha_dash", "numeric_part",
            "bold_heading_only", "numeric_dot_title", "hybrid", "unknown"
        }
        family = data.get("format_family", "unknown")
        if family not in valid_families:
            family = "standard_lettered"

        return FormatClassification(
            format_family=family,
            section_f_split=bool(data.get("section_f_split", False)),
            section_f_labels=list(data.get("section_f_labels", [])),
            has_contents_page=bool(data.get("has_contents_page", False)),
            is_two_column=bool(data.get("is_two_column", False)),
            is_scanned=bool(data.get("is_scanned", False)),
            heading_examples=list(data.get("heading_examples", []))[:6],
            la_name_detected=data.get("la_name_detected") or la_hint,
            la_name_source="document_text" if data.get("la_name_detected")
                           else ("user_provided" if la_hint else "not_detected"),
            classifier_used=True,
            classifier_confidence=data.get("classifier_confidence", "low"),
            classifier_model=self.model,
            pages_sampled=pages_sampled,
            fallback_used=False,
            raw_response=raw_text[:500],
        )

    def _default_classification(self, la_hint=None, fallback_reason="unknown"):
        return FormatClassification(
            format_family="standard_lettered",
            la_name_detected=la_hint,
            la_name_source="user_provided" if la_hint else "not_detected",
            classifier_used=False,
            classifier_confidence="low",
            fallback_used=True,
            raw_response=f"fallback: {fallback_reason}",
        )


def classification_to_parser_args(fc: FormatClassification) -> dict:
    FAMILY_ORDER = {
        "standard_lettered": ["standard_lettered", "alpha_dash",
                               "bold_heading_only", "numeric_dot_title", "numeric_part"],
        "alpha_dash":        ["alpha_dash", "standard_lettered",
                               "bold_heading_only", "numeric_dot_title", "numeric_part"],
        "bold_heading_only": ["bold_heading_only", "standard_lettered",
                               "alpha_dash", "numeric_dot_title", "numeric_part"],
        "numeric_dot_title": ["numeric_dot_title", "standard_lettered",
                               "alpha_dash", "bold_heading_only", "numeric_part"],
        "numeric_part":      ["numeric_part", "standard_lettered",
                               "alpha_dash", "bold_heading_only", "numeric_dot_title"],
        "hybrid":            ["standard_lettered", "alpha_dash",
                               "bold_heading_only", "numeric_dot_title", "numeric_part"],
        "unknown":           ["standard_lettered", "alpha_dash",
                               "bold_heading_only", "numeric_dot_title", "numeric_part"],
    }
    return {
        "preferred_families": FAMILY_ORDER.get(fc.format_family, FAMILY_ORDER["unknown"])
    }


class ClassifierAuditLog:

    def __init__(self, path="classifier_audit.jsonl"):
        self.path = Path(path)

    def record(self, source_filename, la_hint, classification, from_knowledge_bank):
        entry = {
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "source_filename": source_filename,
            "la_hint": la_hint,
            "from_knowledge_bank": from_knowledge_bank,
            "format_family": classification.format_family,
            "classifier_confidence": classification.classifier_confidence,
            "section_f_split": classification.section_f_split,
            "fallback_used": classification.fallback_used,
        }
        try:
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry) + "\n")
        except Exception as exc:
            log.warning("Audit log write failed: %s", exc)
