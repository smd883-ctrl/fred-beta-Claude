import streamlit as st
import fitz  # PyMuPDF
import re
import json
import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="FRED — Families' Rights and Entitlements Directory",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── COLOUR CONSTANTS ─────────────────────────────────────────────────────────

RED    = "#C0392B"
AMBER  = "#D4A017"
GREEN  = "#1E8449"
NAVY   = "#1a2744"
LIGHT  = "#f7f9fc"

GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeA1F9nEdQWkmplbAh973XKq2EsW0bEkhJiw7drhP7BZaPjKQ/viewform"

# ── CUSTOM CSS ───────────────────────────────────────────────────────────────

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500&display=swap');

  html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif;
    background-colour: #f5f5f0;
    colour: #1a1a2e;
  }}

  h1, h2, h3 {{
    font-family: 'DM Serif Display', serif;
    font-weight: 400;
  }}

  /* Hero */
  .hero {{
    background: #2d4a2d;
    colour: white;
    padding: 3rem 2.5rem 2.8rem;
    border-radius: 16px;
    text-align: centre;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
  }}

  .hero::before {{
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 220px; height: 220px;
    border-radius: 50%;
    background: rgba(160,210,130,0.1);
  }}

  .hero h1 {{
    font-family: 'DM Serif Display', serif;
    font-size: 4rem;
    margin: 0;
    colour: #e8f5e0;
    letter-spacing: 0.06em;
    font-weight: 400;
  }}

  .hero .full-name {{
    font-family: 'DM Sans', sans-serif;
    font-size: 0.75rem;
    letter-spacing: 0.22em;
    text-transform: uppercase;
    colour: #9dc98a;
    margin: 0.4rem 0 1rem 0;
    font-weight: 300;
  }}

  .hero .tagline {{
    font-family: 'DM Serif Display', serif;
    font-size: 1.45rem;
    font-weight: 400;
    colour: #e8f5e0;
    margin-bottom: 0.5rem;
    line-height: 1.35;
  }}

  .hero .origin {{
    font-style: italic;
    colour: #7ab870;
    margin-bottom: 1.4rem;
    font-size: 0.93rem;
    font-weight: 300;
  }}

  .hero .sub {{
    colour: #b8d9a8;
    max-width: 540px;
    margin: 0 auto 1.4rem auto;
    font-size: 1rem;
    line-height: 1.7;
    font-weight: 300;
  }}

  .hero .service-line {{
    colour: #7ab870;
    font-size: 0.82rem;
    margin-bottom: 1.8rem;
    letter-spacing: 0.04em;
  }}

  /* Gold CTA button */
  .cta-button {{
    display: inline-block;
    background: #e8c96a;
    colour: #2d3a1a !important;
    font-family: 'DM Sans', sans-serif;
    font-weight: 600;
    font-size: 1rem;
    padding: 0.8rem 2.2rem;
    border-radius: 40px;
    text-decoration: none;
    letter-spacing: 0.02em;
    cursor: pointer;
    border: none;
  }}

  .reassurance {{
    colour: #9dc98a;
    font-size: 0.85rem;
    margin-top: 1rem;
    font-weight: 300;
  }}

  .pricing-hint {{
    colour: #7a8fa8;
    font-size: 0.82rem;
    margin-top: 0.4rem;
  }}

  /* Traffic light badges */
  .badge-red {{
    background: {RED};
    colour: white;
    padding: 0.22rem 0.7rem;
    border-radius: 20px;
    font-weight: 500;
    font-size: 0.78rem;
    display: inline-block;
    margin-bottom: 0.5rem;
    letter-spacing: 0.03em;
  }}
  .badge-amber {{
    background: {AMBER};
    colour: white;
    padding: 0.22rem 0.7rem;
    border-radius: 20px;
    font-weight: 500;
    font-size: 0.78rem;
    display: inline-block;
    margin-bottom: 0.5rem;
    letter-spacing: 0.03em;
  }}
  .badge-green {{
    background: {GREEN};
    colour: white;
    padding: 0.22rem 0.7rem;
    border-radius: 20px;
    font-weight: 500;
    font-size: 0.78rem;
    display: inline-block;
    margin-bottom: 0.5rem;
    letter-spacing: 0.03em;
  }}

  /* Finding cards */
  .finding-red {{
    border-left: 4px solid {RED};
    background: #fdf4f3;
    padding: 1rem 1.2rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 1rem;
  }}
  .finding-amber {{
    border-left: 4px solid {AMBER};
    background: #fdf9f0;
    padding: 1rem 1.2rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 1rem;
  }}
  .finding-green {{
    border-left: 4px solid {GREEN};
    background: #f3faf5;
    padding: 1rem 1.2rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 1rem;
  }}

  /* How it works */
  .hiw-item {{
    display: flex;
    align-items: flex-start;
    gap: 1rem;
    margin-bottom: 0.8rem;
    padding: 0.9rem 1rem;
    background: white;
    border-radius: 10px;
    border: 0.5px solid #dde8dd;
  }}
  .hiw-num {{
    font-family: 'DM Serif Display', serif;
    font-size: 1.3rem;
    colour: #7ab870;
    flex-shrink: 0;
    line-height: 1;
    width: 24px;
  }}

  /* Pricing cards */
  .pricing-card {{
    background: white;
    border: 0.5px solid #d0ddd0;
    border-radius: 12px;
    padding: 1.4rem 1.2rem;
    text-align: centre;
    height: 100%;
  }}
  .pricing-card.featured {{
    border: 1.5px solid #7ab870;
  }}
  .pricing-card .price {{
    font-family: 'DM Serif Display', serif;
    font-size: 2rem;
    font-weight: 400;
    colour: #2d4a2d;
  }}
  .pricing-card .period {{
    font-size: 0.82rem;
    colour: #888;
    font-weight: 300;
  }}
  .pricing-card ul {{
    text-align: left;
    padding-left: 0;
    margin-top: 1rem;
    font-size: 0.88rem;
    line-height: 2;
    list-style: none;
    font-weight: 300;
    colour: #555;
  }}
  .pricing-card ul li::before {{
    content: '·  ';
    colour: #7ab870;
    font-weight: 700;
  }}

  .best-value-tag {{
    background: #2d4a2d;
    colour: #e8f5e0;
    font-size: 0.7rem;
    font-weight: 500;
    padding: 0.2rem 0.7rem;
    border-radius: 20px;
    display: inline-block;
    margin-bottom: 0.5rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }}

  /* Sneak peek */
  .sneak-box {{
    border: 1.5px dashed #7ab870;
    border-radius: 12px;
    padding: 1.5rem;
    background: white;
    margin: 1.5rem 0;
  }}

  /* Document pills */
  .doc-pill {{
    display: inline-block;
    background: #eaf5e0;
    colour: #2d5a2d;
    font-size: 0.78rem;
    padding: 0.28rem 0.8rem;
    border-radius: 20px;
    margin: 0.2rem;
    font-weight: 400;
  }}

  .section-divider {{
    border: none;
    border-top: 0.5px solid #dde8dd;
    margin: 2.5rem 0;
  }}

  .section-label {{
    font-size: 0.7rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    colour: #5a8a5a;
    margin-bottom: 0.6rem;
  }}

  /* Streamlit button override */
  .stButton > button {{
    background: #2d4a2d;
    colour: white;
    border: none;
    border-radius: 40px;
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
    padding: 0.6rem 1.6rem;
    font-size: 0.97rem;
  }}
  .stButton > button:hover {{
    background: #3d5e3d;
    colour: white;
  }}

  /* Primary CTA Streamlit button — gold */
  div[data-testid="stButton"] button[kind="primary"] {{
    background: #e8c96a;
    colour: #2d3a1a;
    border-radius: 40px;
  }}

  footer {{visibility: hidden;}}
  #MainMenu {{visibility: hidden;}}
</style>
""", unsafe_allow_html=True)

# ── EMAILJS ──────────────────────────────────────────────────────────────────


# ── PDF PARSER ────────────────────────────────────────────────────────────────
# Warwickshire format: multiple Section F blocks, two-column tables.
# Strategy: collect ALL text, identify section boundaries, extract each F block.

PROHIBITED_WORDS = [
    "should", "could", "may", "access to", "as needed", "where necessary",
    "as appropriate", "regular", "encouraged", "mindful", "cognisant",
    "holistic approach", "opportunity", "it is expected", "would benefit from",
    "it is recommended", "at their discretion", "where possible",
    "as directed by", "flexible", "flexibility", "responsive", "tailored",
    "embedded",
]

def extract_text_from_pdf(uploaded_file):
    """
    Extract raw text from every page of uploaded PDF.
    Tries text extraction first, then falls back to block extraction.
    Returns extracted text or empty string.
    """
    try:
        data = uploaded_file.read()
        doc  = fitz.open(stream=data, filetype="pdf")
        pages = []
        for page in doc:
            # Try standard text extraction
            text = page.get_text("text")
            if not text.strip():
                # Try block extraction as fallback
                blocks = page.get_text("blocks")
                text = "\n".join(b[4] for b in blocks if isinstance(b[4], str))
            if not text.strip():
                # Try dict extraction as last resort
                d = page.get_text("dict")
                words = []
                for block in d.get("blocks", []):
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            words.append(span.get("text", ""))
                text = " ".join(words)
            pages.append(text)
        result = "\n".join(pages)
        print(f"PDF extracted: {len(result)} chars from {len(doc)} pages")
        return result
    except Exception as e:
        print(f"PDF extraction error: {e}")
        return ""

def extract_text_from_docx(uploaded_file):
    """Extract raw text from uploaded Word document."""
    try:
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""


def extract_text_from_txt(uploaded_file):
    """Extract raw text from plain text file."""
    try:
        return uploaded_file.read().decode("utf-8", errors="ignore")
    except Exception:
        return ""


def extract_text(uploaded_file):
    """Universal text extractor — routes by file type."""
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        text = extract_text_from_pdf(uploaded_file)
        if len(text.strip()) < 50:
            # PDF extracted too little — likely image-based
            return "__PDF_EXTRACTION_FAILED__"
        return text
    elif name.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)
    elif name.endswith(".txt"):
        return extract_text_from_txt(uploaded_file)
    return ""

def find_section_blocks(full_text, section_letter):
    """
    Return list of text blocks for a given section letter (e.g. 'F', 'E', 'B').
    Handles repeated sections (Warwickshire format).
    Skips contents page matches by requiring substantial content (>80 chars).
    """
    # Pattern: "Section X" or just the letter heading used in Warwickshire docs
    pattern = re.compile(
        rf'(?:Section\s+{section_letter}|^{section_letter}\.?\s+[A-Z])',
        re.IGNORECASE | re.MULTILINE
    )
    # Also try simple "Section F" variants
    alt_pattern = re.compile(
        rf'\bSection\s+{section_letter}\b',
        re.IGNORECASE
    )

    positions = []
    for m in alt_pattern.finditer(full_text):
        positions.append(m.start())

    # Deduplicate positions within 50 chars of each other (same heading)
    deduped = []
    for pos in sorted(positions):
        if not deduped or pos - deduped[-1] > 50:
            deduped.append(pos)

    blocks = []
    all_section_starts = sorted(
        [m.start() for m in re.finditer(
            r'\bSection\s+[A-Z]\b', full_text, re.IGNORECASE
        )]
    )

    for i, start in enumerate(deduped):
        # Find where this block ends — next section heading or EOF
        next_starts = [s for s in all_section_starts if s > start + 10]
        end = next_starts[0] if next_starts else len(full_text)
        block = full_text[start:end].strip()
        if len(block) > 80:  # skip contents-page one-liners
            blocks.append(block)

    return blocks


def analyse_section_f(f_blocks):
    """
    Analyse all Section F blocks.
    Returns list of findings dicts with keys:
      tier, title, extract, commentary, delivery_log_required
    """
    findings = []

    if not f_blocks:
        findings.append({
            "tier": "red",
            "title": "Section F not located",
            "extract": "No Section F provision content was identified in this document.",
            "commentary": (
                "Section F is a lawful requirement under the Children and Families Act 2014. "
                "Every EHCP must specify the educational provision in Section F. "
                "If this is a draft document, check whether the section has been completed. "
                "If this is a final EHCP, raise the absence at the earliest opportunity."
            ),
            "delivery_log_required": True,
        })
        return findings

    combined_f = "\n\n".join(f_blocks)

    # ── Check 1: Frequency stated ─────────────────────────────────────────
    freq_pattern = re.compile(r'\b(\d+)\s*(session|hour|minute|time|per week|per term|weekly|daily)\b', re.IGNORECASE)
    has_frequency = bool(freq_pattern.search(combined_f))
    if not has_frequency:
        findings.append({
            "tier": "red",
            "title": "Frequency of provision not specified",
            "extract": _get_short_extract(combined_f, 0, 300),
            "commentary": (
                "Section F must state how often each provision is delivered — "
                "for example, three sessions per week. Without a stated frequency, "
                "the provision cannot be lawfully enforced under the Children and Families Act 2014. "
                "Raise at annual review and request specific numeric frequency for each provision."
            ),
            "delivery_log_required": True,
        })

    # ── Check 2: Duration stated ──────────────────────────────────────────
    dur_pattern = re.compile(r'\b(\d+)\s*(minute|hour|min)\b', re.IGNORECASE)
    has_duration = bool(dur_pattern.search(combined_f))
    if not has_duration:
        findings.append({
            "tier": "red",
            "title": "Duration of sessions not specified",
            "extract": _get_short_extract(combined_f, 0, 300),
            "commentary": (
                "Section F must state the length of each session in minutes or hours. "
                "Without a stated duration, provision is unquantifiable and unenforceable. "
                "Request an amendment to specify session length for every provision."
            ),
            "delivery_log_required": True,
        })

    # ── Check 3: Deliverer role stated ────────────────────────────────────
    role_pattern = re.compile(
        r'\b(SENCO|teaching assistant|TA|therapist|specialist|speech|occupational|'
        r'psychologist|learning support|LSA|teacher|practitioner|coordinator)\b',
        re.IGNORECASE
    )
    has_role = bool(role_pattern.search(combined_f))
    if not has_role:
        findings.append({
            "tier": "red",
            "title": "Deliverer role not specified",
            "extract": _get_short_extract(combined_f, 0, 300),
            "commentary": (
                "Section F must identify who delivers each provision — the role title "
                "and relevant qualification or training level. This is a lawful requirement. "
                "Without a named role, accountability for delivery cannot be established."
            ),
            "delivery_log_required": True,
        })

    # ── Check 4: Prohibited language ─────────────────────────────────────
    found_prohibited = []
    for word in PROHIBITED_WORDS:
        if re.search(rf'\b{re.escape(word)}\b', combined_f, re.IGNORECASE):
            # Find context around the word
            m = re.search(rf'.{{0,60}}\b{re.escape(word)}\b.{{0,60}}', combined_f, re.IGNORECASE)
            ctx = m.group(0).strip() if m else word
            found_prohibited.append((word, ctx))

    if found_prohibited:
        # Group into one red finding
        examples = "; ".join(f'"{w}"' for w, _ in found_prohibited[:5])
        extract  = found_prohibited[0][1] if found_prohibited else ""
        findings.append({
            "tier": "red",
            "title": f"Weak or unenforceable language detected in Section F",
            "extract": f"Language found: {examples}",
            "commentary": (
                "Section F must use 'must' not 'should', 'will' not 'may', and must avoid "
                "vague qualifiers such as 'as appropriate', 'where necessary', or 'flexible'. "
                "Vague language removes enforceability. Each instance should be challenged "
                "and replaced with specific, quantified, unambiguous commitments. "
                "This is a lawful requirement under the Children and Families Act 2014."
            ),
            "delivery_log_required": True,
        })

    # ── Check 5: 'Should' vs 'must' ──────────────────────────────────────
    # Already covered above in prohibited language — skip duplicate.

    # ── Check 6: Delivery log reference ──────────────────────────────────
    log_pattern = re.compile(r'\b(delivery log|record of delivery|session record|log)\b', re.IGNORECASE)
    has_log_ref = bool(log_pattern.search(combined_f))
    if not has_log_ref:
        findings.append({
            "tier": "amber",
            "title": "No delivery log mechanism referenced",
            "extract": "",
            "commentary": (
                "Best practice requires that provision delivery is logged contemporaneously. "
                "A delivery log is the evidence base for the Do stage of the APDR cycle. "
                "Without a delivery log, the Review stage cannot be conducted accurately. "
                "Ask the school what logging system is in place and request access to records."
            ),
            "delivery_log_required": True,
        })

    # ── Check 7: Universal provision substitution ─────────────────────────
    # Note: Warwickshire EHCPs contain a standard disclaimer header:
    # "The provision set out here is in addition to that which all pupils should have access to"
    # This is the LA correctly stating the opposite of substitution — must NOT be flagged.
    # Only flag where universal provision is described AS the specific provision itself.
    universal_pattern = re.compile(
        r'\b(quality first teaching|ordinarily available provision|'
        r'universal provision|whole school approach|available to all pupils)\b',
        re.IGNORECASE
    )
    m = universal_pattern.search(combined_f)
    if m:
        ctx = combined_f[max(0, m.start()-80):m.end()+80].strip()
        # Skip if this is the standard Warwickshire disclaimer header
        disclaimer = re.compile(
            r'provision set out here is in addition to that which all pupils',
            re.IGNORECASE
        )
        if not disclaimer.search(ctx):
            findings.append({
                "tier": "red",
                "title": "Universal provision substituted for specific provision",
                "extract": ctx[:300],
                "commentary": (
                    "Section F must contain provision that is specific to this child, "
                    "above and beyond what is available to all pupils. Universal or "
                    "'quality first' provision does not meet the lawful requirement for "
                    "specified provision in an EHCP. This is the most serious Section F failure. "
                    "Challenge and request replacement with individually specified provision."
                ),
                "delivery_log_required": True,
            })

    # ── Check 8: Recommendation laundering ───────────────────────────────
    launder_pattern = re.compile(
        r'\b(it is recommended|it would be beneficial|it is suggested|'
        r'consideration should be given|it is hoped)\b',
        re.IGNORECASE
    )
    m = launder_pattern.search(combined_f)
    if m:
        ctx = combined_f[max(0, m.start()-80):m.end()+80].strip()
        findings.append({
            "tier": "red",
            "title": "Recommendation language detected — not converted to lawful commitment",
            "extract": ctx[:300],
            "commentary": (
                "Assessment language has been copied into Section F without being converted "
                "into a specified lawful commitment. Section F must contain commitments, "
                "not recommendations. 'It is recommended' has no enforceability. "
                "Each recommendation must be rewritten as a specified, quantified commitment."
            ),
            "delivery_log_required": False,
        })

    # ── Check 9: Dilution clause ──────────────────────────────────────────
    dilution_pattern = re.compile(
        r'\b(subject to|dependent on|contingent on|availability|resources|'
        r'staffing|capacity|where resources allow|budget)\b',
        re.IGNORECASE
    )
    m = dilution_pattern.search(combined_f)
    if m:
        ctx = combined_f[max(0, m.start()-80):m.end()+80].strip()
        findings.append({
            "tier": "red",
            "title": "Dilution clause detected — provision made conditional",
            "extract": ctx[:300],
            "commentary": (
                "Section F provision must not be conditional on school resources, staffing, "
                "or budget. The duty to deliver specified provision is absolute under the "
                "Children and Families Act 2014. Any clause that makes provision conditional "
                "removes its enforceability and must be challenged and removed."
            ),
            "delivery_log_required": False,
        })

    # ── Compliant check ───────────────────────────────────────────────────
    # If freq + duration + role all present and no prohibited language found
    n_red = sum(1 for f in findings if f["tier"] == "red")
    if n_red == 0:
        findings.append({
            "tier": "green",
            "title": "Section F provision appears lawfully specified",
            "extract": "",
            "commentary": (
                "Section F contains stated frequency, duration, and deliverer role. "
                "No prohibited language was detected. Use these compliant entries as "
                "benchmarks when challenging any non-compliant provision in this or "
                "future EHCPs."
            ),
            "delivery_log_required": False,
        })

    return findings


def analyse_section_e(e_blocks):
    """Analyse Section E outcomes."""
    findings = []
    if not e_blocks:
        findings.append({
            "tier": "amber",
            "title": "Section E outcomes not located",
            "extract": "",
            "commentary": (
                "Section E should contain SMART outcomes with a baseline, measurable "
                "progress indicator, and timeframe. If outcomes are absent, raise at "
                "annual review."
            ),
            "delivery_log_required": False,
        })
        return findings

    combined_e = "\n\n".join(e_blocks)

    # Check for baseline
    baseline_pattern = re.compile(r'\b(baseline|currently|starting point|at present)\b', re.IGNORECASE)
    has_baseline = bool(baseline_pattern.search(combined_e))
    if not has_baseline:
        findings.append({
            "tier": "amber",
            "title": "Section E outcomes lack a stated baseline",
            "extract": _get_short_extract(combined_e, 0, 300),
            "commentary": (
                "SMART outcomes require a baseline — where the child is now — so that "
                "progress can be measured. Without a baseline the Review stage of the "
                "APDR cycle cannot be conducted accurately. Request that a baseline "
                "is added to each outcome at the next review."
            ),
            "delivery_log_required": False,
        })

    # Check for timeframe
    time_pattern = re.compile(r'\b(by \w+ 20\d\d|within \d+ (month|week|term|year))\b', re.IGNORECASE)
    has_timeframe = bool(time_pattern.search(combined_e))
    if not has_timeframe:
        findings.append({
            "tier": "amber",
            "title": "Section E outcomes lack a measurable timeframe",
            "extract": _get_short_extract(combined_e, 0, 300),
            "commentary": (
                "Outcomes must include a timeframe — for example, 'by July 2025'. "
                "Without a timeframe, outcomes cannot be objectively reviewed. "
                "Raise at annual review."
            ),
            "delivery_log_required": False,
        })

    if not findings:
        findings.append({
            "tier": "green",
            "title": "Section E outcomes contain baseline and timeframe indicators",
            "extract": "",
            "commentary": (
                "Section E appears to contain measurable outcome language. "
                "Use these as benchmarks."
            ),
            "delivery_log_required": False,
        })

    return findings


def analyse_section_b(b_blocks, f_blocks):
    """
    Cross-reference Section B needs against Section F provision.
    For each need area identified in B, check whether F contains matching provision.
    Produce a 'what good looks like' summary for each gap found.
    """
    findings = []
    if not b_blocks:
        return findings

    combined_b = "\n\n".join(b_blocks)
    combined_f = "\n\n".join(f_blocks) if f_blocks else ""

    # Define need areas with keywords and what good looks like
    need_areas = [
        {
            "name": "Communication and interaction",
            "b_keywords": r'\b(communication|speech|language|SALT|interaction|social communication|autism|ASD|pragmatic)\b',
            "f_keywords": r'\b(speech|language|communication|SALT|social skills|interaction)\b',
            "what_good_looks_like": (
                "Section F should specify: the frequency of speech and language therapy or "
                "targeted communication sessions (e.g. two sessions per week), the duration "
                "of each session in minutes, the role and qualification of the deliverer "
                "(e.g. Speech and Language Therapist or trained TA working to a SALT programmeme), "
                "and the specific strategies to be used. Generalised statements such as "
                "'communication support will be provided' are not sufficient."
            ),
        },
        {
            "name": "Cognition and learning",
            "b_keywords": r'\b(cognition|learning|literacy|numeracy|reading|writing|dyslexia|processing|memory|attention)\b',
            "f_keywords": r'\b(literacy|numeracy|reading|writing|learning support|intervention|programmeme)\b',
            "what_good_looks_like": (
                "Section F should specify: named literacy or learning interventions (not generic "
                "'learning support'), frequency and duration of sessions, the role of the "
                "deliverer and their relevant training, and measurable targets that link "
                "directly to Section E outcomes. 'Access to' a resource is not provision."
            ),
        },
        {
            "name": "Social, emotional and mental health",
            "b_keywords": r'\b(SEMH|emotional|mental health|anxiety|behaviour|wellbeing|regulation|self.esteem|ADHD|attachment)\b',
            "f_keywords": r'\b(emotional|SEMH|regulation|therapeutic|counselling|pastoral|wellbeing|anxiety)\b',
            "what_good_looks_like": (
                "Section F should specify: named therapeutic or emotional support provision "
                "(e.g. weekly one-to-one sessions with a trained SEMH TA or school counsellor), "
                "the frequency and duration, the deliverer's role and relevant training, "
                "and specific strategies referenced from assessment. 'Pastoral support' "
                "without specification does not meet the lawful standard."
            ),
        },
        {
            "name": "Sensory and physical needs",
            "b_keywords": r'\b(sensory|physical|motor|OT|occupational therapy|fine motor|gross motor|proprioception|vestibular|sensory processing)\b',
            "f_keywords": r'\b(OT|occupational|sensory|motor|physical|sensory diet|movement break)\b',
            "what_good_looks_like": (
                "Section F should specify: named occupational therapy or sensory integration "
                "provision, frequency and duration of sessions, the role of the deliverer "
                "(e.g. Occupational Therapist or TA trained to deliver OT programmeme), "
                "and the specific programmeme or strategies. Environmental adjustments should "
                "be listed specifically, not described as generalised 'reasonable adjustments'."
            ),
        },
    ]

    gaps_found = []

    for area in need_areas:
        b_match = re.search(area["b_keywords"], combined_b, re.IGNORECASE)
        f_match = re.search(area["f_keywords"], combined_f, re.IGNORECASE)

        if b_match and not f_match:
            # Need identified in B with no corresponding provision in F — RED
            ctx = combined_b[max(0, b_match.start()-60):b_match.end()+120].strip()
            findings.append({
                "tier": "red",
                "title": f"Need identified in Section B with no provision in Section F — {area['name']}",
                "extract": ctx[:280],
                "commentary": (
                    f"Section B identifies {area['name'].lower()} needs but Section F contains "
                    f"no corresponding provision. Under the Children and Families Act 2014, "
                    f"every need identified in Section B must be met by provision in Section F. "
                    f"This gap must be addressed at annual review.\n\n"
                    f"What good looks like: {area['what_good_looks_like']}"
                ),
                "delivery_log_required": True,
            })
            gaps_found.append(area["name"])

        elif b_match and f_match:
            # Need present in B and F — check F provision is specific enough
            # Grab the F context around the match
            f_ctx = combined_f[max(0, f_match.start()-60):f_match.end()+120].strip()
            # Check for vague language in the matched F provision
            vague = re.search(
                r'\b(support will be provided|will be supported|will have access|'
                r'as appropriate|where necessary|as needed)\b',
                f_ctx, re.IGNORECASE
            )
            if vague:
                findings.append({
                    "tier": "amber",
                    "title": f"Section F provision for {area['name']} may lack specificity",
                    "extract": f_ctx[:280],
                    "commentary": (
                        f"Section B identifies {area['name'].lower()} needs and Section F "
                        f"references corresponding provision, but the language used may be "
                        f"insufficiently specific to be enforceable.\n\n"
                        f"What good looks like: {area['what_good_looks_like']}"
                    ),
                    "delivery_log_required": False,
                })

    # Health needs cross-reference
    health_pattern = re.compile(r'\b(health|medical|therapy|therapist|OT|SALT|physiotherapy)\b', re.IGNORECASE)
    if health_pattern.search(combined_b) and not health_pattern.search(combined_f):
        findings.append({
            "tier": "amber",
            "title": "Health needs in Section B — check Sections C and G for corresponding provision",
            "extract": "",
            "commentary": (
                "Section B describes health-related needs. Sections C and G should contain "
                "corresponding health provision and outcomes. If these sections are empty or "
                "absent, the needs identified in Section B are not being addressed through "
                "the full EHCP framework. Raise at annual review."
            ),
            "delivery_log_required": False,
        })

    return findings


def _get_short_extract(text, start, length):
    """Return a clean short extract for display."""
    extract = text[start:start+length].strip()
    # Clean up whitespace runs
    extract = re.sub(r'\n{3,}', '\n\n', extract)
    extract = re.sub(r' {2,}', ' ', extract)
    return extract


def run_full_analysis(full_text):
    """Master analysis function. Returns all findings."""
    f_blocks = find_section_blocks(full_text, "F")
    e_blocks = find_section_blocks(full_text, "E")
    b_blocks = find_section_blocks(full_text, "B")

    f_findings = analyse_section_f(f_blocks)
    e_findings = analyse_section_e(e_blocks)
    b_findings = analyse_section_b(b_blocks, f_blocks)

    all_findings = f_findings + e_findings + b_findings

    # Sort: red first, amber second, green last
    order = {"red": 0, "amber": 1, "green": 2}
    all_findings.sort(key=lambda x: order.get(x["tier"], 3))

    return all_findings, {
        "f_blocks_found": len(f_blocks),
        "e_blocks_found": len(e_blocks),
        "b_blocks_found": len(b_blocks),
    }


# ── REPORT GENERATORS ─────────────────────────────────────────────────────────

def generate_word_report(findings, child_name="your child", situation="", doc_type="EHCP"):
    """Generate Word document report."""
    doc = Document()

    # Title
    title = doc.add_heading("FRED Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph(f"Families' Rights and Entitlements Directory")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.color.rgb = RGBColor(0x1a, 0x27, 0x44)

    doc.add_paragraph(f"Document type: {doc_type}")
    if situation:
        doc.add_paragraph(f"Context: {situation}")

    doc.add_heading("Summary", 1)
    red_n   = sum(1 for f in findings if f["tier"] == "red")
    amber_n = sum(1 for f in findings if f["tier"] == "amber")
    green_n = sum(1 for f in findings if f["tier"] == "green")
    doc.add_paragraph(
        f"This report identified {red_n} lawful requirement(s) not met (Red), "
        f"{amber_n} best practice gap(s) (Amber), and {green_n} compliant area(s) (Green)."
    )

    # Delivery log note
    needs_log = any(f.get("delivery_log_required") for f in findings)
    if needs_log:
        p = doc.add_paragraph(
            "⚑  One or more findings require a delivery log. "
            "If delivery of provision is not logged, it did not happen. "
            "Request the school's delivery records immediately."
        )
        p.runs[0].bold = True

    doc.add_heading("Findings", 1)

    tier_labels = {"red": "RED — Lawful Requirement Not Met",
                   "amber": "AMBER — Best Practice Gap",
                   "green": "GREEN — Compliant"}
    tier_colors = {"red": RGBColor(0xC0, 0x39, 0x2B),
                    "amber": RGBColor(0xD4, 0xA0, 0x17),
                    "green": RGBColor(0x1E, 0x84, 0x49)}

    for finding in findings:
        tier = finding["tier"]
        h = doc.add_heading(tier_labels.get(tier, tier.upper()), 2)
        h.runs[0].font.color.rgb = tier_colors.get(tier, RGBColor(0, 0, 0))

        doc.add_heading(finding["title"], 3)

        if finding.get("extract"):
            p = doc.add_paragraph(finding["extract"])
            p.runs[0].italic = True

        doc.add_paragraph(finding["commentary"])

        if finding.get("delivery_log_required"):
            p = doc.add_paragraph(
                "Delivery log required: If this provision has been delivered, "
                "there must be a contemporaneous record. Request the delivery log."
            )
            p.runs[0].bold = True

        doc.add_paragraph("")

    doc.add_heading("What next?", 1)
    doc.add_paragraph(
        "Red findings must be addressed at annual review. "
        "Amber findings are worth raising. "
        "Green entries are compliant — use these as benchmarks when challenging non-compliant provision. "
        "\n\nThis report is produced by FRED — Families' Rights and Entitlements Directory. "
        "It provides lawful analysis, not legal advice."
    )

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def generate_pdf_report(findings, child_name="your child", situation="", doc_type="EHCP"):
    """Generate PDF report using ReportLab — substantial, well-spaced output."""
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=2.5*cm, leftMargin=2.5*cm,
        topMargin=2.5*cm, bottomMargin=2.5*cm,
        title="FRED Report",
        author="FRED — Families' Rights and Entitlements Directory",
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'fred_title', fontSize=28, alignment=TA_CENTER,
        textColor=colors.HexColor('#1a2744'),
        spaceAfter=4, spaceBefore=0,
        fontName='Helvetica-Bold',
        leading=34,
    )
    sub_style = ParagraphStyle(
        'fred_sub', fontSize=11, alignment=TA_CENTER,
        textColor=colors.HexColor('#4a5a7a'),
        spaceAfter=6, spaceBefore=0,
        fontName='Helvetica',
        leading=16,
    )
    meta_style = ParagraphStyle(
        'fred_meta', fontSize=10, alignment=TA_CENTER,
        textColor=colors.HexColor('#6a7a90'),
        spaceAfter=20, spaceBefore=0,
        fontName='Helvetica-Oblique',
        leading=14,
    )
    section_h_style = ParagraphStyle(
        'fred_section', fontSize=15, fontName='Helvetica-Bold',
        textColor=colors.HexColor('#1a2744'),
        spaceAfter=8, spaceBefore=20,
        leading=20,
        borderPad=0,
    )
    h2_red = ParagraphStyle(
        'fred_h2red', fontSize=12, fontName='Helvetica-Bold',
        textColor=colors.HexColor('#C0392B'),
        spaceAfter=4, spaceBefore=16, leading=16,
    )
    h2_amber = ParagraphStyle(
        'fred_h2amber', fontSize=12, fontName='Helvetica-Bold',
        textColor=colors.HexColor('#D4A017'),
        spaceAfter=4, spaceBefore=16, leading=16,
    )
    h2_green = ParagraphStyle(
        'fred_h2green', fontSize=12, fontName='Helvetica-Bold',
        textColor=colors.HexColor('#1E8449'),
        spaceAfter=4, spaceBefore=16, leading=16,
    )
    finding_title_style = ParagraphStyle(
        'fred_finding_title', fontSize=11, fontName='Helvetica-Bold',
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=6, spaceBefore=4, leading=15,
    )
    extract_style = ParagraphStyle(
        'fred_extract', fontSize=10, fontName='Helvetica-Oblique',
        textColor=colors.HexColor('#555555'),
        spaceAfter=8, spaceBefore=2, leading=14,
        leftIndent=12, rightIndent=12,
    )
    body_style = ParagraphStyle(
        'fred_body', fontSize=10, fontName='Helvetica',
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=10, spaceBefore=0, leading=15,
    )
    bold_style = ParagraphStyle(
        'fred_bold', fontSize=10, fontName='Helvetica-Bold',
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=8, spaceBefore=4, leading=14,
    )
    footer_style = ParagraphStyle(
        'fred_footer', fontSize=9, fontName='Helvetica-Oblique',
        textColor=colors.HexColor('#888888'),
        spaceAfter=0, spaceBefore=8, leading=13, alignment=TA_CENTER,
    )

    story = []

    # ── Cover block ───────────────────────────────────────────────────────
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph("FRED", title_style))
    story.append(Paragraph("Families' Rights and Entitlements Directory", sub_style))
    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width="100%", thickness=1.5, colour=colors.HexColor('#1a2744'), spaceAfter=8))
    story.append(Paragraph("EHCP Analysis Report", meta_style))
    story.append(Paragraph(f"Document type: {doc_type}", meta_style))
    if situation:
        story.append(Paragraph(f"Context: {situation}", meta_style))
    story.append(HRFlowable(width="100%", thickness=0.5, colour=colors.HexColor('#d0dae8'), spaceAfter=16))
    story.append(Spacer(1, 0.5*cm))

    # ── Summary ───────────────────────────────────────────────────────────
    story.append(Paragraph("Summary", section_h_style))
    story.append(HRFlowable(width="100%", thickness=0.5, colour=colors.HexColor('#d0dae8'), spaceAfter=10))

    red_n   = sum(1 for f in findings if f["tier"] == "red")
    amber_n = sum(1 for f in findings if f["tier"] == "amber")
    green_n = sum(1 for f in findings if f["tier"] == "green")

    story.append(Paragraph(
        f'<font colour="#C0392B"><b>{red_n} Red finding{"s" if red_n != 1 else ""}</b></font> — '
        f'lawful requirement{"s" if red_n != 1 else ""} not met. Must be addressed at annual review.',
        body_style
    ))
    story.append(Paragraph(
        f'<font colour="#D4A017"><b>{amber_n} Amber finding{"s" if amber_n != 1 else ""}</b></font> — '
        f'best practice gap{"s" if amber_n != 1 else ""}. Worth raising at annual review.',
        body_style
    ))
    story.append(Paragraph(
        f'<font colour="#1E8449"><b>{green_n} Green finding{"s" if green_n != 1 else ""}</b></font> — '
        f'compliant. Use as benchmark when challenging non-compliant provision.',
        body_style
    ))

    story.append(Spacer(1, 0.4*cm))

    # Delivery log alert
    needs_log = any(f.get("delivery_log_required") for f in findings)
    if needs_log:
        story.append(HRFlowable(width="100%", thickness=0.5, colour=colors.HexColor('#D4A017'), spaceAfter=6))
        story.append(Paragraph(
            "⚑  Delivery log required",
            bold_style
        ))
        story.append(Paragraph(
            "One or more findings require a delivery log. "
            "If provision has been delivered, there must be a contemporaneous record. "
            "If it is not logged, it did not happen. "
            "Request the school's delivery records immediately.",
            body_style
        ))
        story.append(HRFlowable(width="100%", thickness=0.5, colour=colors.HexColor('#D4A017'), spaceAfter=10))

    story.append(Spacer(1, 0.6*cm))

    # ── Findings ──────────────────────────────────────────────────────────
    story.append(Paragraph("Findings", section_h_style))
    story.append(HRFlowable(width="100%", thickness=0.5, colour=colors.HexColor('#d0dae8'), spaceAfter=12))

    tier_h_styles = {"red": h2_red, "amber": h2_amber, "green": h2_green}
    tier_labels   = {
        "red":   "RED — Lawful Requirement Not Met",
        "amber": "AMBER — Best Practice Gap",
        "green": "GREEN — Compliant",
    }

    for i, finding in enumerate(findings):
        tier = finding["tier"]
        story.append(Paragraph(tier_labels.get(tier, tier.upper()), tier_h_styles.get(tier, body_style)))
        story.append(Paragraph(finding['title'], finding_title_style))

        if finding.get("extract"):
            # Escape any XML special chars in extract
            safe_extract = finding["extract"][:350].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(f'"{safe_extract}"', extract_style))

        # Split commentary on double newlines for readability
        commentary_parts = finding["commentary"].split("\n\n")
        for part in commentary_parts:
            if part.strip():
                story.append(Paragraph(part.strip(), body_style))

        if finding.get("delivery_log_required"):
            story.append(Paragraph(
                "⚑  Delivery log required: if this provision has been delivered, "
                "there must be a contemporaneous record. Request the delivery log.",
                bold_style
            ))

        story.append(Spacer(1, 0.3*cm))
        story.append(HRFlowable(width="100%", thickness=0.3, colour=colors.HexColor('#e0e8f0'), spaceAfter=8))

    story.append(Spacer(1, 0.8*cm))

    # ── What next ─────────────────────────────────────────────────────────
    story.append(Paragraph("What next?", section_h_style))
    story.append(HRFlowable(width="100%", thickness=0.5, colour=colors.HexColor('#d0dae8'), spaceAfter=12))

    story.append(Paragraph(
        "<b>Red findings</b> must be addressed at your annual review. "
        "Where there are shortcomings in the delivery of specified provision, "
        "the school has a duty of care — the duty to deliver what is written in the EHCP "
        "is absolute under the Children and Families Act 2014. It is not discretionary and "
        "it does not depend on resources. If provision has not been delivered, you have the "
        "right to ask for evidence of delivery and to request that the annual review formally "
        "records the gap.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Amber findings</b> are worth raising. They are not lawful failures but they affect "
        "the quality and accountability of your child's support.",
        body_style
    ))
    story.append(Paragraph(
        "<b>Green entries</b> are your benchmarks. Use them when challenging non-compliant provision — "
        "if one section of the EHCP meets the standard, there is no reason another cannot.",
        body_style
    ))

    story.append(Spacer(1, 0.8*cm))
    story.append(HRFlowable(width="100%", thickness=1, colour=colors.HexColor('#1a2744'), spaceAfter=8))
    story.append(Paragraph(
        "This report is produced by FRED — Families' Rights and Entitlements Directory. "
        "It provides lawful analysis, not legal advice. "
        "Reference: Children and Families Act 2014 and SEND Code of Practice 2015.",
        footer_style
    ))

    doc.build(story)
    buf.seek(0)
    return buf


# ── SESSION STATE DEFAULTS ────────────────────────────────────────────────────

def init_state():
    defaults = {
        "stage": "landing",
        "findings": [],
        "parse_meta": {},
        "full_text": "",
        "doc_type": "Final EHCP",
        "situation": "",
        "email_submitted": False,
        "email_address": "",
        "survey_submitted": False,
        "show_full_report": False,
        "relationship_tone": "neutral",
        "upcoming_dates": "",
        "doc_name": "",
        "subscribed": False,
        "corr_analysed": False,
        "show_companion": False,
        "show_amendment": False,
        "show_all_patterns": False,
        "tone_override": None,
        # Document vault — named documents uploaded
        "vault": {},           # {doc_type: {"name": filename, "text": extracted_text}}
        # Correspondence thread — one entry per exchange
        "thread": [],          # [{date, direction, summary, patterns, tone_rec, reply_sent}]
        "thread_context": "",  # running summary FRED uses when new email arrives
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ── HELPER RENDERERS ──────────────────────────────────────────────────────────

def render_finding_card(finding, index=None, show_full=True):
    tier = finding["tier"]
    badge_class = f"badge-{tier}"
    card_class  = f"finding-{tier}"
    tier_labels = {
        "red":   "RED — Lawful Requirement Not Met",
        "amber": "AMBER — Best Practice Gap",
        "green": "GREEN — Compliant",
    }
    label = tier_labels.get(tier, tier.upper())

    html = f"""
    <div class="{card_class}">
      <span class="{badge_class}">{label}</span>
      <p style="font-weight:700;margin:0.4rem 0 0.5rem 0;font-size:1rem;">{finding['title']}</p>
    """
    if finding.get("extract") and show_full:
        html += f'<p style="font-style:italic;colour:#555;font-size:0.9rem;margin:0 0 0.5rem 0;">"{finding["extract"][:280]}…"</p>'

    if show_full:
        html += f'<p style="margin:0;font-size:0.95rem;line-height:1.6;">{finding["commentary"]}</p>'

    if finding.get("delivery_log_required") and show_full:
        html += '<p style="font-weight:700;margin-top:0.7rem;font-size:0.9rem;">⚑ Delivery log required</p>'

    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def render_traffic_light_explainer():
    st.markdown("### How FRED grades your report")
    cols = st.columns(3)
    with cols[0]:
        st.markdown(f"""
        <div style="border-left:5px solid {RED};padding:1rem;background:#fdf4f3;border-radius:0 6px 6px 0;">
          <span class="badge-red">RED</span>
          <p style="font-weight:700;margin:0.4rem 0 0.2rem;">Lawful requirement not met</p>
          <p style="font-size:0.9rem;margin:0;">Must be addressed at annual review. Not lawfully enforceable under the Children and Families Act 2014 as written.</p>
        </div>
        """, unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"""
        <div style="border-left:5px solid {AMBER};padding:1rem;background:#fdf9f0;border-radius:0 6px 6px 0;">
          <span class="badge-amber">AMBER</span>
          <p style="font-weight:700;margin:0.4rem 0 0.2rem;">Best practice gap</p>
          <p style="font-size:0.9rem;margin:0;">Worth raising at annual review. Not a lawful failure but a gap in best practice that may affect your child.</p>
        </div>
        """, unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f"""
        <div style="border-left:5px solid {GREEN};padding:1rem;background:#f3faf5;border-radius:0 6px 6px 0;">
          <span class="badge-green">GREEN</span>
          <p style="font-weight:700;margin:0.4rem 0 0.2rem;">Compliant</p>
          <p style="font-size:0.9rem;margin:0;">Provision meets lawful requirements. Use compliant entries as benchmarks when challenging non-compliant areas.</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)


# ── PAGES ─────────────────────────────────────────────────────────────────────

def page_landing():
    # Hero
    st.markdown("""
    <div class="hero">
      <h1>FRED</h1>
      <p class="full-name">Families' Rights and Entitlements Directory</p>
      <p class="tagline">Know what your child is entitled to.<br>Know when it's not being delivered.</p>
      <p class="origin">Built by a parent who learned the hard way — so you don't have to.</p>
      <p class="sub">
        FRED analyses your child's EHCP against the Children and Families Act 2014 and the SEND Code of Practice.
        Plain-English findings. Red, amber, green. No jargon.
      </p>
      <p class="service-line">One-off report · Annual subscription · Monthly plan</p>
    </div>
    """, unsafe_allow_html=True)

    # CTA
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("Get my report", use_container_width=True):
            st.session_state.stage = "explainer"
            st.rerun()

    st.markdown("""
    <p style="text-align:centre;colour:#6a7a90;font-size:0.9rem;margin-top:0.8rem;">
      Upload first. Decide after. Your report is ready before you pay.
    </p>
    <p style="text-align:centre;colour:#8a9ab0;font-size:0.82rem;">
      From £XX for the full report — or see our subscription plans below
    </p>
    """, unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # How it works
    st.markdown("<h2 style='text-align:centre;'>Everything you need to know.</h2>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    hiw_items = [
        "Gather whatever documents you have — your child's EHCP, EP report, OT report, school correspondence. You don't need everything.",
        "Upload your EHCP. FRED reads it against the Children and Families Act 2014 and the SEND Code of Practice.",
        "FRED gives you one real finding immediately, at no cost. You'll see exactly what kind of analysis you're getting before you decide anything.",
        "Your full report is ready. Red findings are lawful requirements not met. Amber are best practice gaps. Green are compliant — use these as your benchmarks.",
        "Download your report as a Word or PDF document. Take it to your annual review. Use it in correspondence. It's yours.",
        "A subscription gives you the full intelligence layer — correspondence analysis, email drafting, meeting preparation, document vault, and more.",
    ]

    for item in hiw_items:
        st.markdown(f"""
        <div class="hiw-item">
          <div class="hiw-dot"></div>
          <div style="font-size:0.97rem;line-height:1.6;">{item}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Traffic lights
    render_traffic_light_explainer()

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Pricing
    st.markdown("<h2 style='text-align:centre;'>Plans</h2>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    p1, p2, p3 = st.columns(3)
    with p1:
        st.markdown(f"""
        <div class="pricing-card">
          <p style="font-weight:700;font-size:1.1rem;margin-bottom:0.3rem;">One-off Report</p>
          <p class="price">£XX</p>
          <p class="period">single purchase, no commitment</p>
          <ul>
            <li>Full EHCP analysis</li>
            <li>Red / Amber / Green findings</li>
            <li>Word and PDF download</li>
            <li>No ongoing subscription</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

    with p2:
        st.markdown(f"""
        <div class="pricing-card featured">
          <span class="best-value-tag">Best value</span><br>
          <p style="font-weight:700;font-size:1.1rem;margin-bottom:0.3rem;">Annual Subscription</p>
          <p class="price">£XX</p>
          <p class="period">per year — includes first report</p>
          <ul>
            <li>Full EHCP analysis</li>
            <li>Document vault</li>
            <li>Correspondence analysis</li>
            <li>Email drafting</li>
            <li>Meeting preparation</li>
            <li>Annual review preparation</li>
            <li>Reduced renewal rate year two</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

    with p3:
        st.markdown(f"""
        <div class="pricing-card">
          <p style="font-weight:700;font-size:1.1rem;margin-bottom:0.3rem;">Monthly</p>
          <p class="price">£XX</p>
          <p class="period">first month includes report, then £XX/month</p>
          <ul>
            <li>Full EHCP analysis</li>
            <li>Document vault</li>
            <li>Correspondence analysis</li>
            <li>Email drafting</li>
            <li>Cancel anytime</li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br><p style='text-align:centre;colour:#6a7a90;font-size:0.85rem;'>No hidden charges. Your report is ready before you purchase.</p>", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # FAQ
    st.markdown("<h2 style='text-align:centre;'>FAQ</h2>", unsafe_allow_html=True)

    faqs = [
        ("When do I pay?", "After you see one real finding from your EHCP, before you access the full report. You upload first — your report is ready before you pay."),
        ("Is this legal advice?", "No. FRED provides lawful analysis — it tells you what the law requires. It does not tell you what to do about it. That decision is yours."),
        ("What documents do I need?", "Upload whatever you have. An EP report, OT report, school correspondence, diagnosis letter — FRED will work with any of these. If you don't have an EHCP yet, upload what you do have and FRED will advise on next steps to ensure the right support is in place for your child."),
        ("What is an EHCP?", "An Education, Health and Care Plan is a legally binding document setting out a child's needs and the provision that must be made. The word 'must' in that sentence is important."),
        ("What if I only have a draft EHCP?", "FRED analyses draft and final EHCPs differently. For a draft, FRED highlights what should be strengthened before the document is finalised. For a final EHCP, FRED references your rights and the duty to deliver."),
        ("Is my data private?", "Yes. Your documents are used only for your analysis. They are not shared, stored beyond your session, or used to train any system."),
    ]

    for q, a in faqs:
        with st.expander(q):
            st.write(a)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Google Form feedback link — replaces old landing survey
    st.markdown("### Want to shape what FRED becomes?")
    st.markdown("We're in beta. Leave your thoughts — takes two minutes.")
    st.markdown(f"""
    <a href="https://docs.google.com/forms/d/e/1FAIpQLSeA1F9nEdQWkmplbAh973XKq2EsW0bEkhJiw7drhP7BZaPjKQ/viewform" target="_blank" style="
        display:inline-block;
        background:#1a2744;
        colour:white;
        padding:0.7rem 1.8rem;
        border-radius:4px;
        text-decoration:none;
        font-family:'Source Sans 3',sans-serif;
        font-weight:600;
        font-size:0.95rem;
    ">Share your thoughts →</a>
    <p style="font-size:0.8rem;colour:#888;margin-top:0.5rem;">Opens in a new tab. No obligation.</p>
    """, unsafe_allow_html=True)


def page_explainer():
    st.markdown("## Before you upload")
    render_traffic_light_explainer()

    st.markdown("---")
    st.markdown("### What would you like to analyse?")

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("Continue to upload", use_container_width=True):
            st.session_state.stage = "upload"
            st.rerun()


def page_upload():
    """
    Document vault upload page.
    Named document types with plain explanation.
    Confirmation after each upload.
    """
    st.markdown("## Upload your documents")

    st.markdown(f"""
    <div style="background:#eaf5e0;border-radius:10px;padding:1rem 1.3rem;margin-bottom:1.5rem;border:0.5px solid #c0ddb0;">
      <p style="margin:0;font-size:0.95rem;colour:#2d4a2d;line-height:1.7;">
        Upload your child's EHCP first. Then add any other documents you have —
        FRED will read them all and refer back to them when you analyse correspondence
        or request an amendment.
      </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Named document uploads ────────────────────────────────────────────────
    DOCUMENT_TYPES = [
        {
            "key": "ehcp",
            "label": "EHCP",
            "description": "Essential. FRED analyses Section F provision, Section E outcomes, and Section B needs. Everything else is cross-referenced against this.",
            "required": False,
        },
        {
            "key": "ep_report",
            "label": "EP report",
            "description": "FRED checks whether the EP's recommendations have been converted into specified provision in Section F — or whether they have been laundered into vague language.",
            "required": False,
        },
        {
            "key": "ot_report",
            "label": "OT report",
            "description": "FRED checks whether OT recommendations appear in Section F and whether the OT provision loop has been closed — treatment block, discharge, re-referral threshold.",
            "required": False,
        },
        {
            "key": "sen_policy",
            "label": "School SEN policy",
            "description": "FRED cross-references the school's own SEN policy against what is in the EHCP and what the correspondence shows. The school cannot dispute its own policy.",
            "required": False,
        },
        {
            "key": "accessibility_plan",
            "label": "School accessibility plan",
            "description": "FRED checks for unfulfilled commitments — acoustic surveys, staff training, environmental audits — and flags gaps between policy commitment and practice.",
            "required": False,
        },
        {
            "key": "diagnosis",
            "label": "Diagnosis letter",
            "description": "FRED uses this to confirm the need basis for provision and to check whether the diagnosis profile is accurately reflected in Section B.",
            "required": False,
        },
    ]

    vault = st.session_state.vault
    uploaded_count = len(vault)

    for doc in DOCUMENT_TYPES:
        key  = doc["key"]
        already_uploaded = key in vault

        col1, col2 = st.columns([2, 3])
        with col1:
            label = f"{doc['label']} {'✓' if already_uploaded else '— required' if doc['required'] else '— optional'}"
            uploaded = st.file_uploader(
                label,
                type=["pdf", "docx", "txt"],
                key=f"vault_{key}",
                help=doc["description"],
            )
            if uploaded:
                text = extract_text_from_pdf(uploaded) if uploaded.name.endswith(".pdf") else extract_text_from_docx(uploaded)
                vault[key] = {"name": uploaded.name, "text": text}
                st.session_state.vault = vault

        with col2:
            if already_uploaded or uploaded:
                st.markdown(f"""
                <div style="background:#eaf5e0;border-radius:8px;padding:0.6rem 1rem;margin-top:1.6rem;border:0.5px solid #c0ddb0;">
                  <p style="margin:0;font-size:0.85rem;colour:#2d5a2d;font-weight:500;">
                    ✓ FRED has read the {doc['label']}
                  </p>
                  <p style="margin:0.2rem 0 0;font-size:0.8rem;colour:#5a8a5a;">{doc['description']}</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="padding:0.6rem 1rem;margin-top:1.6rem;">
                  <p style="margin:0;font-size:0.82rem;colour:#888;">{doc['description']}</p>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("---")

    # Summary of what's been loaded
    if vault:
        loaded = ", ".join(DOCUMENT_TYPES[i]["label"] for i, d in enumerate(DOCUMENT_TYPES) if d["key"] in vault)
        st.markdown(f"""
        <div style="background:#2d4a2d;border-radius:10px;padding:0.9rem 1.2rem;margin-bottom:1rem;">
          <p style="margin:0;font-size:0.88rem;colour:#9dc98a;font-weight:500;">
            FRED is holding: {loaded}
          </p>
          <p style="margin:0.3rem 0 0;font-size:0.82rem;colour:#7ab870;">
            These will be cross-referenced when you analyse correspondence or request an amendment.
          </p>
        </div>
        """, unsafe_allow_html=True)

    # ── Vault management ─────────────────────────────────────────────────────
    col_a, col_b = st.columns([1, 4])
    with col_a:
        if st.button("Clear vault — start again", key="clear_vault"):
            st.session_state.vault = {}
            st.rerun()
    with col_b:
        if vault:
            st.markdown(
                "<p style='font-size:0.82rem;colour:#888;padding-top:0.4rem;'>"
                "To add more documents, upload them above. To replace a document, "
                "upload a new version — it will overwrite the previous one."
                "</p>",
                unsafe_allow_html=True
            )

    # ── Analyse button ────────────────────────────────────────────────────────
    analyse_clicked = st.button("Analyse my documents", use_container_width=False, key="analyse_top")

    st.markdown("---")
    st.markdown("### A few quick questions")
    st.markdown("<p style='font-size:0.88rem;colour:#666;margin-bottom:1rem;'>These help FRED tailor the analysis. Nothing here is mandatory.</p>", unsafe_allow_html=True)

    doc_type = st.radio(
        "What is the primary document you have uploaded?",
        ["Draft EHCP", "Final EHCP", "EP report", "OT report", "School correspondence", "Other"],
        index=1,
        horizontal=True,
    )

    situation = st.text_area(
        "Briefly describe your current situation (optional)",
        placeholder="e.g. Annual review coming up in March, school not delivering speech therapy, considering appeal...",
        height=80,
    )

    upcoming_dates = st.text_input(
        "Any upcoming dates we should know about? (optional)",
        placeholder="e.g. Annual review 15 March, appeal deadline 1 April"
    )

    relationship_tone = st.select_slider(
        "How would you describe your current relationship with the school?",
        options=["Very difficult", "Difficult", "Neutral", "Generally positive", "Very positive"],
        value="Neutral",
        help="We use this to help shape the tone of any suggested emails."
    )

    st.markdown("---")

    if analyse_clicked:
        if not vault:
            st.error("Please upload at least one document to continue.")
        else:
            with st.spinner("Reading your documents…"):
                # Combine all vault text for analysis
                combined_text = " ".join(v["text"] for v in vault.values())
                findings, meta = run_full_analysis(combined_text)

            st.session_state.findings          = findings
            st.session_state.parse_meta        = meta
            st.session_state.full_text         = combined_text
            st.session_state.doc_type          = doc_type
            st.session_state.situation         = situation
            st.session_state.upcoming_dates    = upcoming_dates
            st.session_state.relationship_tone = relationship_tone
            st.session_state.doc_name          = vault.get("ehcp", {}).get("name", "EHCP")
            st.session_state.stage             = "sneak_peek"
            st.rerun()



def page_sneak_peek():
    findings = st.session_state.findings

    st.markdown("## Your report is ready")
    st.markdown(
        "Here is one real finding from your EHCP — exactly what the full report looks like. "
        "No blur. No teaser. The most significant finding from your document."
    )

    red_findings = [f for f in findings if f["tier"] == "red"]
    preview_finding = red_findings[0] if red_findings else findings[0] if findings else None

    if preview_finding:
        st.markdown('<div class="sneak-box">', unsafe_allow_html=True)
        st.markdown("**Preview — from your EHCP:**")
        render_finding_card(preview_finding, show_full=True)
        st.markdown('</div>', unsafe_allow_html=True)

    red_n   = sum(1 for f in findings if f["tier"] == "red")
    amber_n = sum(1 for f in findings if f["tier"] == "amber")
    green_n = sum(1 for f in findings if f["tier"] == "green")

    st.markdown(f"""
    <div style="background:white;border:1px solid #d0dae8;border-radius:8px;padding:1.2rem;margin:1.2rem 0;">
      <p style="margin:0;font-size:1rem;">
        Your full report contains:&nbsp;
        <span style="colour:{RED};font-weight:700;">{red_n} Red</span> &nbsp;·&nbsp;
        <span style="colour:{AMBER};font-weight:700;">{amber_n} Amber</span> &nbsp;·&nbsp;
        <span style="colour:{GREEN};font-weight:700;">{green_n} Green</span>
      </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ── Email gate ────────────────────────────────────────────────────────
    if not st.session_state.email_submitted:
        st.markdown("### Access your full report")
        st.markdown(
            "Enter your email address to access your full report. "
            "During beta, everything is free — no card details, no obligation."
        )
        st.markdown(f"""
        <div style="background:{NAVY};border-radius:8px;padding:1.5rem 2rem;margin:1rem 0;">
          <p style="colour:#a8b8d8;font-size:0.85rem;margin:0 0 0.3rem;">
            FRED BETA ACCESS
          </p>
          <p style="colour:white;font-size:1.1rem;font-weight:600;margin:0 0 1rem;">
            Your report is waiting. Enter your email to continue.
          </p>
        </div>
        """, unsafe_allow_html=True)

        with st.form("email_gate"):
            email = st.text_input(
                "Email address",
                placeholder="your@email.com",
            )
            submitted = st.form_submit_button("Access my report →", use_container_width=True)
            if submitted:
                if email and "@" in email and "." in email:
                    st.session_state.email_submitted = True
                    st.session_state.email_address   = email
                    print(f"FRED BETA SIGNUP: {email} — doc: {st.session_state.doc_name}")
                    st.rerun()
                else:
                    st.error("Please enter a valid email address.")
    else:
        st.session_state.stage = "full_report"
        st.rerun()


def page_full_report():
    findings  = st.session_state.findings
    doc_type  = st.session_state.doc_type
    situation = st.session_state.situation
    meta      = st.session_state.parse_meta

    red_n   = sum(1 for f in findings if f["tier"] == "red")
    amber_n = sum(1 for f in findings if f["tier"] == "amber")
    green_n = sum(1 for f in findings if f["tier"] == "green")

    st.markdown("## Your full FRED report")

    # Parse meta info
    st.markdown(f"""
    <div style="background:#f0f4fa;border-radius:6px;padding:0.8rem 1.2rem;margin-bottom:1.2rem;font-size:0.9rem;colour:#444;">
      Document type: <b>{doc_type}</b> &nbsp;·&nbsp;
      Section F blocks found: <b>{meta.get('f_blocks_found', 0)}</b> &nbsp;·&nbsp;
      Section E blocks found: <b>{meta.get('e_blocks_found', 0)}</b>
    </div>
    """, unsafe_allow_html=True)

    # Summary
    st.markdown(f"""
    <div style="background:white;border:1px solid #d0dae8;border-radius:8px;padding:1.2rem;margin-bottom:1.5rem;">
      <p style="margin:0;font-size:1.05rem;font-weight:600;">Summary</p>
      <p style="margin:0.5rem 0 0;font-size:0.97rem;">
        <span style="colour:{RED};font-weight:700;">{red_n} Red</span> — lawful requirement(s) not met.
        Must be addressed at annual review.<br>
        <span style="colour:{AMBER};font-weight:700;">{amber_n} Amber</span> — best practice gap(s).
        Worth raising at annual review.<br>
        <span style="colour:{GREEN};font-weight:700;">{green_n} Green</span> — compliant.
        Use as benchmark when challenging non-compliant provision.
      </p>
    </div>
    """, unsafe_allow_html=True)

    needs_log = any(f.get("delivery_log_required") for f in findings)
    if needs_log:
        st.markdown(f"""
        <div style="background:#fff8e1;border:1px solid {AMBER};border-radius:6px;padding:0.9rem 1.2rem;margin-bottom:1.2rem;">
          <p style="margin:0;font-weight:700;">⚑  Delivery log required</p>
          <p style="margin:0.4rem 0 0;font-size:0.9rem;">
            One or more findings require a delivery log.
            If provision has been delivered, there must be a contemporaneous record.
            If not logged, it did not happen.
            Request the school's delivery records immediately.
          </p>
        </div>
        """, unsafe_allow_html=True)

    # Draft-specific note
    if "Draft" in doc_type:
        st.info(
            "This is a draft EHCP. Red findings indicate language and gaps that should be "
            "strengthened before the document is finalised. You have a window to request amendments."
        )

    st.markdown("---")
    st.markdown("### Findings")

    # Render all findings — green only shown if green_n > 0
    for i, finding in enumerate(findings):
        if finding["tier"] == "green" and green_n == 0:
            continue
        render_finding_card(finding, index=i, show_full=True)

    st.markdown("---")
    st.markdown("### What next?")
    st.markdown(
        "**Red findings** must be addressed at your annual review. "
        "Where there are shortcomings in the delivery of specified provision, "
        "the school has a duty of care — the duty to deliver what is written in the EHCP "
        "is absolute under the Children and Families Act 2014. It is not discretionary and "
        "it does not depend on resources. If provision has not been delivered, you have the "
        "right to ask for evidence of delivery and to request that the annual review formally "
        "records the gap."
        "\n\n"
        "**Amber findings** are worth raising. They are not lawful failures but they affect "
        "the quality and accountability of your child's support."
        "\n\n"
        "**Green entries** are your benchmarks. Use them when challenging non-compliant provision — "
        "if one section of the EHCP meets the standard, there is no reason another cannot."
    )

    # ── Subscriber prompt ─────────────────────────────────────────────────
    st.markdown("---")
    if not st.session_state.subscribed:
        st.markdown(f"""
        <div style="background:{NAVY};border-radius:8px;padding:2rem;margin:1rem 0;text-align:centre;">
          <p style="colour:#a8b8d8;font-size:0.85rem;margin:0 0 0.4rem;letter-spacing:0.1em;text-transform:uppercase;">
            FRED BETA — FREE ACCESS
          </p>
          <p style="colour:white;font-size:1.3rem;font-weight:700;margin:0 0 0.6rem;">
            Want the full FRED experience?
          </p>
          <p style="colour:#c8d8f0;font-size:1rem;margin:0 0 1.2rem;max-width:480px;margin-left:auto;margin-right:auto;">
            Subscribe free during beta and unlock correspondence analysis,
            email drafting, meeting preparation, document vault, and more.
          </p>
          <p style="colour:#a8b8d8;font-size:0.85rem;margin:0;">
            Already registered as {st.session_state.email_address if st.session_state.email_address else "a beta tester"}
          </p>
        </div>
        """, unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1,1,1])
        with col2:
            if st.button("Subscribe free — unlock everything", use_container_width=True):
                st.session_state.subscribed = True
                st.session_state.stage = "subscriber"
                st.rerun()
    else:
        if st.button("Go to my FRED workspace →"):
            st.session_state.stage = "subscriber"
            st.rerun()

    # ── Google Form feedback ──────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Help us build FRED properly")
    st.markdown(
        "Two minutes. Completely optional. "
        "Your feedback shapes what FRED becomes."
    )
    st.markdown(f"""
    <a href="https://docs.google.com/forms/d/e/1FAIpQLSeA1F9nEdQWkmplbAh973XKq2EsW0bEkhJiw7drhP7BZaPjKQ/viewform" target="_blank" style="
        display:inline-block;
        background:{NAVY};
        colour:white;
        padding:0.7rem 1.8rem;
        border-radius:4px;
        text-decoration:none;
        font-family:'Source Sans 3',sans-serif;
        font-weight:600;
        font-size:0.95rem;
    ">Leave feedback →</a>
    <p style="font-size:0.8rem;colour:#888;margin-top:0.5rem;">
        Opens in a new tab. Takes 2 minutes.
    </p>
    """, unsafe_allow_html=True)

    # Downloads
    st.markdown("---")
    st.markdown("### Download your report")

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        word_buf = generate_word_report(findings, doc_type=doc_type, situation=situation)
        st.download_button(
            label="Download as Word",
            data=word_buf,
            file_name="FRED_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    with col2:
        pdf_buf = generate_pdf_report(findings, doc_type=doc_type, situation=situation)
        st.download_button(
            label="Download as PDF",
            data=pdf_buf,
            file_name="FRED_report.pdf",
            mime="application/pdf",
        )

    st.markdown(
        "<p style='font-size:0.85rem;colour:#888;'>Word is best for Windows/Office users. PDF is best for Apple devices.</p>",
        unsafe_allow_html=True
    )

    # Subscription prompt
    st.markdown("---")
    st.markdown("### Want the full intelligence layer?")
    st.markdown(
        "A FRED subscription gives you correspondence analysis, email drafting, "
        "meeting preparation, post-meeting summaries, document vault, and annual review preparation. "
        "Subscriptions open at launch."
    )




def page_survey():
    st.markdown("### Your feedback")
    st.markdown("This takes two minutes and helps us build FRED properly.")

    with st.form("survey_form"):
        q1 = st.radio("Did the report identify anything you didn't already know?",
                      ["Yes, something new", "Confirmed what I suspected", "Nothing new", "Not sure"])
        q2 = st.radio("Were the traffic light grades (Red / Amber / Green) clear?",
                      ["Very clear", "Mostly clear", "Somewhat unclear", "Unclear"])
        q3 = st.radio("Was the layout simple to follow?",
                      ["Yes", "Mostly", "No"])
        q4 = st.radio("Would you pay for this report?",
                      ["Yes", "Possibly", "No", "Already would have paid"])
        q5 = st.text_input("What would a fair price be for the full report? (optional)", placeholder="e.g. £15, £25")
        q6 = st.radio("Would you subscribe for ongoing access?",
                      ["Yes", "Possibly", "No"])
        q7 = st.radio("What personalisation matters most to you?",
                      ["Colour theme", "Text size", "Both", "Not bothered"])
        q8 = st.text_area("Anything else you'd like to tell us? (optional)", height=80)

        st.markdown("---")
        st.markdown("**Would you like to be notified when FRED launches?**")
        notify = st.radio("", ["Yes, notify me", "No thank you"], key="survey_notify")
        notify_email = ""
        if notify == "Yes, notify me":
            notify_email = st.text_input("Email address for launch notification", key="survey_email")

        submitted = st.form_submit_button("Submit feedback")
        if submitted:
            extra_dict = {
                "new_info": q1, "tl_clear": q2, "layout": q3,
                "would_pay": q4, "fair_price": q5, "subscribe": q6,
                "personalise": q7, "other": q8, "notify": notify,
            }
            email_to_send = notify_email if notify_email else "no-email@fred.invalid"
            send_emailjs(email_to_send, "survey", extra_dict)
            st.session_state.survey_submitted = True
            st.success("Thank you. Your feedback helps us build FRED properly.")


# ── CORRESPONDENCE MODULE ─────────────────────────────────────────────────────

def detect_tone_recommendation(text, patterns):
    """
    Auto-detect recommended response tone from correspondence language and patterns.
    Returns dict: {recommendation, reasoning, confidence}
    """
    red_count   = sum(1 for p in patterns if p["tier"] == "red")
    amber_count = sum(1 for p in patterns if p["tier"] == "amber")

    # Signals for formal response
    formal_signals = re.compile(
        r"\b(following legal advice|our solicitor|legal services|"
        r"we are unable|complaints procedure|without prejudice|"
        r"we must inform|we are not in a position|governing body|"
        r"we refute|we dispute|we deny|permanent exclusion|"
        r"fixed term exclusion|resources do not allow)\b",
        re.IGNORECASE
    )

    # Signals for collaborative response
    collab_signals = re.compile(
        r"\b(we would like to|working together|we appreciate|"
        r"we understand your concerns|happy to discuss|"
        r"we are committed|please let us know|we value|"
        r"thank you for|we recognise|we are sorry|"
        r"we will look into|we want to support)\b",
        re.IGNORECASE
    )

    formal_count = len(formal_signals.findall(text))
    collab_count = len(collab_signals.findall(text))

    if red_count >= 2 or formal_count >= 2:
        return {
            "recommendation": "formal",
            "label": "Formal written response",
            "reasoning": (
                f"The correspondence contains {'serious patterns' if red_count >= 2 else 'formal or legal language'} "
                f"that suggest the school is managing rather than engaging. "
                f"A formal written response that references Section F specifically "
                f"is likely to be more effective than a collaborative one. "
                f"Put everything in writing and request a written response within five working days."
            ),
            "confidence": "high" if red_count >= 2 and formal_count >= 1 else "moderate",
        }
    elif collab_count >= 2 and red_count == 0:
        return {
            "recommendation": "collaborative",
            "label": "Collaborative response",
            "reasoning": (
                "The language in this correspondence is constructive and appears to be "
                "engaging with your concerns. A collaborative tone is likely to preserve "
                "the relationship while still holding the school to account. "
                "You can be warm and firm at the same time."
            ),
            "confidence": "moderate",
        }
    else:
        return {
            "recommendation": "neutral",
            "label": "Measured response",
            "reasoning": (
                "The correspondence does not show strong signals in either direction. "
                "A measured written response — factual, specific, referencing the EHCP — "
                "is appropriate. Neither overly formal nor overly warm."
            ),
            "confidence": "low",
        }


def add_to_thread(direction, summary, patterns, tone_rec, reply_sent=""):
    """Add an exchange to the correspondence thread."""
    if "thread" not in st.session_state:
        st.session_state.thread = []
    entry = {
        "date": datetime.datetime.now().strftime("%d %B %Y"),
        "direction": direction,   # "from_school" or "to_school"
        "summary": summary,
        "patterns": [p["name"] for p in patterns],
        "tone_rec": tone_rec.get("recommendation", ""),
        "reply_sent": reply_sent,
    }
    st.session_state.thread.append(entry)
    # Update running context summary
    summaries = [f"{e['date']}: {e['summary']}" for e in st.session_state.thread[-5:]]
    st.session_state.thread_context = " | ".join(summaries)


def check_thread_for_similar(patterns, environment):
    """
    Check whether current correspondence matches anything in the thread history.
    Returns a list of similar past entries.
    """
    if not st.session_state.get("thread"):
        return []
    current_pattern_names = {p["name"] for p in patterns}
    similar = []
    for entry in st.session_state.thread[:-1]:  # exclude current
        past_names = set(entry.get("patterns", []))
        overlap = current_pattern_names & past_names
        if overlap or (environment and environment in entry.get("summary", "")):
            similar.append({**entry, "overlap": list(overlap)})
    return similar


# ── CORRESPONDENCE PATTERN LIBRARY ───────────────────────────────────────────
# Built from real tribunal bundle — Duffy v Brookhurst Primary School 2019

CORRESPONDENCE_PATTERNS = [
    {
        "id": "resources_defence",
        "name": "The Resources Defence",
        "tier": "red",
        "triggers": r"\b(resources|cost involved|beyond its requirements|school budget|cannot fund|not funded|staffing constraints|financial)\b",
        "explanation": "The school is suggesting that cost or resources limit what they can provide. The duty to deliver specified EHCP provision is absolute under the Children and Families Act 2014. It does not depend on the school's budget. Where additional funding is needed, the responsibility lies with the LA.",
        "fred_question": "Ask the school in writing to identify which provisions in Section F they are delivering and which they are not delivering due to resource constraints. Any gap becomes a matter for the LA to fund.",
    },
    {
        "id": "relationship_threat",
        "name": "The Relationship Breakdown Threat",
        "tier": "red",
        "triggers": r"\b(placement.*fail|break.*down|irrevocably|if.*situation.*not.*change|not.*engage|refuse.*communicate|consequences)\b",
        "explanation": "The school is implying that asserting your rights will damage the placement. This inverts responsibility. The duty to deliver provision exists regardless of whether parents are in dispute. A threat of placement breakdown in response to a parent exercising their rights is pressure, not professional dialogue.",
        "fred_question": "Note the date and exact wording. Respond in writing acknowledging receipt only. Do not respond to the threat. Ask instead for written confirmation of what provision is currently being delivered and what the support plan is.",
    },
    {
        "id": "best_interests_redirect",
        "name": "The Best Interests Redirect",
        "tier": "amber",
        "triggers": r"\b(best interest|not right for|another.*placement|more.*suitable|alternative.*provision|better.*placed|different.*school)\b",
        "explanation": "The school is using welfare language to suggest your child would be better placed elsewhere. Any professional opinion about placement must be made through the EHCP review process with evidence — not in correspondence responding to a complaint.",
        "fred_question": "Ask what evidence base supports this view, who assessed it, and whether it has been formally submitted to the LA as part of the EHCP review process.",
    },
    {
        "id": "blip_normalisation",
        "name": "The Blip Normalisation",
        "tier": "amber",
        "triggers": r"\b(only a blip|usual for his age|usual for her age|all children|typical.*behaviour|normal.*development|peers.*same|nothing.*unusual)\b",
        "explanation": "The school is comparing your child to neurotypical peers to minimise the significance of an incident. Your child has an EHCP precisely because their needs are not the same as their peers. 'Usual for their age' is not a relevant benchmark.",
        "fred_question": "Ask the school to confirm whether this assessment was made with reference to your child's EHCP profile and diagnosis. Ask them to reconsider in light of Section B of the plan.",
    },
    {
        "id": "complaint_deflection",
        "name": "The Complaint Policy Deflection",
        "tier": "amber",
        "triggers": r"\b(complaints policy|complaints procedure|formal complaint|pursue.*further|following.*legal.*advice|refer you to)\b",
        "explanation": "The school is directing you to the complaints procedure rather than engaging with the substance. The complaints procedure does not pause statutory obligations or EHCP timelines. It is a separate process.",
        "fred_question": "Note that the complaints procedure and the statutory EHCP process are separate. Continue through the EHCP route in parallel. Respond confirming you have noted the reference and are continuing to seek resolution of the provision concern.",
    },
    {
        "id": "good_week_signal",
        "name": "The Good Week Signal",
        "tier": "amber",
        "triggers": r"\b(good week|positive week|doing well|made good progress|responding well|settled week|happy in school)\b",
        "explanation": "Positive weekly summaries create a paper trail that can contradict the severity of your child's needs if an incident occurs in the same period. A log that records only positives is not an accurate delivery record.",
        "fred_question": "Ask for the full incident record alongside positive weekly summaries. If both a positive summary and a behavioural incident occurred in the same week, ask the school to explain the discrepancy.",
    },
    {
        "id": "reintegration_promise",
        "name": "The Reintegration Promise Without a Plan",
        "tier": "red",
        "triggers": r"\b(reintegration|return.*plan|back.*school|phased.*return|managed.*return|support.*return)\b",
        "explanation": "A verbal or vague reintegration promise is not a reintegration plan. A child with an EHCP who has been excluded requires a documented reintegration plan that references the EHCP provision. Without it there is no Do stage.",
        "fred_question": "Request the written reintegration plan within five working days, referencing the EHCP explicitly. If no written plan exists, confirm that in writing and ask when one will be provided.",
    },
    {
        "id": "staffing_change",
        "name": "Unstated Staffing Change",
        "tier": "red",
        "triggers": r"\b(new.*staff|change.*support|different.*adult|new.*TA|new.*teaching assistant|cover|temporary|supply)\b",
        "explanation": "For a child with ASD or similar needs, key adult relationships are often specified in the EHCP. A staffing change affecting 1:1 support is a provision change. It should not happen without parents being informed and without a transition plan.",
        "fred_question": "Ask the school to confirm in writing whether any staffing changes affecting your child's 1:1 support have been made since the last annual review, and what transition planning was put in place.",
    },
    {
        "id": "legal_misrepresentation",
        "name": "Legal Obligation Misrepresented",
        "tier": "red",
        "triggers": r"\b(no legal obligation|not required by law|not a legal requirement|beyond what is required|not legally obliged|discretionary)\b",
        "explanation": "The school may be correctly stating there is no obligation to provide 'the best possible education' — but this is not the relevant obligation. The obligation under the Children and Families Act 2014 is to deliver every provision specified in Section F. That obligation is absolute, not discretionary.",
        "fred_question": "Ask the school to confirm delivery of each named provision in Section F. The question is not whether they are providing the best possible education — it is whether they are delivering what the EHCP specifies.",
    },
    {
        "id": "homeschool_discrepancy",
        "name": "Home/School Discrepancy Used Against Parents",
        "tier": "amber",
        "triggers": r"\b(home.*school.*different|parents.*expectations|at home.*different|parental.*concern|parents.*report|home.*school.*gap)\b",
        "explanation": "The school is using differences between home and school presentation to question parental credibility. A child who presents differently at home and at school may be experiencing a stress response that is discharged at home. This is a clinical question, not a credibility question.",
        "fred_question": "Ask the school what steps they have taken to investigate the home/school discrepancy as a clinical question. Ask whether it has been formally assessed and what the findings were.",
    },
    {
        "id": "reassurance_without_evidence",
        "name": "Reassurance Without Evidence",
        "tier": "amber",
        "triggers": r"\b(we are not concerned|no concerns at this time|fully supported|meeting his needs|meeting her needs|you would be our first|runs continuously|we monitor closely)\b",
        "explanation": "The correspondence contains reassurance language without documentary support. Reassurance is not evidence. An assertion that provision is in place or that needs are being met requires a delivery log, not a statement of confidence.",
        "fred_question": "Ask what the evidence base is for this statement. Request the delivery log for the relevant provision over the past half term. If no log exists, note that in writing.",
    },
    {
        "id": "implicit_admission",
        "name": "Implicit Admission of Current Gap",
        "tier": "amber",
        "triggers": r"\b(going forward|in future|we will ensure|we have identified|we are working to improve|we recognise|we are reviewing|steps have been taken)\b",
        "explanation": "This language implies current practice is inadequate while presenting future improvement as reassurance. A commitment to improve in future is an admission of a current gap. Note the date — this correspondence is evidence that the gap existed at this point.",
        "fred_question": "Ask the school to confirm in writing what the current position is — not what it will be. Note that this correspondence has been retained and dated.",
    },
    {
        "id": "behaviour_framing",
        "name": "SEND Need Framed as Behaviour Problem",
        "tier": "red",
        "triggers": r"\b(behaviour.*pattern|pattern.*behaviour|conduct|expectations.*behaviour|"
                    r"struggling to meet expectations|behaviour.*concern|behavioural incident|"
                    r"standards expected|reminded.*standards|reflect on his behaviour|"
                    r"reflect on her behaviour|out of circulation|time to reflect)\b",
        "explanation": (
            "The school is describing what may be a SEND need as a behaviour problem. "
            "For a child with an EHCP, any incident during unstructured time — lunchtime, "
            "break, transitions — must first be examined against what provision was in place "
            "at that moment. A child with documented sensory or social communication needs "
            "who becomes dysregulated in an unstructured environment is not misbehaving — "
            "they are experiencing a provision failure."
        ),
        "fred_question": (
            "Ask the school what provision from Section F was in place during the unstructured "
            "period when this incident occurred. If 1:1 support, sensory breaks, or structured "
            "lunchtime activity is specified in the EHCP, ask whether it was being delivered "
            "at the time of the incident."
        ),
    },
    {
        "id": "veiled_threat",
        "name": "Veiled Threat — Next Steps Language",
        "tier": "red",
        "triggers": r"\b(next steps|may need to consider|if.*persist|if.*continue|"
                    r"cannot continue|placement.*appropriate|alternative.*provision|"
                    r"exploring.*options|reviewing.*placement|not the right environment|"
                    r"may not be meeting|current environment.*not)\b",
        "explanation": (
            "The correspondence contains language that implies escalation — toward exclusion, "
            "alternative provision, or placement change — dressed as concern. "
            "'Next steps in terms of support and provision' following a behavioural incident "
            "is a soft version of a placement threat. Note the date and exact wording. "
            "This creates a paper trail that the school may use to justify future action."
        ),
        "fred_question": (
            "Ask the school to clarify in writing what 'next steps' means specifically. "
            "Is the school considering a change to provision, a change to placement, or "
            "a referral? Each has different statutory implications. "
            "Any change to provision must go through the EHCP review process."
        ),
    },
    {
        "id": "home_responsibility_redirect",
        "name": "Home Responsibility Redirect",
        "tier": "amber",
        "triggers": r"\b(reinforce at home|support at home|work with.*at home|"
                    r"we would ask.*home|parental.*support|home.*behaviour|"
                    r"consistent.*home|same.*home|boundaries at home|routines at home)\b",
        "explanation": (
            "The school is redirecting responsibility for a school-based incident to the home. "
            "School behaviour during school hours is a school provision question, not a home "
            "behaviour question. An incident in the dining hall is governed by what provision "
            "was in place in the dining hall — not by what happens at home."
        ),
        "fred_question": (
            "Note this redirect. Respond by asking what provision was in place at school "
            "during the incident rather than engaging with the home behaviour framing. "
            "The EHCP provision is a school responsibility, not a shared one."
        ),
    },
    {
        "id": "unstructured_time_gap",
        "name": "Incident During Unstructured Time — Provision Check",
        "tier": "red",
        "triggers": r"\b(lunchtime|lunch time|break time|breaktime|unstructured|"
                    r"less structured|playground|dinner time|free time|between lessons|"
                    r"transition.*incident|incident.*transition|corridor.*incident)\b",
        "explanation": (
            "The incident occurred during an unstructured period — lunchtime, break, or "
            "transition. For a child with ASD, SEMH needs, or sensory processing difficulties, "
            "unstructured time is consistently the highest-risk period of the school day. "
            "The EHCP should specify what support is in place during these periods. "
            "If it does not, that is a provision gap. If it does and the support was not "
            "in place, that is a delivery failure."
        ),
        "fred_question": (
            "Ask the school to confirm in writing what provision from Section F is specifically "
            "in place during lunchtime, break time, and transitions. "
            "Then ask whether that provision was being delivered at the time of the incident. "
            "Request the delivery log for the relevant period."
        ),
    },
    {
        "id": "monitoring_without_action",
        "name": "Monitor and Observe — No Specified Action",
        "tier": "amber",
        "triggers": r"\b(monitor.*closely|continue to monitor|keep an eye|observe.*situation|"
                    r"watching.*closely|review.*situation|situation.*closely|monitor.*progress)\b",
        "explanation": (
            "The school has committed to monitoring without specifying what action will be "
            "taken or what threshold will trigger a review. 'We will monitor closely' is "
            "not a provision — it is an observation. Monitoring without a specified response "
            "plan is not a graduated approach under the SEND Code of Practice."
        ),
        "fred_question": (
            "Ask the school what specifically they will be monitoring, what the review "
            "mechanism is, who is responsible, and what action will be taken if the "
            "monitoring identifies further difficulty. Ask for this in writing."
        ),
    },
]

ENVIRONMENT_TRIGGERS = {
    "food hall": "food hall / canteen",
    "canteen": "food hall / canteen",
    "dining hall": "food hall / canteen",
    "dining room": "food hall / canteen",
    "lunch hall": "food hall / canteen",
    "corridor": "corridor / transition space",
    "hallway": "corridor / transition space",
    "playground": "playground / outdoor space",
    "outside": "playground / outdoor space",
    "assembly": "assembly hall",
    "hall": "assembly hall",
    "classroom": "classroom environment",
    "gym": "sports hall / gym",
    "sports hall": "sports hall / gym",
    "pe": "sports hall / gym",
    "toilet": "toilet facilities",
    "bathroom": "toilet facilities",
    "bus": "school transport",
    "minibus": "school transport",
}

SENSORY_CHECKLIST = {
    "food hall / canteen": [
        ("Sound", [
            "Overall noise level — crowded dining hall acoustics",
            "Cutlery, chair scraping, tray noise",
            "Echo and reverberation in the space",
            "Kitchen extraction fans and cooking sounds",
            "Unpredictable loud sounds — dropped trays, shouting",
        ]),
        ("Light", [
            "Strip lighting or fluorescent flicker",
            "Overall brightness",
            "Glare from windows or surfaces",
        ]),
        ("Smell", [
            "Food odours — specific foods that may be aversive",
            "Cleaning products",
            "Accumulated smell of multiple food types",
        ]),
        ("Physical", [
            "Crowding — proximity to other pupils",
            "Queue management — unpredictable movement",
            "Seating — choice or assigned, proximity to others",
            "Clear exit route visible",
        ]),
        ("Timing", [
            "Hunger state on arrival — time since last food",
            "Blood sugar — snack provided before lunch?",
            "Preceding activity — what happened immediately before",
        ]),
        ("Predictability", [
            "Does the child know what food is available before arriving",
            "Consistent routine day to day",
            "Visual support for the lunch process",
        ]),
    ],
    "corridor / transition space": [
        ("Sound", [
            "Noise level during transition — multiple classes moving",
            "Unpredictable sounds — lockers, doors, shouting",
            "Echo in narrow spaces",
        ]),
        ("Physical", [
            "Crowding during peak transition times",
            "Physical contact from other pupils",
            "Width of corridor — personal space available",
        ]),
        ("Predictability", [
            "Consistency of route taken",
            "Whether the child knows destination in advance",
            "Early departure to avoid peak crowding",
            "Named adult support during transition",
        ]),
    ],
    "playground / outdoor space": [
        ("Sound", [
            "Unpredictable noise — shouting, ball games",
            "Wind noise",
        ]),
        ("Social", [
            "Unstructured peer interaction — no adult mediation",
            "Safe space available if overwhelmed",
            "Structured activity available as alternative",
        ]),
        ("Sensory", [
            "Bright sunlight — visual sensitivity",
            "Temperature extremes",
            "Physical play — proprioceptive seeking or avoiding",
        ]),
    ],
}

DEFAULT_CHECKLIST = [
    ("Sound", ["Overall noise level", "Unpredictable sounds", "Echo or reverberation"]),
    ("Light", ["Brightness", "Fluorescent lighting", "Glare"]),
    ("Smell", ["Specific odours that may be aversive"]),
    ("Physical", ["Crowding", "Personal space", "Exit routes"]),
    ("Predictability", ["Routine consistency", "Advance warning of changes"]),
    ("Timing", ["Time of day", "Preceding activity", "Hunger or fatigue state"]),
]


def detect_environment(text):
    text_lower = text.lower()
    for trigger, env_name in ENVIRONMENT_TRIGGERS.items():
        if trigger in text_lower:
            return env_name
    return None


def get_checklist(env_name):
    return SENSORY_CHECKLIST.get(env_name, DEFAULT_CHECKLIST)


def detect_patterns(text, ehcp_text=""):
    """
    Run all correspondence patterns against text.
    If EHCP text provided, adds EHCP cross-reference findings.
    Returns list sorted red first.
    """
    matched = []
    for pattern in CORRESPONDENCE_PATTERNS:
        if re.search(pattern["triggers"], text, re.IGNORECASE):
            m = re.search(pattern["triggers"], text, re.IGNORECASE)
            ctx = ""
            if m:
                ctx = text[max(0, m.start()-100):m.end()+120].strip()
            matched.append({**pattern, "extract": ctx[:250]})

    # EHCP cross-reference — unstructured time
    if ehcp_text and re.search(
        r"\b(lunchtime|lunch time|break time|unstructured|dining hall|canteen|playground)\b",
        text, re.IGNORECASE
    ):
        ehcp_lunch = re.search(
            r"\b(lunchtime|unstructured|break.*supervision|lunch.*support|"
            r"lunch.*adult|supervision.*lunch|1:1.*lunch|full.time.*support)\b",
            ehcp_text, re.IGNORECASE
        )
        if ehcp_lunch:
            matched.append({
                "id": "ehcp_xref_lunch",
                "name": "EHCP Cross-Reference — Lunchtime Provision",
                "tier": "red",
                "triggers": "",
                "explanation": (
                    "Your EHCP references provision during lunchtime or unstructured periods. "
                    "This incident occurred during exactly that period. "
                    "The question is not whether the behaviour occurred — it is whether the "
                    "provision specified in the EHCP was being delivered when it did."
                ),
                "fred_question": (
                    "Request the delivery log for lunchtime provision on the day of this incident. "
                    "Ask the school to confirm in writing whether the named provision from "
                    "Section F was in place at the time. If it was not, that is a delivery "
                    "failure, not a behaviour problem."
                ),
                "extract": "",
            })
        else:
            matched.append({
                "id": "ehcp_xref_lunch_gap",
                "name": "EHCP Gap — No Lunchtime Provision Specified",
                "tier": "red",
                "triggers": "",
                "explanation": (
                    "This incident occurred during lunchtime but the EHCP does not appear "
                    "to specify provision for this period. For a child with documented "
                    "sensory or social communication needs, unstructured lunchtime is a "
                    "high-risk period. The absence of specified provision is itself a gap "
                    "that should be addressed at the next annual review."
                ),
                "fred_question": (
                    "At the next annual review, request that Section F is amended to specify "
                    "what support is in place during all unstructured periods — lunchtime, "
                    "break, and transitions. Ask the school in writing what support was "
                    "in place at the time of this incident."
                ),
                "extract": "",
            })

    # Sort: red first, amber second
    order = {"red": 0, "amber": 1, "green": 2}
    matched.sort(key=lambda x: order.get(x["tier"], 3))
    return matched


def analyse_policy(policy_text):
    findings = []
    audit_pattern = re.compile(
        r"\b(will conduct|will undertake|will carry out|will review|annual.*audit|"
        r"acoustic.*survey|accessibility.*audit|risk.*assessment.*will|will.*assess|plan to|intend to)\b",
        re.IGNORECASE
    )
    for m in audit_pattern.finditer(policy_text):
        ctx = policy_text[max(0, m.start()-60):m.end()+100].strip()
        findings.append({
            "tier": "amber",
            "title": "Policy commitment — check whether actioned",
            "extract": ctx[:280],
            "commentary": (
                "This policy commits to an action — audit, survey, or review. "
                "Ask the school for evidence it has been completed and the date it was last carried out. "
                "If it has not been done, this is a gap between policy commitment and practice."
            ),
        })
        break  # One instance is enough

    year_pattern = re.compile(r"\b(20\d{2})\b")
    years = [int(y) for y in year_pattern.findall(policy_text)]
    current_year = datetime.datetime.now().year
    if years:
        most_recent = max(years)
        age = current_year - most_recent
        if age >= 2:
            findings.append({
                "tier": "amber",
                "title": f"Policy is {age} years old — commitments may be overdue",
                "extract": f"Most recent year reference: {most_recent}",
                "commentary": (
                    f"This policy was last updated in {most_recent}. Any commitments — "
                    f"audits, assessments, training — should have been actioned in the "
                    f"intervening {age} years. Ask for evidence of actions taken since that date."
                ),
            })

    if re.search(r"\b(acoustic|strip light|fluorescent|sensory|ear defender|quiet area|calm space|accessibility)\b", policy_text, re.IGNORECASE):
        findings.append({
            "tier": "green",
            "title": "Policy references sensory or accessibility provision",
            "extract": "",
            "commentary": (
                "The policy references sensory or accessibility provision. "
                "Cross-reference this against your child's EHCP Section F. "
                "If the policy commits to provision that is absent from Section F, "
                "that gap is evidenced by the school's own document."
            ),
        })

    if re.search(r"\b(staff.*training|trained.*staff|autism.*training|SEND.*training|awareness.*training)\b", policy_text, re.IGNORECASE):
        findings.append({
            "tier": "amber",
            "title": "Policy commits to staff training — ask for evidence",
            "extract": "",
            "commentary": (
                "The policy references staff training in relation to SEND or specific needs. "
                "Ask the school to confirm what training has been provided to staff working "
                "directly with your child, and when it was last updated."
            ),
        })

    return findings


def render_pattern_card(p, index):
    tier = p["tier"]
    tier_label = "RED — Lawful concern" if tier == "red" else "AMBER — Pattern detected"
    st.markdown(f"""
    <div class="finding-{tier}" style="margin-bottom:1.2rem;">
      <span class="badge-{tier}">{tier_label}</span>
      <p style="font-weight:700;margin:0.5rem 0 0.3rem;font-size:1rem;">{p['name']}</p>
      <p style="margin:0 0 0.6rem;font-size:0.93rem;line-height:1.6;">{p['explanation']}</p>
      <div style="background:rgba(0,0,0,0.04);border-radius:4px;padding:0.6rem 0.8rem;margin-bottom:0.4rem;">
        <p style="font-weight:700;margin:0 0 0.2rem;font-size:0.88rem;text-transform:uppercase;letter-spacing:0.05em;">The question to ask:</p>
        <p style="margin:0;font-size:0.93rem;font-style:italic;">{p['fred_question']}</p>
      </div>
      {f'<p style="font-size:0.82rem;colour:#666;margin:0.4rem 0 0;border-left:3px solid #ddd;padding-left:0.6rem;">{p["extract"]}…</p>' if p.get("extract") else ""}
    </div>
    """, unsafe_allow_html=True)


def generate_amendment_word(environment, confirmed_items, today):
    """Generate Word document for amendment request."""
    doc = Document()
    title = doc.add_heading("EHCP Amendment Request", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Date: {today}")
    doc.add_paragraph(f"Re: Request to amend Section B and Section F — {environment}")
    doc.add_paragraph("")

    doc.add_paragraph("Dear [SENCO name],")
    doc.add_paragraph("")
    doc.add_paragraph(
        f"I am writing to request a formal amendment to my child's Education, Health and Care Plan, "
        f"specifically to Sections B and F, following observations made in the {environment} environment."
    )
    doc.add_paragraph("")
    doc.add_paragraph("The following factors have been identified through observation:")

    for item in confirmed_items:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(item)

    doc.add_paragraph("")
    doc.add_paragraph(
        "These observations correspond to needs already described in Section B of the current plan. "
        "The current Section F does not specify provision for this environment in sufficient detail "
        "to be enforceable. I am requesting that the following be added to Section F:"
    )

    additions = doc.add_paragraph(style="List Bullet")
    additions.add_run(f"Specific provision for the {environment} environment, including named adjustments")
    doc.add_paragraph(style="List Bullet").add_run("Named arrangements to be in place on every occasion my child uses this space")
    doc.add_paragraph(style="List Bullet").add_run("A half-termly review mechanism to confirm adjustments are in place")

    doc.add_paragraph("")
    doc.add_paragraph(
        "I would be grateful for written confirmation that this request has been received "
        "and will be considered at the earliest opportunity."
    )
    doc.add_paragraph("")
    doc.add_paragraph("Yours sincerely,")
    doc.add_paragraph("[Your name]")
    doc.add_paragraph(f"[Date: {today}]")

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def page_correspondence():
    """Enhanced correspondence intelligence — polished for tester release."""

    # ── Empty state / intro ───────────────────────────────────────────────────
    st.markdown("## Correspondence analysis")

    # Link back to report
    if st.session_state.get("findings"):
        st.markdown(
            f"<a href='#' onclick='void(0)' style='font-size:0.85rem;colour:{NAVY};'>← View my EHCP report</a>",
            unsafe_allow_html=True
        )
        if st.button("← Back to my EHCP report", key="back_to_report"):
            st.session_state.stage = "full_report"
            st.rerun()

    st.markdown(f"""
    <div style="background:#f0f4fa;border-radius:8px;padding:1.2rem 1.5rem;margin-bottom:1.5rem;">
      <p style="margin:0;font-size:0.97rem;line-height:1.7;">
        Upload a letter or email from school or the LA and FRED will read it for you —
        identifying patterns, root causes, and the right question to ask next.
        This is different from your EHCP report: the EHCP report analyses the document,
        correspondence analysis reads the relationship.
      </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Uploads ───────────────────────────────────────────────────────────────
    st.markdown("### Upload correspondence")

    input_method = st.radio(
        "How do you want to add your correspondence?",
        ["Upload file (PDF or Word)", "Paste text directly"],
        horizontal=True,
        key="input_method"
    )

    col1, col2 = st.columns(2)

    if input_method == "Upload file (PDF or Word)":
        with col1:
            email1 = st.file_uploader(
                "Most recent letter or email",
                type=["pdf", "docx", "txt"], key="email1"
            )
            email2 = st.file_uploader(
                "Earlier correspondence — optional",
                type=["pdf", "docx", "txt"], key="email2"
            )
        with col2:
            policy_file = st.file_uploader(
                "School policy — optional",
                type=["pdf", "docx", "txt"], key="policy_upload",
                help="Upload the school accessibility, SEN, or behaviour policy."
            )
            st.markdown(
                "<p style='font-size:0.82rem;colour:#666;margin-top:0.3rem;'>"
                "Upload the school's accessibility or SEN policy and FRED will check whether "
                "they have met their own commitments."
                "</p>", unsafe_allow_html=True
            )
        # Set paste vars to empty
        paste_text1 = ""
        paste_date1 = ""
        paste_text2 = ""
        paste_date2 = ""
    else:
        with col1:
            st.markdown("**Most recent correspondence**")
            paste_date1 = st.text_input(
                "Date of this correspondence",
                placeholder="e.g. 12 May 2025",
                key="paste_date1"
            )
            paste_text1 = st.text_area(
                "Paste the text here",
                placeholder="Paste the full text of the email or letter...",
                height=180,
                key="paste_text1"
            )
        with col2:
            st.markdown("**Earlier correspondence — optional**")
            paste_date2 = st.text_input(
                "Date of this correspondence",
                placeholder="e.g. 3 April 2025",
                key="paste_date2"
            )
            paste_text2 = st.text_area(
                "Paste the text here",
                placeholder="Paste the full text of an earlier email or letter...",
                height=180,
                key="paste_text2"
            )
        # Set file vars to None
        email1 = None
        email2 = None
        policy_file = None


    st.markdown("---")

    # ── Tone question ─────────────────────────────────────────────────────────
    # Tone will be auto-detected after analysis — store manual override option
    tone_override = st.session_state.get("tone_override", None)
    if tone_override:
        st.markdown(f"""
        <div style="background:#eaf5e0;border-radius:8px;padding:0.6rem 1rem;margin-bottom:0.5rem;">
          <p style="margin:0;font-size:0.85rem;colour:#2d5a2d;">
            Tone set to: <b>{tone_override}</b> —
            <a href="#" onclick="void(0)" style="colour:#5a8a5a;">change</a>
          </p>
        </div>
        """, unsafe_allow_html=True)
    tone_q = tone_override or "auto"

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Analyse button ────────────────────────────────────────────────────────
    analyse = st.button("Analyse correspondence", use_container_width=False, key="analyse_corr")

    if not email1 and not analyse:
        st.markdown(
            "<p style='colour:#888;font-size:0.9rem;margin-top:0.5rem;'>"
            "Upload at least one piece of correspondence to begin."
            "</p>",
            unsafe_allow_html=True
        )

    if analyse and not email1:
        st.error("Please upload at least one piece of correspondence to continue.")

    has_content = (email1 is not None) or (paste_text1.strip() != "")

    if analyse and not has_content:
        st.error("Please upload a file or paste some correspondence text to continue.")

    # Run analysis only when button clicked, then store in session state
    if analyse and has_content:
        st.session_state.corr_analysed = False  # reset to re-run
        
        # Progress steps
        progress_text = st.empty()
        progress_text.markdown("*Reading correspondence…*")

        if email1:
            text1 = extract_text(email1)
            if text1 == "__PDF_EXTRACTION_FAILED__":
                progress_text.empty()
                st.warning(
                    "FRED couldn't extract text from this PDF — it may be an image-based file. "
                    "Try using the 'Paste text directly' option instead, or save the email as a "
                    "Word document (.docx) and upload that."
                )
                st.stop()
        else:
            date_prefix = f"[Date: {paste_date1}]\n" if paste_date1 else ""
            text1 = date_prefix + paste_text1

        if email2:
            text2 = extract_text(email2)
            if text2 == "__PDF_EXTRACTION_FAILED__":
                text2 = ""
        elif paste_text2.strip():
            date_prefix2 = f"[Date: {paste_date2}]\n" if paste_date2 else ""
            text2 = date_prefix2 + paste_text2
        else:
            text2 = ""
        combined = text1 + "\n\n" + text2

        progress_text.markdown("*Identifying patterns…*")
        ehcp_text = st.session_state.get("vault", {}).get("ehcp", {}).get("text", "")
        matched_patterns = detect_patterns(combined, ehcp_text)

        progress_text.markdown("*Checking environment…*")
        environment = detect_environment(combined)

        policy_text = ""
        policy_findings = []
        if policy_file:
            progress_text.markdown("*Reading school policy…*")
            policy_text = extract_text_from_pdf(policy_file) if policy_file.name.endswith(".pdf") else extract_text_from_docx(policy_file)
            policy_findings = analyse_policy(policy_text)

        progress_text.markdown("*Detecting tone…*")
        tone_rec = detect_tone_recommendation(combined, matched_patterns)

        progress_text.markdown("*Checking correspondence history…*")
        similar_past = check_thread_for_similar(matched_patterns, environment)

        progress_text.empty()

        today = datetime.datetime.now().strftime("%d %B %Y")

        # Add to thread — deduplicate by content hash
        pattern_names = [p["name"] for p in matched_patterns[:3]]
        summary = f"Email from school — {len(matched_patterns)} pattern(s) detected"
        if environment:
            summary += f" — {environment} referenced"
        # Only add if this is genuinely new (not a resubmit)
        existing = st.session_state.get("thread", [])
        if not existing or existing[-1].get("summary") != summary:
            add_to_thread("from_school", summary, matched_patterns, tone_rec)

    # ── Display results from session state (persists across reruns) ──────────
    if st.session_state.get("corr_analysed"):
        matched_patterns = st.session_state.get("all_patterns", [])
        environment      = st.session_state.get("corr_environment")
        policy_findings  = st.session_state.get("corr_policy", [])
        tone_rec         = st.session_state.get("corr_tone_rec", {})
        similar_past     = st.session_state.get("corr_similar", [])
        draft_reply      = st.session_state.get("draft_reply", "")
        top_two          = st.session_state.get("top_two_patterns", [])
        today            = datetime.datetime.now().strftime("%d %B %Y")

        red_n, amber_n, summary_text, summary_colour = st.session_state.get(
            "corr_summary_bar", (0, 0, "Analysis complete.", "#2d4a2d")
        )
        red_n   = sum(1 for p in matched_patterns if p["tier"] == "red")
        amber_n = sum(1 for p in matched_patterns if p["tier"] == "amber")

        if red_n == 0 and amber_n == 0:
            summary_colour = GREEN
            summary_text = "No major patterns detected in this correspondence."
        elif red_n > 0:
            summary_colour = RED
            summary_text = f"{red_n} serious pattern{'s' if red_n > 1 else ''} and {amber_n} amber signal{'s' if amber_n != 1 else ''} detected."
        else:
            summary_colour = AMBER
            summary_text = f"{amber_n} pattern{'s' if amber_n > 1 else ''} detected. No immediate lawful concerns."

        st.markdown(f"""
        <div style="background:{summary_colour};border-radius:6px;padding:0.9rem 1.2rem;margin:1rem 0 1.5rem;">
          <p style="colour:white;font-weight:700;margin:0;font-size:1rem;">{summary_text}</p>
        </div>
        """, unsafe_allow_html=True)

        # ── Patterns ──────────────────────────────────────────────────────────
        st.markdown("### Patterns detected")

        if not matched_patterns:
            st.markdown(f"""
            <div class="finding-green">
              <span class="badge-green">No recognised patterns</span>
              <p style="margin:0.5rem 0 0;font-size:0.95rem;">
                No recognised school correspondence patterns were detected.
                This may mean the correspondence is straightforward — or that the language
                used does not match known patterns. Read the correspondence carefully
                before assuming it is without issue.
              </p>
            </div>
            """, unsafe_allow_html=True)
        else:
            # Show top 5, offer more
            show_all = st.session_state.get("show_all_patterns", False)
            display = matched_patterns if show_all else matched_patterns[:5]

            for i, p in enumerate(display):
                render_pattern_card(p, i)

            if len(matched_patterns) > 5 and not show_all:
                remaining = len(matched_patterns) - 5
                if st.button(f"Show {remaining} more pattern{'s' if remaining > 1 else ''}"):
                    st.session_state.show_all_patterns = True
                    st.rerun()
            elif show_all and len(matched_patterns) > 5:
                if st.button("Show fewer"):
                    st.session_state.show_all_patterns = False
                    st.rerun()

        # ── Environment / root cause ──────────────────────────────────────────
        if environment:
            st.markdown("---")
            st.markdown(f"### Root cause — {environment}")
            st.markdown(f"""
            <div style="background:#f0f4fa;border-radius:6px;padding:0.9rem 1.2rem;margin-bottom:1rem;">
              <p style="margin:0;font-size:0.95rem;">
                The correspondence references <b>{environment}</b> as a context for difficulty.
                Before responding, use the checklist below to identify what specifically within
                that environment may be the trigger. Tick everything that applies or has been observed.
                FRED will generate a root cause summary and — if needed — a draft amendment request.
              </p>
            </div>
            """, unsafe_allow_html=True)

            checklist = get_checklist(environment)
            confirmed_items = []

            for category, items in checklist:
                st.markdown(f"**{category}**")
                cols = st.columns(2)
                for idx, item in enumerate(items):
                    with cols[idx % 2]:
                        if st.checkbox(item, key=f"chk_{category}_{idx}"):
                            confirmed_items.append(f"{category}: {item}")

            if confirmed_items:
                st.markdown("---")
                st.markdown("### Root cause confirmed")

                st.markdown(f"""
                <div class="finding-red">
                  <span class="badge-red">{len(confirmed_items)} factor{'s' if len(confirmed_items) > 1 else ''} confirmed — {environment}</span>
                  <p style="margin:0.5rem 0 0;font-size:0.95rem;">
                    These factors should be documented and cross-referenced against Section B and
                    Section F of the EHCP. If confirmed factors are not addressed in the current
                    provision, this is evidence for an EHCP amendment.
                  </p>
                </div>
                """, unsafe_allow_html=True)

                # Immediate Do
                st.markdown("**Immediate steps — check what is already in the EHCP:**")
                st.markdown(
                    "<p style='font-size:0.9rem;colour:#555;margin-bottom:0.8rem;'>"
                    "Before requesting an amendment, check whether these provisions — "
                    "which may already be in Section F — apply specifically to this environment. "
                    "If they do, ask the school in writing to confirm they are in place."
                    "</p>",
                    unsafe_allow_html=True
                )
                existing = [
                    "Ear defenders — access to sensory toolkit",
                    "Entry to the space before other pupils — to reduce crowding",
                    "Named trusted adult present during this activity",
                    "Visual schedule — advance notice of what will happen",
                    "Safe exit route — agreed way to leave if overwhelmed",
                    "Snack provision before the activity — if hunger is a factor",
                    "Quiet alternative — access to a calm space as an option",
                ]
                for prov in existing:
                    st.checkbox(prov, key=f"existing_{prov[:25]}")

                # Amendment request — behind a button
                st.markdown("---")
                if st.button("Generate EHCP amendment request", key="gen_amendment"):
                    st.session_state.show_amendment = True

                if st.session_state.get("show_amendment"):
                    st.markdown(f"""
                    <div style="background:{NAVY};border-radius:8px;padding:1.2rem 1.5rem;margin:1rem 0;">
                      <p style="colour:#a8b8d8;font-size:0.8rem;margin:0 0 0.3rem;letter-spacing:0.08em;text-transform:uppercase;">EHCP AMENDMENT FLAG</p>
                      <p style="colour:white;font-weight:700;margin:0 0 0.4rem;font-size:1rem;">New evidence — {environment} — {today}</p>
                      <p style="colour:#c8d8f0;margin:0;font-size:0.9rem;">
                        The observations above constitute new evidence about your child's sensory needs
                        in this environment. This should be formally documented and used to inform
                        the next EHCP review. Edit the draft below before sending.
                      </p>
                    </div>
                    """, unsafe_allow_html=True)

                    draft = f"""Dear [SENCO name],

I am writing to request a formal amendment to [child's name]'s Education, Health and Care Plan — specifically to Sections B and F — following observations in the {environment} on {today}.

The following has been identified:

{chr(10).join(f"- {item}" for item in confirmed_items)}

The current Section F does not specify provision for this environment in sufficient detail to be enforceable. I am requesting:

- Specific named provision for the {environment} environment
- Named arrangements confirmed in writing, in place on every occasion
- A half-termly review to confirm adjustments remain in place

Please confirm receipt of this request and advise when it will be considered.

Yours sincerely,
[Your name]
{today}"""

                    st.text_area("Draft amendment request — edit before sending:", value=draft, height=280, key="amendment_text")

                    # Word download
                    word_buf = generate_amendment_word(environment, confirmed_items, today)
                    st.download_button(
                        "Download as Word document",
                        data=word_buf,
                        file_name=f"FRED_amendment_request_{today.replace(' ','_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

                    # Log to knowledge bank
                    if "knowledge_bank" not in st.session_state:
                        st.session_state.knowledge_bank = []
                    entry = {
                        "date": today,
                        "environment": environment,
                        "confirmed": confirmed_items,
                        "amendment_flagged": True,
                    }
                    if entry not in st.session_state.knowledge_bank:
                        st.session_state.knowledge_bank.append(entry)
                    print(f"FRED AMENDMENT FLAG: {today} — {environment} — {len(confirmed_items)} factors")

        # ── Policy findings ───────────────────────────────────────────────────
        if policy_findings:
            st.markdown("---")
            st.markdown("### School policy cross-reference")
            st.markdown(
                "<p style='font-size:0.95rem;margin-bottom:1rem;'>"
                "FRED has read the uploaded school policy and found the following in relation "
                "to the correspondence. The school cannot dispute its own policy."
                "</p>",
                unsafe_allow_html=True
            )
            for f in policy_findings:
                tier = f["tier"]
                st.markdown(f"""
                <div class="finding-{tier}">
                  <span class="badge-{tier}">POLICY</span>
                  <p style="font-weight:700;margin:0.4rem 0 0.3rem;">{f['title']}</p>
                  {f'<p style="font-style:italic;colour:#555;font-size:0.88rem;margin:0 0 0.4rem;">"{f["extract"]}"</p>' if f.get("extract") else ""}
                  <p style="margin:0;font-size:0.93rem;">{f['commentary']}</p>
                </div>
                """, unsafe_allow_html=True)

        # ── Tone recommendation — auto-detected ──────────────────────────────
        st.markdown("---")
        st.markdown("### Tone recommendation")

        rec_colour = {"formal": "#C0392B", "collaborative": "#1E8449", "neutral": "#D4A017"}
        rec_bg    = {"formal": "#fdf4f3", "collaborative": "#f3faf5", "neutral": "#fdf9f0"}
        r = tone_rec["recommendation"]

        st.markdown(f"""
        <div style="border-left:4px solid {rec_colour.get(r,'#888')};
                    background:{rec_bg.get(r,'#f9f9f9')};
                    border-radius:0 8px 8px 0;padding:1rem 1.2rem;margin-bottom:1rem;">
          <p style="font-weight:600;margin:0 0 0.3rem;font-size:0.95rem;">
            FRED recommends: {tone_rec['label']}
          </p>
          <p style="margin:0;font-size:0.9rem;line-height:1.6;colour:#444;">
            {tone_rec['reasoning']}
          </p>
          <p style="margin:0.6rem 0 0;font-size:0.78rem;colour:#888;">
            Confidence: {tone_rec['confidence'].title()} —
            based on language patterns in this correspondence.
            This is a recommendation, not an instruction.
          </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("**Does this match your experience of the relationship?**")
        tone_confirm = st.radio(
            "",
            ["Yes — use recommended tone",
             "No — it should be more formal",
             "No — it should be more collaborative",
             "I'll decide later"],
            key="tone_confirm",
            label_visibility="collapsed",
            horizontal=True,
        )
        if "formal" in tone_confirm:
            st.session_state.tone_override = "formal"
        elif "collaborative" in tone_confirm:
            st.session_state.tone_override = "collaborative"
        elif "Yes" in tone_confirm:
            st.session_state.tone_override = r

        # ── Draft reply ───────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### Draft reply")

        st.markdown("""
        <p style='font-size:0.92rem;colour:#444;margin-bottom:0.3rem;'>
          This draft email has been generated for you. Please check to ensure it is your voice before sending.
        </p>
        <p style='font-size:0.85rem;colour:#7ab870;margin-bottom:1rem;font-style:italic;'>
          A focused question in writing is harder to ignore than five.
        </p>
        """, unsafe_allow_html=True)

        # ── Build prioritised two-question draft ──────────────────────────────
        red_patterns   = [p for p in matched_patterns if p["tier"] == "red"]
        amber_patterns = [p for p in matched_patterns if p["tier"] == "amber"]
        tone_label     = st.session_state.get("tone_override", tone_rec.get("recommendation", "neutral"))

        # Priority order for which patterns generate a question in the draft
        # Everything else goes to meeting companion only
        PRIORITY_IDS = [
            "behaviour_framing",
            "unstructured_time_gap",
            "ehcp_xref_lunch",
            "ehcp_xref_lunch_gap",
            "veiled_threat",
            "reintegration_promise",
            "staffing_change",
            "legal_misrepresentation",
            "resources_defence",
            "monitoring_without_action",
            "home_responsibility_redirect",
            "reassurance_without_evidence",
            "implicit_admission",
        ]

        QUESTION_PARAGRAPHS = {
            "behaviour_framing": (
                "I would be grateful if you could confirm in writing what provision from "
                "Section F of [child's name]'s EHCP was specifically in place at the time of "
                "this incident, and provide the delivery log for that period. "
                "I ask because the incident occurred during a time when the EHCP specifies "
                "support should be in place."
            ),
            "unstructured_time_gap": (
                "Could you confirm in writing what support was in place during lunchtime on "
                "the day of this incident, and share the relevant delivery records with me?"
            ),
            "ehcp_xref_lunch": (
                "The EHCP specifies provision during unstructured periods including lunchtime. "
                "I would ask you to confirm whether that provision was being delivered at the "
                "time of this incident, and to provide the delivery log."
            ),
            "ehcp_xref_lunch_gap": (
                "I note that the EHCP does not currently specify provision for lunchtime. "
                "Given that this incident occurred during that period, I would like to raise "
                "this at the next annual review as a gap that needs to be addressed."
            ),
            "veiled_threat": (
                "I would ask you to clarify in writing what 'next steps in terms of support "
                "and provision' means specifically. Any change to provision or placement must "
                "go through the EHCP review process under the Children and Families Act 2014."
            ),
            "reintegration_promise": (
                "Could you provide the written reintegration plan, specifically referencing "
                "the provision in Section F of the EHCP? I would ask for this within "
                "five working days."
            ),
            "staffing_change": (
                "Could you confirm whether there have been any changes to [child's name]'s "
                "1:1 support since the last annual review, and what transition planning "
                "was put in place for any such changes?"
            ),
            "legal_misrepresentation": (
                "I would ask the school to confirm delivery of each provision specified in "
                "Section F of the EHCP. The relevant obligation under the Children and "
                "Families Act 2014 is to deliver what the EHCP specifies."
            ),
            "resources_defence": (
                "Could you identify in writing which provisions in Section F are currently "
                "being delivered and whether any are not being delivered due to resource "
                "constraints? Any such gaps become a matter for the LA to address."
            ),
            "monitoring_without_action": (
                "Could you clarify what specifically will be monitored, what the review "
                "mechanism is, who is responsible, and what action will be taken if the "
                "monitoring identifies further difficulty? I would ask for this in writing."
            ),
            "home_responsibility_redirect": (
                "I would like to focus on what provision was in place at school at the time "
                "of the incident, as this is governed by the EHCP. Could you confirm this "
                "in writing?"
            ),
            "reassurance_without_evidence": (
                "Could you share the evidence base for the assurances given in your email, "
                "including the delivery log for the relevant provision over the past half term?"
            ),
            "implicit_admission": (
                "Could you confirm in writing what the current position is — specifically "
                "what is in place now, rather than what will be put in place going forward?"
            ),
        }

        # Select top two priority patterns that have a question
        priority_sorted = sorted(
            [p for p in matched_patterns if p["id"] in PRIORITY_IDS],
            key=lambda x: PRIORITY_IDS.index(x["id"]) if x["id"] in PRIORITY_IDS else 99
        )
        top_two = priority_sorted[:2]
        reserve  = matched_patterns[2:]  # everything else goes to meeting companion

        # Build the letter
        openings = {
            "formal": "Dear [Name],\n\nThank you for your email dated [date]. I am writing to follow up on the points raised.",
            "collaborative": "Dear [Name],\n\nThank you for letting us know about the incident on [date]. I wanted to follow up on a couple of points.",
            "neutral": "Dear [Name],\n\nThank you for your email dated [date]. I would like to raise a couple of points in response.",
        }
        opening = openings.get(tone_label, openings["neutral"])

        # Intro line before questions
        intro = (
            "I have a couple of specific questions I would be grateful if you could "
            "address in writing."
        )

        # Numbered questions
        question_lines = []
        for i, p in enumerate(top_two):
            q = QUESTION_PARAGRAPHS.get(p["id"], "")
            if q:
                question_lines.append(f"{i+1}. {q}")

        closings = {
            "formal": "I would be grateful for a written response within five working days.\n\nYours sincerely,\n[Your name]",
            "collaborative": "I look forward to hearing from you.\n\nKind regards,\n[Your name]",
            "neutral": "I look forward to your response.\n\nYours sincerely,\n[Your name]",
        }
        closing = closings.get(tone_label, closings["neutral"])

        # Build full letter with opening
        if question_lines:
            draft_parts = [opening, intro] + question_lines + [closing]
        else:
            draft_parts = [opening, closing]
        draft_reply = "\n\n".join(p for p in draft_parts if p.strip())

        # Store draft and reserve patterns in session state for meeting companion
        st.session_state["draft_reply"]        = draft_reply
        st.session_state["reserve_patterns"]   = reserve
        st.session_state["top_two_patterns"]   = top_two
        st.session_state["all_patterns"]       = matched_patterns
        st.session_state["correspondence_date"]= paste_date1 if not email1 else "[date]"
        st.session_state["corr_summary_bar"]   = (red_n, amber_n, summary_text, summary_colour)
        st.session_state["corr_environment"]   = environment
        st.session_state["corr_policy"]        = policy_findings
        st.session_state["corr_tone_rec"]      = tone_rec
        st.session_state["corr_similar"]       = similar_past
        st.session_state["corr_analysed"]      = True
        st.rerun()
def page_subscriber():
    """
    The subscriber environment — feels like a different product.
    Beta: all features accessible, no payment required.
    """
    name = st.session_state.email_address if st.session_state.email_address else "there"

    st.markdown(f"""
    <div style="background:{NAVY};border-radius:8px;padding:2rem;margin-bottom:1.5rem;">
      <p style="colour:#a8b8d8;font-size:0.85rem;margin:0 0 0.3rem;letter-spacing:0.1em;text-transform:uppercase;">
        FRED BETA — SUBSCRIBER WORKSPACE
      </p>
      <h2 style="colour:white;margin:0 0 0.4rem;font-family:'Playfair Display',serif;">
        Welcome to FRED
      </h2>
      <p style="colour:#c8d8f0;margin:0;font-size:0.95rem;">
        Everything is unlocked. You have full access during beta.
      </p>
    </div>
    """, unsafe_allow_html=True)

    # Workspace tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📄 My Report", "✉️ Correspondence", "📋 Meeting Prep", "🗂️ Evidence Bank"])

    with tab1:
        st.markdown("### Your EHCP report")
        if st.session_state.findings:
            st.markdown("Your report is ready. Use the buttons below to download it.")
            col1, col2 = st.columns(2)
            with col1:
                word_buf = generate_word_report(
                    st.session_state.findings,
                    doc_type=st.session_state.doc_type,
                    situation=st.session_state.situation
                )
                st.download_button(
                    "Download as Word",
                    data=word_buf,
                    file_name="FRED_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            with col2:
                pdf_buf = generate_pdf_report(
                    st.session_state.findings,
                    doc_type=st.session_state.doc_type,
                    situation=st.session_state.situation
                )
                st.download_button(
                    "Download as PDF",
                    data=pdf_buf,
                    file_name="FRED_report.pdf",
                    mime="application/pdf",
                )
            st.markdown("---")
            st.markdown("### All findings")
            for finding in st.session_state.findings:
                render_finding_card(finding, show_full=True)
        else:
            st.info("No report yet. Go back to the home page and upload your EHCP.")
            if st.button("Upload an EHCP"):
                st.session_state.stage = "upload"
                st.rerun()

    with tab2:
        st.markdown("### Correspondence analysis")
        st.markdown(
            "Upload emails from school or the LA. FRED will read them for you — "
            "tone, intent, gaps, and what to do next."
        )
        email1 = st.file_uploader("Most recent email (PDF or Word)", type=["pdf", "docx", "txt"], key="sub_email1")
        email2 = st.file_uploader("Previous email — optional", type=["pdf", "docx", "txt"], key="sub_email2")

        tone_q = st.radio(
            "How is the school engaging right now?",
            ["Genuinely trying to help", "Going through the motions", "Actively avoiding or obstructing"],
            help="You know the relationship. FRED doesn't. Your answer calibrates the analysis."
        )

        if st.button("Analyse correspondence") and email1:
            with st.spinner("Reading correspondence…"):
                text1 = extract_text_from_pdf(email1) if email1.name.endswith(".pdf") else extract_text_from_docx(email1)
                text2 = ""
                if email2:
                    text2 = extract_text_from_pdf(email2) if email2.name.endswith(".pdf") else extract_text_from_docx(email2)
                combined = text1 + "\n\n" + text2
                matched = detect_patterns(combined)
                tone_rec = detect_tone_recommendation(combined, matched)

            st.markdown("### Patterns detected")
            if matched:
                for p in matched[:5]:
                    st.markdown(f"""
                    <div class="finding-{p['tier']}">
                      <span class="badge-{p['tier']}">{p['name']}</span>
                      <p style="margin:0.4rem 0 0.3rem;font-size:0.93rem;">{p['explanation']}</p>
                      <p style="font-style:italic;font-size:0.88rem;colour:#555;margin:0;">{p['fred_question']}</p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("No major patterns detected in this correspondence.")

            st.markdown(f"""
            <div style="background:#eaf5e0;border-radius:8px;padding:0.8rem 1.1rem;margin-top:1rem;">
              <p style="font-weight:600;margin:0 0 0.2rem;font-size:0.9rem;">Tone recommendation: {tone_rec['label']}</p>
              <p style="margin:0;font-size:0.85rem;colour:#555;">{tone_rec['reasoning']}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")
            st.info(
                "Hold is a valid strategic choice. "
                "Sometimes the most powerful response is a short warm acknowledgement "
                "that signals knowledge without displaying it."
            )

    with tab3:
        st.markdown("### Meeting preparation")
        st.markdown(
            "Use this before your annual review or any meeting with school or the LA."
        )

        upcoming = st.session_state.get("upcoming_dates", "")
        if upcoming:
            st.info(f"Upcoming date noted: {upcoming}")

        st.markdown("**Before the meeting — check these:**")
        checklist = [
            "Request the agenda in writing at least five working days before",
            "Ask for the delivery log for each provision in Section F",
            "Confirm who will be attending and in what capacity",
            "Take a copy of the EHCP and this report",
            "Note the date, time, and attendees at the start of the meeting",
            "Request that minutes are circulated within five working days",
            "Follow up any verbal agreements in writing within 24 hours",
        ]
        for item in checklist:
            st.checkbox(item, key=f"check_{item[:20]}")

        st.markdown("---")
        st.markdown("**After the meeting:**")
        post = st.text_area(
            "Note anything agreed verbally that needs to be confirmed in writing",
            placeholder="e.g. SENCO agreed to provide weekly session records from next term...",
            height=100,
        )
        if post:
            st.success(
                "Noted. Follow this up in writing within 24 hours. "
                "Subject line: 'Confirmation of agreements — [date] meeting'. "
                "Copy in the SENCO and the LA caseworker if present."
            )

    with tab4:
        st.markdown("### Evidence bank")
        st.markdown(
            "FRED records confirmed findings from correspondence analysis here, "
            "dated and stored for the duration of your session. "
            "Use these at annual review."
        )
        bank = st.session_state.get("knowledge_bank", [])
        held = st.session_state.get("held_findings", [])
        if bank:
            for entry in bank:
                confirmed_list = entry.get("confirmed", [])
                items_html = "".join(f'<p style="font-size:0.85rem;margin:0.1rem 0;">• {item}</p>' for item in confirmed_list)
                amendment_html = "<p style='font-size:0.82rem;font-weight:700;colour:#C0392B;margin-top:0.4rem;'>⚑ EHCP amendment flagged</p>" if entry.get("amendment_flagged") else ""
                st.markdown(f"""
                <div style="background:white;border:1px solid #d0dae8;border-radius:6px;padding:1rem 1.2rem;margin-bottom:0.8rem;">
                  <p style="font-weight:700;margin:0 0 0.2rem;">{entry['date']} — {entry['environment']}</p>
                  <p style="margin:0 0 0.4rem;font-size:0.88rem;colour:#555;">{len(confirmed_list)} confirmed factor(s)</p>
                  {items_html}
                  {amendment_html}
                </div>
                """, unsafe_allow_html=True)
        elif held:
            for entry in held:
                st.markdown(f"""
                <div style="background:white;border:1px solid #d0dae8;border-radius:6px;padding:1rem 1.2rem;margin-bottom:0.8rem;">
                  <p style="font-weight:700;margin:0 0 0.2rem;">{entry['date']}</p>
                  <p style="margin:0;font-size:0.88rem;colour:#555;">{entry.get('patterns',0)} pattern(s) — held without action</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown(
                "<p style='colour:#888;font-size:0.9rem;'>"
                "No evidence entries yet. Confirmed root cause findings from "
                "correspondence analysis will appear here with dates."
                "</p>",
                unsafe_allow_html=True
            )
        st.markdown(
            "<p style='font-size:0.82rem;colour:#aaa;margin-top:1rem;'>"
            "Evidence bank is stored for this session. "
            "Screenshot entries before closing your browser."
            "</p>",
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown(f"""
    <a href="https://docs.google.com/forms/d/e/1FAIpQLSeA1F9nEdQWkmplbAh973XKq2EsW0bEkhJiw7drhP7BZaPjKQ/viewform" target="_blank" style="
        display:inline-block;
        background:#e8eef8;
        colour:{NAVY};
        padding:0.6rem 1.4rem;
        border-radius:4px;
        text-decoration:none;
        font-family:'Source Sans 3',sans-serif;
        font-weight:600;
        font-size:0.9rem;
    ">Leave feedback on FRED →</a>
    """, unsafe_allow_html=True)



# ── NAVIGATION ────────────────────────────────────────────────────────────────

def render_nav():
    col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 3])
    with col1:
        if st.button("Home"):
            st.session_state.stage = "landing"
            st.rerun()
    with col2:
        if st.button("My report"):
            if st.session_state.findings:
                st.session_state.stage = "full_report"
                st.rerun()
    with col3:
        if st.session_state.subscribed:
            if st.button("My workspace"):
                st.session_state.stage = "subscriber"
                st.rerun()
    with col4:
        if st.button("Correspondence"):
            st.session_state.stage = "correspondence"
            st.rerun()


# ── ROUTER ────────────────────────────────────────────────────────────────────

render_nav()
st.markdown("<hr style='margin:0 0 1.5rem;border-colour:#d0dae8;'>", unsafe_allow_html=True)

stage = st.session_state.stage

if stage == "landing":
    page_landing()
elif stage == "explainer":
    page_explainer()
elif stage == "upload":
    page_upload()
elif stage == "sneak_peek":
    page_sneak_peek()
elif stage == "full_report":
    page_full_report()
elif stage == "subscriber":
    page_subscriber()
elif stage == "correspondence":
    page_correspondence()
else:
    page_landing()
