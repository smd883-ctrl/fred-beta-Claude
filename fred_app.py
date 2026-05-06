import streamlit as st
import fitz  # PyMuPDF
import re
import json
import requests
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

# ── CUSTOM CSS ───────────────────────────────────────────────────────────────

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&family=Source+Sans+3:wght@400;600&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Source Sans 3', sans-serif;
    background-color: {LIGHT};
    color: #1a1a2e;
  }}

  h1, h2, h3 {{
    font-family: 'Playfair Display', serif;
  }}

  .hero {{
    background: {NAVY};
    color: white;
    padding: 3rem 2rem;
    border-radius: 8px;
    text-align: center;
    margin-bottom: 2rem;
  }}

  .hero h1 {{
    font-size: 4rem;
    margin: 0;
    color: white;
    letter-spacing: 0.12em;
  }}

  .hero .full-name {{
    font-family: 'Source Sans 3', sans-serif;
    font-size: 0.9rem;
    letter-spacing: 0.25em;
    text-transform: uppercase;
    color: #a8b8d8;
    margin: 0.3rem 0 0.8rem 0;
  }}

  .hero .tagline {{
    font-size: 1.4rem;
    font-weight: 600;
    color: white;
    margin-bottom: 0.5rem;
  }}

  .hero .origin {{
    font-style: italic;
    color: #a8b8d8;
    margin-bottom: 1.5rem;
    font-size: 1rem;
  }}

  .hero .sub {{
    color: #c8d8f0;
    max-width: 560px;
    margin: 0 auto 1.5rem auto;
    font-size: 1.05rem;
    line-height: 1.6;
  }}

  .hero .service-line {{
    color: #a8b8d8;
    font-size: 0.95rem;
    margin-bottom: 1.8rem;
  }}

  .cta-button {{
    display: inline-block;
    background: white;
    color: {NAVY} !important;
    font-family: 'Source Sans 3', sans-serif;
    font-weight: 700;
    font-size: 1.1rem;
    padding: 0.85rem 2.4rem;
    border-radius: 4px;
    text-decoration: none;
    letter-spacing: 0.04em;
    cursor: pointer;
    border: none;
  }}

  .reassurance {{
    color: #a8b8d8;
    font-size: 0.9rem;
    margin-top: 1rem;
  }}

  .pricing-hint {{
    color: #7a8fa8;
    font-size: 0.85rem;
    margin-top: 0.4rem;
  }}

  /* Traffic light badges */
  .badge-red {{
    background: {RED};
    color: white;
    padding: 0.25rem 0.75rem;
    border-radius: 3px;
    font-weight: 700;
    font-size: 0.85rem;
    display: inline-block;
    margin-bottom: 0.5rem;
  }}
  .badge-amber {{
    background: {AMBER};
    color: white;
    padding: 0.25rem 0.75rem;
    border-radius: 3px;
    font-weight: 700;
    font-size: 0.85rem;
    display: inline-block;
    margin-bottom: 0.5rem;
  }}
  .badge-green {{
    background: {GREEN};
    color: white;
    padding: 0.25rem 0.75rem;
    border-radius: 3px;
    font-weight: 700;
    font-size: 0.85rem;
    display: inline-block;
    margin-bottom: 0.5rem;
  }}

  /* Finding cards */
  .finding-red {{
    border-left: 5px solid {RED};
    background: #fdf4f3;
    padding: 1rem 1.2rem;
    border-radius: 0 6px 6px 0;
    margin-bottom: 1rem;
  }}
  .finding-amber {{
    border-left: 5px solid {AMBER};
    background: #fdf9f0;
    padding: 1rem 1.2rem;
    border-radius: 0 6px 6px 0;
    margin-bottom: 1rem;
  }}
  .finding-green {{
    border-left: 5px solid {GREEN};
    background: #f3faf5;
    padding: 1rem 1.2rem;
    border-radius: 0 6px 6px 0;
    margin-bottom: 1rem;
  }}

  /* How it works bullets */
  .hiw-item {{
    display: flex;
    align-items: flex-start;
    gap: 1rem;
    margin-bottom: 1rem;
    padding: 0.8rem 1rem;
    background: white;
    border-radius: 6px;
    border: 1px solid #e0e8f0;
  }}
  .hiw-dot {{
    width: 10px;
    height: 10px;
    background: {NAVY};
    border-radius: 50%;
    margin-top: 6px;
    flex-shrink: 0;
  }}

  /* Pricing cards */
  .pricing-card {{
    background: white;
    border: 1px solid #d0dae8;
    border-radius: 8px;
    padding: 1.5rem;
    text-align: center;
    height: 100%;
  }}
  .pricing-card.featured {{
    border: 2px solid {NAVY};
    position: relative;
  }}
  .pricing-card .price {{
    font-size: 2rem;
    font-weight: 700;
    color: {NAVY};
  }}
  .pricing-card .period {{
    font-size: 0.85rem;
    color: #666;
  }}
  .pricing-card ul {{
    text-align: left;
    padding-left: 1rem;
    margin-top: 1rem;
    font-size: 0.9rem;
    line-height: 1.8;
  }}

  .best-value-tag {{
    background: {NAVY};
    color: white;
    font-size: 0.75rem;
    font-weight: 700;
    padding: 0.2rem 0.6rem;
    border-radius: 2px;
    display: inline-block;
    margin-bottom: 0.5rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }}

  /* Sneak peek box */
  .sneak-box {{
    border: 2px dashed {NAVY};
    border-radius: 8px;
    padding: 1.5rem;
    background: white;
    margin: 1.5rem 0;
  }}

  /* Section divider */
  .section-divider {{
    border: none;
    border-top: 1px solid #d0dae8;
    margin: 2.5rem 0;
  }}

  /* Suppress Streamlit defaults */
  .stButton > button {{
    background: {NAVY};
    color: white;
    border: none;
    border-radius: 4px;
    font-family: 'Source Sans 3', sans-serif;
    font-weight: 600;
    padding: 0.6rem 1.6rem;
    font-size: 1rem;
  }}
  .stButton > button:hover {{
    background: #253560;
    color: white;
  }}

  footer {{visibility: hidden;}}
  #MainMenu {{visibility: hidden;}}
</style>
""", unsafe_allow_html=True)

# ── EMAILJS ──────────────────────────────────────────────────────────────────

def send_emailjs(email, source, extra_dict=None):
    """
    Fire EmailJS using variable names that match template_44rbysv exactly.
    Template vars: {{name}}, {{email}}, {{user_email}}, {{message}},
                   {{timestamp}}, {{q1}} through {{q9}}
    """
    import datetime
    if extra_dict is None:
        extra_dict = {}
    try:
        service_id  = st.secrets.get("EMAILJS_SERVICE_ID", "service_8fbqfzn")
        template_id = st.secrets.get("EMAILJS_TEMPLATE_ID", "template_44rbysv")
        public_key  = st.secrets.get("EMAILJS_PUBLIC_KEY",  "ieFEY10YArTGltwrf")
        timestamp = datetime.datetime.now().strftime("%d %b %Y %H:%M")
        params = {
            "name":       email,
            "email":      email,
            "user_email": email,
            "message":    f"Source: {source}",
            "timestamp":  timestamp,
            "q1": extra_dict.get("new_info", ""),
            "q2": extra_dict.get("tl_clear", ""),
            "q3": extra_dict.get("layout", ""),
            "q4": extra_dict.get("would_pay", ""),
            "q5": extra_dict.get("fair_price", ""),
            "q6": extra_dict.get("subscribe", ""),
            "q7": extra_dict.get("personalise", ""),
            "q8": extra_dict.get("other", ""),
            "q9": extra_dict.get("notify", ""),
        }
        payload = {
            "service_id":      service_id,
            "template_id":     template_id,
            "user_id":         public_key,
            "template_params": params,
        }
        response = requests.post(
            "https://api.emailjs.com/api/v1.0/email/send",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=8,
        )
        print(f"EmailJS [{source}]: {response.status_code} — {response.text}")
    except Exception as e:
        print(f"EmailJS error ({source}): {e}")


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
    """Extract raw text from every page of uploaded PDF."""
    try:
        data = uploaded_file.read()
        doc  = fitz.open(stream=data, filetype="pdf")
        pages = []
        for page in doc:
            pages.append(page.get_text("text"))
        return "\n".join(pages)
    except Exception as e:
        return ""

def extract_text_from_docx(uploaded_file):
    """Extract raw text from uploaded Word document."""
    try:
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
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
                "(e.g. Speech and Language Therapist or trained TA working to a SALT programme), "
                "and the specific strategies to be used. Generalised statements such as "
                "'communication support will be provided' are not sufficient."
            ),
        },
        {
            "name": "Cognition and learning",
            "b_keywords": r'\b(cognition|learning|literacy|numeracy|reading|writing|dyslexia|processing|memory|attention)\b',
            "f_keywords": r'\b(literacy|numeracy|reading|writing|learning support|intervention|programme)\b',
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
                "(e.g. Occupational Therapist or TA trained to deliver OT programme), "
                "and the specific programme or strategies. Environmental adjustments should "
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
    tier_colours = {"red": RGBColor(0xC0, 0x39, 0x2B),
                    "amber": RGBColor(0xD4, 0xA0, 0x17),
                    "green": RGBColor(0x1E, 0x84, 0x49)}

    for finding in findings:
        tier = finding["tier"]
        h = doc.add_heading(tier_labels.get(tier, tier.upper()), 2)
        h.runs[0].font.color.rgb = tier_colours.get(tier, RGBColor(0, 0, 0))

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
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1a2744'), spaceAfter=8))
    story.append(Paragraph("EHCP Analysis Report", meta_style))
    story.append(Paragraph(f"Document type: {doc_type}", meta_style))
    if situation:
        story.append(Paragraph(f"Context: {situation}", meta_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#d0dae8'), spaceAfter=16))
    story.append(Spacer(1, 0.5*cm))

    # ── Summary ───────────────────────────────────────────────────────────
    story.append(Paragraph("Summary", section_h_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#d0dae8'), spaceAfter=10))

    red_n   = sum(1 for f in findings if f["tier"] == "red")
    amber_n = sum(1 for f in findings if f["tier"] == "amber")
    green_n = sum(1 for f in findings if f["tier"] == "green")

    story.append(Paragraph(
        f'<font color="#C0392B"><b>{red_n} Red finding{"s" if red_n != 1 else ""}</b></font> — '
        f'lawful requirement{"s" if red_n != 1 else ""} not met. Must be addressed at annual review.',
        body_style
    ))
    story.append(Paragraph(
        f'<font color="#D4A017"><b>{amber_n} Amber finding{"s" if amber_n != 1 else ""}</b></font> — '
        f'best practice gap{"s" if amber_n != 1 else ""}. Worth raising at annual review.',
        body_style
    ))
    story.append(Paragraph(
        f'<font color="#1E8449"><b>{green_n} Green finding{"s" if green_n != 1 else ""}</b></font> — '
        f'compliant. Use as benchmark when challenging non-compliant provision.',
        body_style
    ))

    story.append(Spacer(1, 0.4*cm))

    # Delivery log alert
    needs_log = any(f.get("delivery_log_required") for f in findings)
    if needs_log:
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#D4A017'), spaceAfter=6))
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
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#D4A017'), spaceAfter=10))

    story.append(Spacer(1, 0.6*cm))

    # ── Findings ──────────────────────────────────────────────────────────
    story.append(Paragraph("Findings", section_h_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#d0dae8'), spaceAfter=12))

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
        story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor('#e0e8f0'), spaceAfter=8))

    story.append(Spacer(1, 0.8*cm))

    # ── What next ─────────────────────────────────────────────────────────
    story.append(Paragraph("What next?", section_h_style))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#d0dae8'), spaceAfter=12))

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
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1a2744'), spaceAfter=8))
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
        "survey_submitted": False,
        "show_full_report": False,
        "relationship_tone": "neutral",
        "upcoming_dates": "",
        "doc_name": "",
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
        html += f'<p style="font-style:italic;color:#555;font-size:0.9rem;margin:0 0 0.5rem 0;">"{finding["extract"][:280]}…"</p>'

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
      <p class="tagline">Know what your child is entitled to. Know when it's not being delivered.</p>
      <p class="origin">Built by a parent who learned the hard way — so you don't have to.</p>
      <p class="sub">
        FRED analyses your child's EHCP against the Children and Families Act 2014 and the SEND Code of Practice.
        You get a plain-English report that tells you exactly what's lawfully required, what's missing, and what to do next.
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
    <p style="text-align:center;color:#6a7a90;font-size:0.9rem;margin-top:0.8rem;">
      Upload first. Decide after. Your report is ready before you pay.
    </p>
    <p style="text-align:center;color:#8a9ab0;font-size:0.82rem;">
      From £XX for the full report — or see our subscription plans below
    </p>
    """, unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # How it works
    st.markdown("<h2 style='text-align:center;'>Everything you need to know.</h2>", unsafe_allow_html=True)
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
    st.markdown("<h2 style='text-align:center;'>Plans</h2>", unsafe_allow_html=True)
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

    st.markdown("<br><p style='text-align:center;color:#6a7a90;font-size:0.85rem;'>No hidden charges. Your report is ready before you purchase.</p>", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # FAQ
    st.markdown("<h2 style='text-align:center;'>FAQ</h2>", unsafe_allow_html=True)

    faqs = [
        ("When do I pay?", "After you see one real finding from your EHCP, before you access the full report. You upload first — your report is ready before you pay."),
        ("Is this legal advice?", "No. FRED provides lawful analysis — it tells you what the law requires. It does not tell you what to do about it. That decision is yours."),
        ("What documents do I need?", "Your child's EHCP is essential. Everything else — EP reports, OT reports, correspondence — adds to the analysis but is not required to start."),
        ("What is an EHCP?", "An Education, Health and Care Plan is a legally binding document setting out a child's needs and the provision that must be made. The word 'must' in that sentence is important."),
        ("What if I only have a draft EHCP?", "FRED analyses draft and final EHCPs differently. For a draft, FRED highlights what should be strengthened before the document is finalised. For a final EHCP, FRED references your rights and the duty to deliver."),
        ("Is my data private?", "Yes. Your documents are used only for your analysis. They are not shared, stored beyond your session, or used to train any system."),
    ]

    for q, a in faqs:
        with st.expander(q):
            st.write(a)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Survey
    st.markdown("### Interested in FRED?")
    st.markdown("We're in beta. Leave your email and we'll let you know when we launch.")
    with st.form("landing_survey", clear_on_submit=True):
        survey_email = st.text_input("Your email address")
        notify = st.radio("Would you like to be notified when FRED launches?", ["Yes, notify me", "No thank you"])
        submitted = st.form_submit_button("Submit")
        if submitted and survey_email:
            send_emailjs(survey_email, "landing_survey", {"notify": notify})
            st.success("Thank you. No obligation — we'll only contact you about the launch.")


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
    st.markdown("## Upload your documents")
    st.markdown("Upload your child's EHCP below. You can add further documents — EP report, OT report, school policy — using the expander beneath.")

    # Main upload
    main_file = st.file_uploader(
        "Upload EHCP (PDF or Word document)",
        type=["pdf", "docx"],
        key="main_upload"
    )

    # Optional additional
    with st.expander("Add another document (optional)"):
        st.markdown("You can add an EP report, OT report, diagnosis letter, or school policy. FRED will cross-reference it with the EHCP.")
        extra_file = st.file_uploader(
            "Additional document",
            type=["pdf", "docx"],
            key="extra_upload"
        )

    st.markdown("---")
    st.markdown("### A few quick questions")
    st.markdown("These help FRED tailor the analysis. Answer what you can — nothing here is mandatory.")

    doc_type = st.radio(
        "Is this a draft or final EHCP?",
        ["Draft EHCP", "Final EHCP"],
        index=1,
        horizontal=True,
    )

    situation = st.text_area(
        "Briefly describe your current situation (optional)",
        placeholder="For example: annual review coming up in March, school not delivering speech therapy, considering appeal...",
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
    )

    st.markdown("---")

    if st.button("Analyse my documents", use_container_width=False):
        if main_file is None:
            st.error("Please upload your EHCP to continue.")
        else:
            with st.spinner("Reading your document…"):
                if main_file.name.endswith(".pdf"):
                    full_text = extract_text_from_pdf(main_file)
                else:
                    full_text = extract_text_from_docx(main_file)

                extra_text = ""
                if extra_file is not None:
                    if extra_file.name.endswith(".pdf"):
                        extra_text = extract_text_from_pdf(extra_file)
                    else:
                        extra_text = extract_text_from_docx(extra_file)

                combined_text = full_text + "\n\n" + extra_text
                findings, meta = run_full_analysis(combined_text)

            st.session_state.findings         = findings
            st.session_state.parse_meta       = meta
            st.session_state.full_text        = combined_text
            st.session_state.doc_type         = doc_type
            st.session_state.situation        = situation
            st.session_state.upcoming_dates   = upcoming_dates
            st.session_state.relationship_tone = relationship_tone
            st.session_state.doc_name         = main_file.name
            st.session_state.stage            = "sneak_peek"
            st.rerun()


def page_sneak_peek():
    findings = st.session_state.findings

    st.markdown("## Your report is ready")
    st.markdown(
        "Here is one real finding from your EHCP analysis. "
        "This is exactly what the full report looks like — no preview blur, no teaser. "
        "The most significant finding is shown below."
    )

    # Show the top red finding (or first finding if no reds)
    red_findings = [f for f in findings if f["tier"] == "red"]
    preview_finding = red_findings[0] if red_findings else findings[0] if findings else None

    if preview_finding:
        st.markdown('<div class="sneak-box">', unsafe_allow_html=True)
        st.markdown("**Preview finding — from your EHCP:**")
        render_finding_card(preview_finding, show_full=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Summary counts
    red_n   = sum(1 for f in findings if f["tier"] == "red")
    amber_n = sum(1 for f in findings if f["tier"] == "amber")
    green_n = sum(1 for f in findings if f["tier"] == "green")

    st.markdown(f"""
    <div style="background:white;border:1px solid #d0dae8;border-radius:8px;padding:1.2rem;margin:1.2rem 0;">
      <p style="margin:0;font-size:1rem;">
        Your full report contains:&nbsp;
        <span style="color:{RED};font-weight:700;">{red_n} Red</span> &nbsp;·&nbsp;
        <span style="color:{AMBER};font-weight:700;">{amber_n} Amber</span> &nbsp;·&nbsp;
        <span style="color:{GREEN};font-weight:700;">{green_n} Green</span>
      </p>
      <p style="margin:0.5rem 0 0 0;font-size:0.9rem;color:#555;">
        Upload first. Decide after. Your report is ready before you pay.
      </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Get your full report")
    st.markdown("Enter your email address below. During beta, full access is free — no card details required. Your email is how we send you your report and keep you updated.")

    if not st.session_state.email_submitted:
        with st.form("email_capture"):
            email = st.text_input("Your email address")
            submitted = st.form_submit_button("Get my full report")
            if submitted:
                if email and "@" in email:
                    send_emailjs(email, "beta_signup", {"message": f"Beta signup — doc: {st.session_state.doc_name}"})
                    st.session_state.email_submitted = True
                    st.session_state.show_full_report = True
                    st.rerun()
                else:
                    st.error("Please enter a valid email address.")
    else:
        if st.button("View my full report"):
            st.session_state.show_full_report = True
            st.session_state.stage = "full_report"
            st.rerun()

    if st.session_state.show_full_report and st.session_state.email_submitted:
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
    <div style="background:#f0f4fa;border-radius:6px;padding:0.8rem 1.2rem;margin-bottom:1.2rem;font-size:0.9rem;color:#444;">
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
        <span style="color:{RED};font-weight:700;">{red_n} Red</span> — lawful requirement(s) not met.
        Must be addressed at annual review.<br>
        <span style="color:{AMBER};font-weight:700;">{amber_n} Amber</span> — best practice gap(s).
        Worth raising at annual review.<br>
        <span style="color:{GREEN};font-weight:700;">{green_n} Green</span> — compliant.
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
        "\n\n"
        "Take your report to your annual review meeting. "
        "You can also upload correspondence — emails from school or the LA — "
        "and FRED will analyse them for you."
    )

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
        "<p style='font-size:0.85rem;color:#888;'>Word is best for Windows/Office users. PDF is best for Apple devices.</p>",
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

    if not st.session_state.survey_submitted:
        st.markdown("---")
        page_survey()


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
            email_to_send = notify_email if notify_email else survey_email if survey_email else "no-email@fred.invalid"
            send_emailjs(email_to_send, "survey", extra_dict)
            st.session_state.survey_submitted = True
            st.success("Thank you. Your feedback helps us build FRED properly.")


# ── CORRESPONDENCE MODULE ─────────────────────────────────────────────────────

def page_correspondence():
    st.markdown("## Correspondence analysis")
    st.markdown(
        "Upload up to two recent emails from the school or LA. "
        "FRED will analyse them for you — tone, intent, gaps, and what to do next. "
        "This feature is part of the subscription. It is included in beta."
    )

    email1 = st.file_uploader("Most recent email (PDF or Word)", type=["pdf", "docx"], key="email1")
    email2 = st.file_uploader("Previous email (optional)", type=["pdf", "docx"], key="email2")

    tone_q = st.radio(
        "How would you describe how the school is engaging right now?",
        [
            "Genuinely trying to help",
            "Going through the motions",
            "Actively avoiding or obstructing",
        ],
        help="You know the relationship. FRED doesn't. Your answer calibrates the tone of the analysis."
    )

    if st.button("Analyse correspondence") and email1:
        with st.spinner("Reading correspondence…"):
            if email1.name.endswith(".pdf"):
                text1 = extract_text_from_pdf(email1)
            else:
                text1 = extract_text_from_docx(email1)

            text2 = ""
            if email2:
                if email2.name.endswith(".pdf"):
                    text2 = extract_text_from_pdf(email2)
                else:
                    text2 = extract_text_from_docx(email2)

            combined = text1 + "\n\n" + text2
            briefing = analyse_correspondence(combined, tone_q)

        st.markdown("### Briefing")
        for item in briefing:
            st.markdown(f"""
            <div class="finding-{item['tier']}">
              <span class="badge-{item['tier']}">{item['label']}</span>
              <p style="font-weight:700;margin:0.4rem 0 0.3rem;">{item['title']}</p>
              <p style="margin:0;font-size:0.95rem;">{item['detail']}</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### What would you like to do?")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.button("Draft a response email")
        with col2:
            st.button("Get a briefing note")
        with col3:
            st.button("Hold — save to vault without acting")

        st.info(
            "Hold is a valid strategic choice. "
            "FRED saves the finding with a date. You can return to it at any time. "
            "Sometimes the most powerful response is no response."
        )


def analyse_correspondence(text, tone):
    """
    Three-finding briefing from correspondence.
    Returns list of up to 3 finding dicts for the briefing display.
    """
    findings = []

    # Signal 1: Reassurance without evidence
    reassurance_pattern = re.compile(
        r'\b(we are not concerned|no concerns|doing well|runs continuously|'
        r'you would be our first point of contact|we would contact you|'
        r'fully supported|fully meeting)\b',
        re.IGNORECASE
    )
    m = reassurance_pattern.search(text)
    if m:
        ctx = text[max(0, m.start()-100):m.end()+100].strip()
        findings.append({
            "tier": "amber",
            "label": "PATTERN",
            "title": "Reassurance without evidence",
            "detail": (
                f'The correspondence contains reassurance language without documentary support. '
                f'"{ctx[:200]}…" — '
                "Reassurance is not evidence. Ask what the evidence base is for this statement."
            ),
        })

    # Signal 2: HOY involvement
    hoy_pattern = re.compile(r'\b(head of year|HOY|form tutor|year team)\b', re.IGNORECASE)
    if hoy_pattern.search(text):
        findings.append({
            "tier": "amber",
            "label": "SIGNAL",
            "title": "Head of Year involvement detected",
            "detail": (
                "Communications initiated or copied to the Head of Year typically indicate "
                "a disruption management pathway, not a SEND support pathway. "
                "Check whether this correspondence references your child's EHCP or their needs, "
                "or whether it is focused on peer disruption or behaviour."
            ),
        })

    # Signal 3: Provision substitution
    sub_pattern = re.compile(
        r'\b(ordinarily available|quality first|whole class|classroom support|'
        r'all our students|all pupils)\b',
        re.IGNORECASE
    )
    m = sub_pattern.search(text)
    if m:
        ctx = text[max(0, m.start()-80):m.end()+80].strip()
        findings.append({
            "tier": "red",
            "label": "RED",
            "title": "Universal provision substituted for specified EHCP provision",
            "detail": (
                f'"{ctx[:200]}…" — '
                "The correspondence describes universal or classroom provision in response "
                "to a question about your child's EHCP provision. The question was not answered — "
                "it was redirected. Request a specific answer about the delivery of named provision in Section F."
            ),
        })

    # Signal 4: Implicit admission
    admit_pattern = re.compile(
        r'\b(we are working to improve|going forward|in future|we will ensure|'
        r'we have identified|we recognise that|we are reviewing)\b',
        re.IGNORECASE
    )
    m = admit_pattern.search(text)
    if m:
        ctx = text[max(0, m.start()-80):m.end()+80].strip()
        findings.append({
            "tier": "amber",
            "label": "SIGNAL",
            "title": "Implicit admission of current gap",
            "detail": (
                f'"{ctx[:200]}…" — '
                "This language implies that current practice is inadequate while presenting "
                "future improvement as reassurance. Note the date of this correspondence. "
                "This is an admission of a current gap, not a resolution of it."
            ),
        })

    # Vacuum detection
    vacuum_pattern = re.compile(
        r'\b(as discussed|as previously|on a number of occasions|we have tried|'
        r'we have been working|ongoing concern)\b',
        re.IGNORECASE
    )
    if vacuum_pattern.search(text):
        findings.append({
            "tier": "amber",
            "label": "SIGNAL",
            "title": "References to undocumented history",
            "detail": (
                "The correspondence references previous discussions or attempts without "
                "providing documentation. Request the records that evidence these conversations. "
                "If they do not exist, the history did not happen in any lawful sense."
            ),
        })

    # Tone note
    tone_note = {
        "Genuinely trying to help": "School appears to be engaging genuinely. Hold this finding and use it as context — a school that is trying may respond better to collaborative framing.",
        "Going through the motions": "School appears to be going through the motions. Written requests with specific reference to EHCP section numbers will be more effective than verbal escalation.",
        "Actively avoiding or obstructing": "School appears to be avoiding or obstructing. Every communication should be in writing. Reference specific statutory duties. Consider whether escalation to the LA is appropriate.",
    }.get(tone, "")

    if tone_note:
        findings.append({
            "tier": "amber",
            "label": "TONE",
            "title": "Relationship read",
            "detail": tone_note,
        })

    return findings[:3]  # Three finding maximum


# ── NAVIGATION ────────────────────────────────────────────────────────────────

def render_nav():
    col1, col2, col3, col4 = st.columns([1, 1, 1, 4])
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
        if st.button("Correspondence"):
            st.session_state.stage = "correspondence"
            st.rerun()


# ── ROUTER ────────────────────────────────────────────────────────────────────

render_nav()
st.markdown("<hr style='margin:0 0 1.5rem;border-color:#d0dae8;'>", unsafe_allow_html=True)

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
elif stage == "correspondence":
    page_correspondence()
else:
    page_landing()
