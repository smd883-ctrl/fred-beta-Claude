FRED — SESSION STARTER
Paste this at the top of every new Claude conversation before describing the task
---
How to work with me — read this first
You are helping me build FRED (Families' Rights and Entitlements Directory), a Streamlit web app hosted on Streamlit Community Cloud via GitHub repo `smd883-ctrl/fred-beta-Claude`. The main file is `fred_app.py` — 3000+ lines. Do not rewrite it. One change at a time.
I am not a developer. I can follow precise instructions but I need them written for someone working directly in GitHub's web editor, not a local IDE.
Before every code change tell me:
Which file we are changing
Roughly what line number to look for (I will use Ctrl+F to search)
The exact search string to find the location — short, unique, copy-pasteable
What to select and what to replace it with
Whether this is a complete function replacement or a targeted line change
After every code change tell me:
What to test and what result to expect
What a passing result looks like in plain English
What a failing result looks like and what to do if it happens
Why we are doing each thing: Before writing any code, tell me in one short paragraph what problem we are solving and why it matters for FRED. I need to understand the reasoning, not just follow instructions. This is my tool and I need to understand what is inside it.
How to give me code:
Give me complete copy-pasteable blocks — no ellipsis, no "add the rest as before"
If replacing a function, give me the entire function
If replacing a block, give me the entire block with enough surrounding context to find it
Never assume I can see what you are referring to — name it explicitly every time
One change at a time. Make one change. Wait for me to test it. Confirm it works. Then move to the next. Do not give me three changes at once unless they are genuinely inseparable.
If I report an error: Ask me to paste the exact error text. Read it carefully before responding. Give me one targeted fix, not a rewrite.
If I cannot find something: Give me an alternative search string. Then another. If I still cannot find it, offer to give me the complete function to replace in one go.
What not to do:
Do not rewrite `fred_app.py` from scratch
Do not touch `ehcp_parser/` files without flagging "this needs overseer review"
Do not make changes to fix one thing that might break another without warning me first
Do not assume I know where something is — always tell me
Do not give me partial code blocks with "..." in them
At the end of every session: Produce a Session Debrief PDF, numbered sequentially. Cover: what was built, why each decision was made, what was tested and the result, what comes next. This is not optional — it is how the tool's intelligence is preserved between sessions.
What I need from you above everything else: Patience. Clear instructions. One step at a time. Tell me what we are doing and why before you tell me how. I am building something that matters — a tool that helps parents fight for their children's legal entitlements. Every change should make that tool more accurate, more honest, and more useful to the people who need it.
---
What FRED is
FRED (Families' Rights and Entitlements Directory) is a Streamlit web app that helps parents of children with EHCPs (Education, Health and Care Plans) understand what their child is legally entitled to and whether it is being delivered.
Built by a parent. Not a legal service. Plain English findings. Red/Amber/Green output.
---
The legal standard FRED works to
Every RED finding must be citable. The three sources are:
Children and Families Act 2014 s42 — the duty to deliver specified Section F provision is absolute. It does not depend on school resources, budget, or staffing.
SEND Code of Practice 2015 para 9.69 — Section F must be "detailed and specific and normally quantified".
SOS!SEN EHCP Infosheet 2022 — explicit list of terms that must be challenged in Section F because they make provision unauditable and unenforceable.
A parent must be able to take a FRED finding to an annual review and defend it by citing one of these sources.
---
The core philosophy — never break this
Explicit failure over silent failure. Wrong output with no warning is worse than an honest error. If FRED cannot find Section F, it says so clearly.
Grounded in the legal standard. Every RED finding must be citable. Opinion is not a finding.
Plain English output. The analysis is deterministic. The output is for a parent who may never have read the legislation. The gap between those two things is FRED's job.
Stability before sophistication. One change at a time. Test before the next change. Never rewrite working functions.
The scales principle. Every decision tips toward the child and family. Ask of every finding: "If this fires, who does it protect?" If the answer is the institution, the finding is correct. If the answer is unclear, the finding needs more thought.
---
What's already built and working
Live app: Streamlit, hosted on Streamlit Community Cloud via GitHub repo `smd883-ctrl/fred-beta-Claude`
GitHub file structure:
```
fred_app.py                   — main Streamlit app (3000+ lines, do not rewrite)
requirements.txt              — streamlit, pymupdf, python-docx, reportlab, requests, pdfplumber
sandbox/
  provision_inventory.py      — provision categorisation module
  __init__.py
ehcp_parser/                  — core parser (do not modify without overseer review)
  __init__.py
  core.py
  models.py
  section_registry.py
  classifier.py
  pipeline.py
```
---
Current state of fred_app.py — what is built and working
PDF extraction
Primary extractor: pdfplumber with x-coordinate column separation
Detects two-column layout (Warwickshire/Capita Synergy format)
Left column (provision) extracted first, right column (deliverer) appended with label
Fallback: PyMuPDF if pdfplumber fails
Split threshold: 65% of page width (~390pt on A4)
Section block finder
find_section_blocks() uses need-area group headings as boundaries
Correctly finds 4 Section F blocks on Warwickshire EHCP (one per need area)
Need area boundaries: Communication and Interaction, Social Emotional and Mental Health, Cognition and Learning, Physical and/or Sensory
Section F analysis — analyse_section_f()
Two separate vague language checks, each producing its own RED finding:
VAGUE_MODAL — commitment failures. These replace "will receive" with something conditional or aspirational. The fix is always to rewrite as "will receive."
Current list: "would benefit from", "should receive", "may be provided", "could receive", "might benefit", "it is recommended", "it would be helpful", "it is suggested", "it is hoped"
VAGUE_QUALIFIER — specificity failures. These state provision without enough detail to audit. The fix is always to quantify: frequency, duration, named role.
Current list: "access to", "as appropriate", "where necessary", "as needed", "regular", "flexible", "tailored", "embedded", "holistic", "opportunities", "encouraged", "supported to", "up to", "as directed", "as advised", "as required"
Dilution clause — fires only on conditional constructions, not standalone words.
Triggers: "subject to staffing/resources/funding/budget/availability", "where resources allow", "if staffing allows", "where budget permits", "budget permitting", "dependent on availability", "subject to funding"
Does NOT fire on: "sensory resources", "staffing arrangements", "available support" — these are legitimate provision language.
Also checked in Section F:
Frequency of provision (number + session/hour/week etc) — RED if absent
Duration of sessions (minutes or hours) — RED if absent
Deliverer role specified — RED if absent
Universal provision substituted for specific provision — RED if detected
Recommendation language in Section F ("it is recommended") — RED, not converted to commitment
Delivery log mechanism referenced — AMBER if absent
Section E analysis — analyse_section_e()
Three categories of outcome language, each treated differently:
Achievement language — correct. "Will be able to", "will independently", "will demonstrate", "will initiate". Describes what the child can do. Can be assessed at annual review.
Activity language — RED if no achievement language present, AMBER if mixed. "Will receive", "will attend", "will be supported". Describes provision not outcomes. Belongs in Section F not Section E.
Vague aspiration language — always RED. "Will have developed", "will have improved", "will have made progress", "will have gained". Cannot be failed at annual review. Protects the institution, not the child. Always RED — this is the clearest application of the scales principle.
Timeframe check — extended regex recognises all EHCP-relevant formats:
"by [month] 20XX" — dated
"within X months/weeks/terms/years"
"by the end of Key Stage X" / "by the end of the Key Stage X"
"by the end of Year X" / "by Year X"
"by the next annual review" / "at the annual review"
"by the age of X" / "by age X"
"before transition to secondary" / "before starting college"
"by the end of primary" / "before leaving secondary"
DO NOT check for baselines. Baselines belong in assessment reports (EP report, SALT report, OT report), not in the EHCP. The EHCP is a commitment document. Checking for baselines in Section E is incorrect and was removed.
Section B analysis — analyse_section_b()
Checks four need areas against Section F:
Communication and Interaction
Cognition and Learning
Social, Emotional and Mental Health
Sensory and Physical
Each area: if identified in Section B with no provision in Section F → RED. If Section F provision is vague → AMBER. Each finding includes "what good looks like" guidance for the parent.
Provision library — PROVISION_LIBRARY
Structured dictionary of known legitimate SEND provision terms, four categories matching Section B need areas. Powers:
B to F cross-reference (need identified, provision named or absent)
Provision inventory (plain English summary)
Correspondence green card (correct SEND language recognised in school emails)
Extract display
Every finding shows a clearly labelled block: "From your EHCP:" above the italic extract text. Character limit 400. Applies in app display, Word download, and PDF download.
Correspondence module
Pattern library of 18 patterns. Each pattern has: tier (RED/AMBER/GREEN), explanation, and "the question to ask".
RED patterns: Resources Defence, Relationship Breakdown Threat, Reintegration Promise Without Plan, Unstated Staffing Change, Legal Obligation Misrepresented, Behaviour Need Framed as Behaviour Problem, Veiled Threat, Incident During Unstructured Time, Process Substitution, Escalation Sequence, Retrospective Justification, Provision Substitution
AMBER patterns: Best Interests Redirect, Blip Normalisation, Complaint Policy Deflection, Good Week Signal, Home/School Discrepancy, Collaborative Mask, Home Responsibility Redirect, Monitor Without Action, Reassurance Without Evidence, Implicit Admission
GREEN pattern: Correct SEND Language Used — acknowledges when a school uses specific provision language (sensory diet, Zones of Regulation, trusted adult, co-regulation) and prompts the parent to check whether it appears in Section F and whether a delivery log exists.
Tone scoring: Counts formal signals, collaborative signals, escalation signals, and operational SEND signals. Collaborative mask detection: if collaborative language AND escalation language appear together, recommendation is formal regardless of warm tone.
Environment detection: Identifies where an incident occurred (food hall, corridor, playground etc) and surfaces a sensory/environmental checklist for that space.
Draft reply generator: Builds a focused written response based on the top two detected patterns and the recommended tone.
---
What is NOT yet built
EP recommendation laundering check (Task 2.3) — BUILT. Four helper functions: _split_into_sentences(), _extract_key_terms(), _will_sentences_from_section_f(), _score_match(). Main function: analyse_report_recommendations(). Three constants: RECOMMENDATION_TRIGGERS, _REC_STOPWORDS, _REC_MATCH_THRESHOLD (=2). Follow-up tasks still needed:
2.3A — Dynamic child name extraction. Extract child's name from EHCP at parse time (near 'dob' or 'date of birth'), add to stopwords dynamically. Store in session state. Priority.
2.3B — Plain English narrative output. Rewrite RED and GREEN finding text — no algorithm language visible to parent. Priority.
2.3C — Continue growing _REC_STOPWORDS as new test documents reveal weak terms.
B to F gap check enhancement (Task 2.4) — use PROVISION_LIBRARY categories for richer matching. Current version partially works.
Plain English provision summary (Task 2.5) — "Your child's EHCP commits to: ..." at top of report. Not yet built.
Supabase connection — not yet built. Planned for Phase 0 before beta.
AUTH (user login) — not yet built. Planned for Phase 0.
Persistence (documents and findings surviving page refresh) — not yet built.
---
Infrastructure direction — decided, not yet built
Database and auth: Supabase
Payments: Stripe (hosted checkout, no custom UI)
Hosting: Streamlit Community Cloud now, Render when needed
Builder dashboard: Private Streamlit page connected to Supabase
API cost management: Hash every PDF, cache classifier results in Supabase
Five Supabase tables needed:
`users` — handled by Supabase Auth
`subscriptions` — tier, status, Stripe customer ID, renewal date
`document_logs` — PDF hash, LA, format type, timestamp, user ID
`format_cache` — PDF hash → format JSON
`unknown_formats` — hash, first 3 pages text, date, user ID, status
---
Rules for every session
Make one change at a time
Always give copy-paste ready code
Always say which file and which function to change
Never rewrite fred_app.py from scratch
Never touch ehcp_parser/ files without flagging "this needs overseer review"
Do not add baselines to Section E analysis under any circumstances
Do not merge VAGUE_MODAL and VAGUE_QUALIFIER back into a single list
The test: Upload Warwickshire EHCP → Section F blocks found = 4. If this breaks, stop and flag immediately. Do not continue.
---
At the end of every session
Produce a Session Debrief document. Number sequentially from where we are. The debrief covers:
What was built and what changed
Why each decision was made and what standard it was grounded in
What was tested and the result
What comes next
Write it in plain English. Save as PDF. This is the intelligence layer of FRED — the reasoning behind the code.