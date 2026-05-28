FRED Build Method
The permanent record — method statement, build log, and decisions
Living document. Updated after every session. Last updated: Session Debrief 05, 27 May 2026.
---
What this document is
FRED is built session by session. The code lives in GitHub. The reasoning behind every decision lives here.
This document has four jobs:
Method statement — how FRED is built and why it works the way it does
Build log — what has been done, in what order, and what the current state is
Decision record — why each significant decision was made, grounded in the legal standard
Recovery document — if something breaks, this is where you find out what it is and why it was built that way
If you are a new builder starting a session, read this before touching any code.
---
The product — what FRED is for
A parent uploads their child's EHCP.
FRED tells them in plain English:
What Section F commits to — the provision inventory
Where the language is vague or unenforceable — the red findings
Where EP/SALT/OT recommendations were not converted to commitments
Where needs in Section B have no provision in Section F
When the parent gets an email from school saying provision isn't happening:
FRED finds the Section F commitment
Shows it to the parent
Tells them what to ask for in writing
That's the product. Everything built serves that.
---
The legal standard
Every RED finding must be citable. The three sources:
Children and Families Act 2014 s42 — the duty to deliver specified Section F provision is absolute. It does not depend on school resources, budget, or staffing.
SEND Code of Practice 2015 para 9.69 — Section F must be "detailed and specific and normally quantified."
SOS!SEN EHCP Infosheet 2022 — explicit list of terms that must be challenged in Section F because they make provision unauditable and unenforceable.
A parent must be able to take a FRED finding to an annual review and defend it by citing one of these.
---
The five principles — never break these
1. Explicit failure over silent failure
Wrong output with no warning is worse than an honest error. If FRED cannot find Section F, it says so clearly. If a pattern fires on something ambiguous, the finding explains the ambiguity.
2. Grounded in the legal standard
Every RED finding must be citable. Opinion is not a finding. A finding that cannot be grounded in the Children and Families Act 2014, the SEND Code of Practice, or SOS!SEN guidance is not a finding.
3. Plain English output
The analysis is deterministic. The output is for a parent who may never have read the legislation. The gap between those two things is FRED's job.
4. Stability before sophistication
One change at a time. Test before the next change. Never rewrite working functions. No change to ehcp_parser/ files without overseer review.
5. The scales principle
Every decision tips toward the child and family. Ask of every finding: "If this fires, who does it protect?" If the answer is the institution, the finding is correct. If the answer is unclear, the finding needs more thought.
A false positive RED that a parent can verify in 30 seconds is better than a missed RED that costs a child an enforceable right.
---
The build system
Three roles, strictly separated
The overseer (this conversation)
Holds the full build tracker. Receives session debriefs. Updates this document. Writes the brief for the next task. Does not write code. Does not make build decisions mid-session.
Build sessions (separate Claude conversations)
One task per conversation. Receives the session starter, the current fred_app.py, and the task brief. Nothing else. Has no access to previous sessions or the broader roadmap. One task, done properly, tested, debrief produced.
Session debriefs (PDF, numbered sequentially)
The intelligence layer. Every session ends with a debrief covering what was built, why each decision was made, what was tested, and what comes next. Brought to the overseer after every session. Numbered 01, 02, 03 and so on.
The rule that prevents things going wrong
One task per conversation. No exceptions. A build chat that finishes one task and starts another is already drifting.
The test — non-negotiable
Before and after every change:
Go to the Streamlit app
Upload the Warwickshire EHCP
Confirm Section F blocks found = 4
Confirm findings make sense
If Section F blocks ≠ 4, stop. Do not continue. Come back to the overseer.
---
File structure and rules
```
fred_app.py                   — main app, 3000+ lines. Never rewrite. One change at a time.
requirements.txt              — streamlit, pymupdf, python-docx, reportlab, requests, pdfplumber
FRED_SESSION_STARTER.md       — paste at the top of every new builder chat
FRED_BUILD_METHOD.md          — this document
sandbox/
  provision_inventory.py      — provision categorisation module
  __init__.py
ehcp_parser/                  — core parser. DO NOT MODIFY without overseer review.
  __init__.py
  core.py
  models.py
  section_registry.py
  classifier.py
  pipeline.py
```
ehcp_parser/ is protected. Any change to these files requires the overseer to review it first. Flag with "this needs overseer review" and stop.
---
How to edit fred_app.py
Use the GitHub web editor. On the main repo page at github.com/smd883-ctrl/fred-beta-Claude, press the full stop key ( . ) on your keyboard. The page reloads into a browser-based VS Code editor. No installation required.
To find something: Click fred_app.py in the left panel. Press Ctrl+F. Search for a short unique string.
To save changes: Click the branch icon (Y shape) in the left sidebar. Type a commit message. Click Commit and Push.
The app restarts on Streamlit automatically within 1–3 minutes of committing.
Note: Ctrl+S in the web editor has no visible confirmation. This is normal. Changes are held in memory until you commit.
---
Infrastructure decisions — decided, not yet built
Component	Decision
Database and auth	Supabase
Payments	Stripe — hosted checkout, no custom UI
Hosting	Streamlit Community Cloud now, Render when needed
Builder dashboard	Private Streamlit page connected to Supabase
API cost management	Hash every PDF, cache classifier results in Supabase
Five Supabase tables needed
`users` — handled by Supabase Auth
`subscriptions` — tier, status, Stripe customer ID, renewal date
`document_logs` — PDF hash, LA, format type, timestamp, user ID
`format_cache` — PDF hash → format JSON (never classify the same document twice)
`unknown_formats` — hash, first 3 pages text, date, user ID, status
`provision_patterns` — anonymised tagged findings for best practice database (opt-in)
---
Build tracker — current state
Current as of Session Debrief 05, 27 May 2026.
Phase 0 — Pre-Beta (do before testers see FRED)
#	Task	Status
0.1	UI refresh — replace hard green, improve readability	🔴 TO DO
0.2	Supabase project setup — six core tables	🔴 TO DO
0.3	Supabase AUTH — email/password login for testers	🔴 TO DO
0.4	Document store — save extracted text against user account	🔴 TO DO
0.5	Findings store — save findings with timestamp	🔴 TO DO
0.6	Session persistence — user returns and picks up where they left off	🔴 TO DO
0.7	Multi-document cross-reference within session	🟡 PARTIAL — vault exists
Phase 1 — Parser Foundation
#	Task	Status
1.1	Build ehcp_parser folder and files	✅ DONE
1.2	Fix find_section_blocks() boundary bug	✅ DONE
1.3	Confirm Section F blocks = 4 on Warwickshire	✅ DONE
1.4	Fix dilution clause false positive	✅ DONE
1.5	Fix Section E timeframe false positive	✅ DONE
1.6	Two-column PDF extraction (Warwickshire)	🟡 DONE — not fully tested on other LA formats
1.7	AI format classifier	🔴 TO DO — after Phase 0
1.8	Multi-LA format testing	🔴 TO DO — after classifier built
Phase 2 — Analysis Layer
#	Task	Status
2.1	Rewrite Section F language check	✅ DONE
2.2	Build provision inventory output	✅ DONE — PROVISION_LIBRARY built
2.3	Build EP recommendation check	✅ DONE
2.3A	Dynamic child name extraction	✅ DONE — extract_child_name() added
2.3B	Plain English narrative output — RED and GREEN rewritten	✅ DONE
2.3C	Grow _REC_STOPWORDS as new documents reveal weak terms	🟡 ONGOING
2.3D	Fix extract text truncation in VAGUE_MODAL and VAGUE_QUALIFIER	🔴 TO DO — priority
2.4	Section B to F gap check — enhance with PROVISION_LIBRARY	🔴 TO DO
2.5	Plain English provision summary — Your child's EHCP commits to…	🔴 TO DO
Phase 3 — Knowledge Bank
#	Task	Status
3.1	Store provision inventory per child	🔴 TO DO — enabled by Phase 0
3.2	Cross-reference school emails vs Section F	🟡 PARTIAL
3.3	Flag when email contradicts EHCP commitment	🔴 TO DO
3.4	Annual review preparation checklist	🔴 TO DO
3.5	APDR reader architecture	🔴 TO DO — after Phase 0
Phase 4 — Product
#	Task	Status
4.1	Pricing and payment integration — Stripe	🔴 TO DO — after beta
4.2	User accounts — enabled by Phase 0 Supabase AUTH	🔴 TO DO
4.3	Multi-document upload (EHCP + EP + OT)	🟡 PARTIAL
4.4	Report branding and quality	🟡 IN PROGRESS
4.5	Hosting migration — Render or Railway	🔴 TO DO — after paying users
4.6	Builder dashboard — private Streamlit + Supabase	🔴 TO DO — after beta
Phase 5 — EHC Request Template (new)
#	Task	Status
5.1	Rights-first introduction — what parents are entitled to before any question	🔴 TO DO
5.2	Guided template — seven sections, plain English prompts, examples	🔴 TO DO
5.3	FRED review of draft — checks against Section 36 CFA 2014 threshold	🔴 TO DO
5.4	SENDIASS lookup — postcode in, local contact out. Standalone component.	🔴 TO DO
5.5	Output — bullet point parent statement, evidence checklist, Word download	🔴 TO DO
5.6	Anonymised data contribution — opt-in, tags need areas and LA only	🔴 TO DO
Phase 6 — Best Practice Database (new)
#	Task	Status
6.1	provision_patterns Supabase table — tagged anonymised findings	🔴 TO DO
6.2	Opt-in consent built into template flow from day one	🔴 TO DO
6.3	Builder dashboard aggregate view — patterns by LA, need area, provision gap	🔴 TO DO
6.4	Template improvement loop — successful request language feeds back into prompts	🔴 TO DO
Correspondence module — ahead of schedule
#	Task	Status
C.1	Pattern library — initial 12 patterns	✅ DONE
C.2	Six new patterns including collaborative mask	✅ DONE
C.3	Tone scoring with escalation detection	✅ DONE
C.4	Environment-specific sensory checklists	✅ DONE
C.5	Draft reply generator	✅ DONE
C.6	Extract display — From your EHCP label, 400 char limit	✅ DONE
---
Key decisions — the reasoning behind the code
These decisions shaped FRED in ways that are not obvious from the code alone.
VAGUE_MODAL and VAGUE_QUALIFIER are separate lists (Debrief 01)
The original PROHIBITED_WORDS list mixed weak modal verbs with vague qualifiers. They now produce two distinct RED findings with different legal explanations and different remedies. A commitment failure (VAGUE_MODAL) requires rewriting as "will receive." A specificity failure (VAGUE_QUALIFIER) requires quantifying the provision. The distinction matters because the remedy is different.
Dilution clause fires on constructions, not words (Debrief 01)
The original pattern fired on single words: "resources", "staffing", "availability." This produced false positives on "sensory resources" and "staffing arrangements." The rewritten pattern fires only on conditional constructions: "subject to staffing", "where resources allow." A finding that fires incorrectly undermines trust in the tool.
Baseline check removed from Section E (Debrief 02)
Baselines belong in assessment reports, not in the EHCP. The EHCP is a commitment document. Checking for baselines in Section E penalised correctly structured EHCPs and was actively misleading parents. Removed entirely.
Vague aspiration language is always RED (Debrief 02)
"Will have developed", "will have improved", "will have made progress" cannot be failed at annual review. The school can always point to some improvement. The language protects the institution from accountability while appearing child-centred. Always RED. This is the clearest application of the scales principle.
Classifier before public release — not after (Debrief 03)
The SEND parent network is fast, tight, and unforgiving. A broken report shared in a Facebook group does more damage than a polished one does good. The AI format classifier must be built before any user touches FRED. Classifier scope is fixed: first 3 pages only, JSON output, four format types, one API call per upload, never touches provision content.
Unknown format messaging places responsibility correctly (Debrief 03)
When FRED encounters an unknown PDF format the message reads: "FRED has not seen this EHCP format before. This is likely due to the way your local authority has structured the document rather than anything in your plan itself." This is accurate, honest, and does not undermine the parent's confidence in their own document.
Recommendation matching uses a two-term threshold (Debrief 04)
One matching term is too easy — any two documents about the same child share single words. Three or more is too strict — professionals routinely rephrase when writing from a report into an EHCP. Two is the working balance, with stopwords ensuring the matched terms are substantive provision terms, not noise.
Vagueness and presence are separate checks — never merge them (Debrief 04)
Task 2.3 answers one question: is the substance of this recommendation in Section F at all? Vagueness is checked separately by analyse_section_f(). Merging the two would make findings harder to maintain, harder to test, and harder to explain to a parent.
Child name extracted dynamically, never hardcoded (Debrief 05)
extract_child_name() searches the EHCP text in two passes — first for an explicit label, then for two adjacent Title-case words near the date of birth marker. The name is added dynamically to _REC_STOPWORDS and stored in st.session_state['child_name']. Without this, the child's name — present in almost every sentence of a professional report — scores false matches.
Persistence before beta — revised build order (strategic)
FRED's real value is layered intelligence — EHCP plus EP report plus correspondence plus history, documents talking to each other across time. Without persistence, testers experience a walk-in clinic. The product is a patient record. Build Supabase persistence before testers see FRED.
EHC request template is a second front door (strategic)
Parents who do not yet have an EHCP are more numerous, more frightened, and currently less supported. A guided request template reaches a different parent at the most daunting moment of the process. Parents who use the template and get an EHCP come back to FRED to analyse it. The two features feed each other.
Anonymised pattern data compounds over time (strategic)
Every template submission and EHCP analysis generates anonymised pattern data. Aggregated across hundreds of parents, this becomes the only real-time evidence base for how EHCPs are actually working across England. The database is a product asset, not a byproduct. Designed into Supabase from the start with a clear opt-in consent model.
SENDIASS lookup is a reusable standalone component (strategic)
Built once, used in multiple places: the request template before submission, the correspondence module when escalation language is detected, the annual review preparation checklist. Every parent FRED works with should know their SENDIASS service exists.
GREEN cards are deliberate (Debrief 01)
FRED rewards correct practice. A tool that only produces RED and AMBER becomes adversarial and loses credibility. GREEN cards make FRED a fair witness. They also give parents a benchmark when challenging non-compliant provision.
---
Session debrief index
Session	What it covered	Status
Debrief 01	Section F rewrite, VAGUE_MODAL/QUALIFIER, dilution clause, provision library, 6 correspondence patterns, tone scoring	✅ Received
Debrief 02	Section E rewrite, baseline removal, timeframe regex extended, three outcome language categories, scales principle	✅ Received
Debrief 03	Two-column PDF extraction, extract display fix, infrastructure strategy, five decisions identified	✅ Received
Debrief 04	Task 2.3 EP recommendation check built and tested, three design decisions, three false positive iterations fixed	✅ Received
Debrief 05	Tasks 2.3A and 2.3B — dynamic child name extraction, plain English narrative rewrite. GitHub web editor established.	✅ Received
Debrief 06		Pending
Debrief 07		Pending
Debrief 08		Pending

Debrief 09		Pending
Debrief 10		Pending
---
Operations one-pager
The things you need if something breaks.
What	Where
Live app	Streamlit Community Cloud — your deployment URL
Code	github.com/smd883-ctrl/fred-beta-Claude
Main file	fred_app.py in the repo
Edit code	Press . on the GitHub repo page
Roll back a change	GitHub → fred_app.py → History → pick a version → Restore
App not restarting	Go to Streamlit Cloud dashboard, click Reboot
Something broke and you don't know why	Come to the overseer conversation. Paste the error. Do not guess.
The recovery rule: If Section F blocks ≠ 4 on the Warwickshire EHCP, something broke. Roll back the last commit in GitHub before doing anything else.
---
FRED — Families' Rights and Entitlements Directory. Built by a parent. Not a legal service. Plain English findings grounded in the Children and Families Act 2014 and the SEND Code of Practice 2015.