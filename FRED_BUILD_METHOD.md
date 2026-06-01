FRED
Families’ Rights and Entitlements Directory
THE FRED BUILD METHOD
Method statement · Build log · Decision record · Recovery document
Living document — updated after every session. Current to Session Debrief 09, 31 May 2026.

1. What FRED is for
A parent uploads their child’s EHCP. FRED tells them in plain English: what Section F commits to, where the language is vague or unenforceable, where EP/SALT/OT recommendations were not converted to commitments, and where needs in Section B have no provision in Section F.
When the parent gets an email from school saying provision isn’t happening: FRED finds the Section F commitment, shows it to the parent, and tells them what to ask for in writing.
That’s the product. Everything built serves that.



2. The legal standard
Children and Families Act 2014 s42 — the duty to deliver specified Section F provision is absolute. It does not depend on school resources, budget, or staffing.
SEND Code of Practice 2015 para 9.69 — Section F must be ‘detailed and specific and normally quantified’.
SOS!SEN EHCP Infosheet 2022 — explicit list of terms that must be challenged in Section F.



3. The five principles — never break these

1
Explicit failure over silent failure
Wrong output with no warning is worse than an honest error. If FRED cannot find Section F, it says so clearly.


2
Grounded in the legal standard
Every RED finding must be citable. SOS!SEN, CoP para 9.69, Children and Families Act 2014. Opinion is not a finding.


3
Plain English output
The analysis is deterministic. The output is for a parent who may never have read the legislation. The gap between those two things is FRED’s job.


4
Stability before sophistication
One change at a time. Test before the next change. Never rewrite working functions. No change to ehcp_parser/ without overseer review.


5
The scales principle
Every decision tips toward the child and family. Ask of every finding: ‘If this fires, who does it protect?’ If the answer is the institution, the finding is correct.




4. Perceived value and real value — always together
Established in Session 07. FRED must deliver two things simultaneously:
Real value — accurate, citable, legally grounded analysis that a parent can take to an annual review and defend.
Perceived value — an experience that makes the parent feel equipped, seen, and in control from the first moment they open their report.
These are not in tension. A tool that is legally accurate but emotionally alienating will not be used. A tool that feels reassuring but is not grounded will not be trusted. FRED must be both.
The ‘Your child’s EHCP commits to:’ section is the clearest expression of this. It is legally real — every sentence is drawn directly from Section F and represents an absolute obligation. But its placement before the findings changes the emotional experience. The parent first sees what they are owed. Then they see where it is failing. That order matters.
Apply this as a test to every new feature: not just ‘is this accurate?’ but ‘does this make the parent feel equipped?’



5. The build system
Three roles, strictly separated
The overseer holds the full build tracker, receives session debriefs, updates this document, and writes the brief for the next task. It does not write code.
Build sessions get one task per conversation: the session starter, the current fred_app.py, and the task brief. Nothing else. One task, done properly, tested, debrief produced.
Session debriefs are the intelligence layer. Every session ends with a PDF debrief covering what was built, why each decision was made, what was tested, and what comes next. Numbered 01, 02, 03 and so on.

The rule that prevents things going wrong
One task per conversation. No exceptions. A build chat that finishes one task and starts another is already drifting.

The stability test — non-negotiable
Before and after every change: upload the Warwickshire EHCP, confirm Section F blocks found = 4. If the number changes, stop. Do not continue. Come back to the overseer.



6. File structure and editing
Primary editing environment: GitHub web editor. On the repo page at github.com/smd883-ctrl/fred-beta-Claude, press the full stop key ( . ) to open the browser-based editor. Ctrl+F to search. Branch icon to commit and push. App restarts on Streamlit within 1–3 minutes.

#
Task
Status
fred_app.py
Main app, 3000+ lines. Never rewrite. One change at a time.
PROTECTED
FRED_SESSION_STARTER.md
Paste at top of every new builder chat.
LIVE
FRED_BUILD_METHOD.md
This document. Update after every session.
LIVE
ehcp_parser/
Core parser. DO NOT modify without overseer review.
PROTECTED
sandbox/provision_inventory.py
Provision categorisation module.
ACTIVE




7. Infrastructure decisions — decided, not yet built
#
Task
Status
Database and auth
Supabase
DECIDED
Payments
Stripe — hosted checkout, no custom UI
DECIDED
Hosting
Streamlit Community Cloud now, Render when needed
DECIDED
Builder dashboard
Private Streamlit page connected to Supabase
DECIDED
API cost management
Hash every PDF, cache classifier results in Supabase
DECIDED


Six Supabase tables needed
users — handled by Supabase Auth
subscriptions — tier, status, Stripe customer ID, renewal date
document_logs — PDF hash, LA, format type, timestamp, user ID
format_cache — PDF hash to format JSON. Never classify the same document twice.
unknown_formats — hash, first 3 pages text, date, user ID, status
provision_patterns — anonymised tagged findings for best practice database (opt-in)



8. Build tracker — current state
Current as of Session Debrief 09, 31 May 2026.

Phase 0 — Pre-Beta
#
Task
Status
0.1
UI refresh — replace hard green, improve readability
🔴 TO DO
0.2
Supabase project setup — six core tables
🔴 TO DO
0.3
Supabase AUTH — email/password login for testers
🔴 TO DO
0.4
Document store — save extracted text against user account
🔴 TO DO
0.5
Findings store — save findings with timestamp
🔴 TO DO
0.6
Session persistence — user returns and picks up where they left off
🔴 TO DO
0.7
Multi-document cross-reference within session
🟡 PARTIAL — vault exists


Phase 1 — Parser Foundation
#
Task
Status
1.1
Build ehcp_parser folder and files
✅ DONE
1.2
Fix find_section_blocks() boundary bug
✅ DONE
1.3
Confirm Section F blocks = 4 on Warwickshire
✅ DONE
1.4
Fix dilution clause false positive
✅ DONE
1.5
Fix Section E timeframe false positive
✅ DONE
1.6
Two-column PDF extraction (Warwickshire)
🟡 DONE — column scrambling partially resolved, Approach A pending
1.7
AI format classifier
🔴 TO DO — after Phase 0
1.8
Multi-LA format testing
🔴 TO DO — after classifier


Phase 2 — Analysis Layer
#
Task
Status
2.1
Rewrite Section F language check — VAGUE_MODAL and VAGUE_QUALIFIER
✅ DONE
2.2
Build provision inventory output — PROVISION_LIBRARY built
✅ DONE
2.3
Build EP recommendation check
✅ DONE
2.3A
Dynamic child name extraction — extract_child_name()
✅ DONE
2.3B
Plain English narrative output — RED and GREEN rewritten
✅ DONE
2.3C
Grow _REC_STOPWORDS as new documents reveal weak terms
🟡 ONGOING
2.3D
Fix extract text truncation — sentence-aware, trigger highlighted in red
✅ DONE
2.4
Section B to F gap check — two-stage PROVISION_LIBRARY matching
✅ DONE
2.5
Plain English provision summary — Your child’s EHCP commits to…
✅ DONE
2.5A
Download parity — commitments summary in Word and PDF downloads
🔴 TO DO — priority


Phase 3 — Knowledge Bank
#
Task
Status
3.1
Store provision inventory per child
🔴 TO DO — enabled by Phase 0
3.2
Cross-reference school emails vs Section F
🟡 PARTIAL
3.3
Flag when email contradicts EHCP commitment
🔴 TO DO
3.4
Annual review preparation checklist
🔴 TO DO
3.5
APDR reader architecture
🔴 TO DO — after Phase 0


Phase 4 — Product
#
Task
Status
4.1
Pricing and payment integration — Stripe
🔴 TO DO — after beta
4.2
User accounts — enabled by Phase 0 Supabase AUTH
🔴 TO DO
4.3
Multi-document upload (EHCP + EP + OT)
🟡 PARTIAL
4.4
Report branding and quality
🟡 IN PROGRESS
4.5
Hosting migration — Render or Railway
🔴 TO DO — after paying users
4.6
Builder dashboard — private Streamlit + Supabase
🔴 TO DO — after beta


Phase 5 — EHC Request Template
#
Task
Status
5.1
Rights-first introduction — what parents are entitled to
🔴 TO DO
5.2
Guided template — seven sections, plain English prompts
🔴 TO DO
5.3
FRED review of draft — checks against Section 36 CFA 2014
🔴 TO DO
5.4
SENDIASS lookup — postcode in, local contact out
🔴 TO DO
5.5
Output — bullet point parent statement, evidence checklist
🔴 TO DO
5.6
Anonymised data contribution — opt-in
🔴 TO DO


Phase 6 — Best Practice Database
#
Task
Status
6.1
provision_patterns Supabase table — tagged anonymised findings
🔴 TO DO
6.2
Opt-in consent built into template flow from day one
🔴 TO DO
6.3
Builder dashboard aggregate view — patterns by LA, need area, gap
🔴 TO DO
6.4
Template improvement loop — successful language feeds back into prompts
🔴 TO DO


Correspondence module — ahead of schedule
#
Task
Status
C.1
Pattern library — 18 patterns total
✅ DONE
C.2
Tone scoring with escalation and collaborative mask detection
✅ DONE
C.3
Environment-specific sensory checklists
✅ DONE
C.4
Draft reply generator
✅ DONE
C.5
Extract display — From your EHCP label, 400 char limit
✅ DONE




9. Key decisions — the reasoning behind the code
VAGUE_MODAL and VAGUE_QUALIFIER are separate lists (Debrief 01)
Two distinct RED findings with different legal explanations and different remedies. A commitment failure (VAGUE_MODAL) requires rewriting as ‘will receive’. A specificity failure (VAGUE_QUALIFIER) requires quantifying the provision. The remedy is different so the finding must be different.

Dilution clause fires on constructions, not words (Debrief 01)
The original pattern fired on single words: ‘resources’, ‘staffing’. This produced false positives on ‘sensory resources’ and ‘staffing arrangements’. The rewritten pattern fires only on conditional constructions: ‘subject to staffing’, ‘where resources allow’. A finding that fires incorrectly undermines trust in the tool.

Baseline check removed from Section E (Debrief 02)
Baselines belong in assessment reports, not in the EHCP. The EHCP is a commitment document. Checking for baselines in Section E penalised correctly structured EHCPs and was actively misleading parents. Removed entirely. Do not reinstate under any circumstances.

Vague aspiration language is always RED (Debrief 02)
‘Will have developed’, ‘will have improved’, ‘will have made progress’ cannot be failed at annual review. The language protects the institution from accountability while appearing child-centred. Always RED. This is the clearest application of the scales principle.

Classifier before public release — not after (Debrief 03)
The SEND parent network is fast, tight, and unforgiving. A broken report shared in a Facebook group does more damage than a polished one does good. Classifier scope is fixed: first 3 pages only, JSON output, four format types, one API call per upload, never touches provision content.

Recommendation matching uses a two-term threshold (Debrief 04)
One matching term is too easy. Three or more is too strict. Two is the working balance, with stopwords ensuring the matched terms are substantive provision terms, not noise.

Vagueness and presence are separate checks — never merge them (Debrief 04)
Task 2.3 answers one question: is the substance of this recommendation in Section F at all? Vagueness is checked separately by analyse_section_f(). Merging the two would make findings harder to maintain, harder to test, and harder to explain to a parent.

Child name extracted dynamically, never hardcoded (Debrief 05)
extract_child_name() searches the EHCP text in two passes. The name is added dynamically to _REC_STOPWORDS and stored in st.session_state[‘child_name’]. Without this, the child’s name scores false matches throughout every document.

Two-stage B-to-F matching (Debrief 07)
Stage 1: broad keyword match catches topic presence. Stage 2: PROVISION_LIBRARY precision match checks for named provision strategies. A miss on both fires RED. A match on either confirms provision is present. This prevents ‘communication support will be provided’ from suppressing a gap finding when no named communication provision exists in Section F.

Commitments summary comes before findings (Debrief 07)
build_ehcp_commitments_summary() extracts ‘will’ sentences from Section F, filters administrative language, deduplicates, and displays up to 20 commitments at the top of the report before any findings. The parent first sees what they are owed. Then they see where it is failing. This is perceived value and real value working together. The order changes the emotional experience without changing the analysis.

AI may assist format reading but never provision analysis (Debrief 06)
Vision-based extraction (Approach B for column scrambling) is acceptable because AI reads layout only — furniture, not meaning. The extracted text enters the same deterministic pipeline as today. The analysis layer does not know or care how the text arrived. Every RED finding remains deterministic and citable.

Persistence before beta — revised build order (strategic)
FRED’s real value is layered intelligence — EHCP plus EP report plus correspondence plus history. Without persistence, testers experience a walk-in clinic. The product is a patient record. Build Supabase persistence before testers see FRED.

EHC request template is a second front door (strategic)
Parents who do not yet have an EHCP are more numerous, more frightened, and currently less supported. A guided request template reaches a different parent at the most daunting moment. Parents who use the template and get an EHCP come back to FRED to analyse it. The two features feed each other.

Anonymised pattern data compounds over time (strategic)
Every template submission and EHCP analysis generates anonymised pattern data. Aggregated across hundreds of parents, this becomes the only real-time evidence base for how EHCPs are actually working across England. The database is a product asset, not a byproduct.

SENDIASS lookup is a reusable standalone component (strategic)
Built once, used in multiple places: the request template before submission, the correspondence module when escalation language is detected, the annual review preparation checklist. Every parent FRED works with should know their SENDIASS service exists.



10. Session debrief index

Session
What it covered
Status
Debrief 01
Section F rewrite, VAGUE_MODAL/QUALIFIER, dilution clause, provision library, 6 correspondence patterns, tone scoring
✅ Received
Debrief 02
Section E rewrite, baseline removal, timeframe regex, three outcome categories, scales principle
✅ Received
Debrief 03
Two-column PDF extraction, extract display fix, infrastructure strategy, five decisions identified
✅ Received
Debrief 04
Task 2.3 EP recommendation check, three design decisions, false positive iterations fixed
✅ Received
Debrief 05
Tasks 2.3A and 2.3B — child name extraction, plain English narrative. GitHub web editor established.
✅ Received
Debrief 06
Task 2.3D — sentence-aware extract display, trigger highlighted in red. Tasks 2.4 and 2.5 designed.
✅ Received
Debrief 07
Tasks 2.4 and 2.5 — two-stage B-to-F matching, plain English commitments summary. Perceived value principle established.
✅ Received
Debrief 08
Tasks 2.5A and 1.6A — download parity achieved, column threshold set to 60%, page-break extraction identified
✅ Received
Debrief 09
Supabase infrastructure — project created, six tables, real AUTH, beta bypass, graceful degradation
✅ Received
Debrief 10


Pending




11. Operations one-pager
The things you need if something breaks.

What
Where
Live app
Streamlit Community Cloud — your deployment URL
Code
github.com/smd883-ctrl/fred-beta-Claude
Edit code
Press . on the GitHub repo page
Roll back a change
GitHub → fred_app.py → History → pick a version → Restore
App not restarting
Go to Streamlit Cloud dashboard, click Reboot
Something broke
Come to the overseer conversation. Paste the error. Do not guess.
Section F blocks ≠ 4
Roll back the last commit in GitHub before doing anything else.



FRED — Families’ Rights and Entitlements Directory · Built by a parent. Not a legal service. Plain English findings grounded in the Children and Families Act 2014 and the SEND Code of Practice 2015.
