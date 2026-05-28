FRED — Session Debrief 01
Families' Rights and Entitlements Directory
Intelligence layer, decisions, and reasoning
---
What this document is
FRED is built session by session. Each session produces working code — but the reasoning behind each decision, the legal standard it is grounded in, and the thinking that shaped it lives only in the chat. This document captures that reasoning for the first major development session: the Section F language check rewrite, the provision library, and the correspondence intelligence layer.
This is not a technical manual. It is an explanation of why FRED works the way it does — written for the person who built it, so that the intelligence behind the tool is not locked inside code that only makes sense if you were there.
---
1. The Section F language check — what changed and why
1.1 The original problem
The original FRED had a single list called PROHIBITED_WORDS. It contained a mix of terms — some weak modal verbs, some vague qualifiers — and produced one generic finding: 'weak or unenforceable language detected'. The finding was technically correct but legally shallow. It told a parent something was wrong without explaining what kind of wrong, or why that specific kind of language mattered.
The deeper problem was that PROHIBITED_WORDS was based on intuition, not on a documented legal standard. When challenged — by a school, a SENCO, or a tribunal — a parent couldn't point to why 'would benefit from' was a problem. FRED needed to be grounded in something citable.
1.2 The legal standard
The SOS!SEN EHCP Infosheet (2022) provides an explicit list of terms that must be challenged in Section F. These include: 'will benefit from', 'access to', 'opportunities are', 'regular', 'up to', 'as advised', 'as required', 'may be helpful', 'contacts', and 'adults'. The reasoning is that each of these makes provision impossible to audit or enforce.
The SEN Code of Practice (paragraph 9.69) requires Section F to be 'detailed and specific and normally quantified'. The Children and Families Act 2014 creates an absolute duty on the LA to secure what is specified in Section F. The enforcement route — Judicial Review — is only available if the provision is specific enough to be audited. Vague language removes enforceability entirely.
1.3 The decision — two lists, two findings
Rather than one list, FRED now has two:
VAGUE_MODAL — terms that weaken the commitment itself. 'Would benefit from', 'should receive', 'may be provided'. These replace 'will receive' — the only formulation that creates a lawful duty — with something conditional or aspirational. This is a commitment failure.
VAGUE_QUALIFIER — terms that state provision without specificity. 'Access to', 'as appropriate', 'regular', 'up to'. These make provision impossible to audit. You cannot check whether 'regular' sessions are happening. You cannot enforce 'up to X hours' because it allows zero delivery. This is a specificity failure.
The distinction matters because the remedies are different. A commitment failure requires the sentence to be rewritten with 'will receive'. A specificity failure requires the provision to be quantified — frequency, duration, deliverer. A parent who understands this can have a much more precise conversation at annual review.
1.4 Showing the actual EHCP sentence
The original finding showed the trigger word. The new finding shows the actual sentence from the EHCP in which the trigger word appeared. This is a small technical change with significant practical impact. A parent can now take the report to an annual review and point to a specific sentence in the document and say: this sentence is the problem, and here is why.
---
2. The dilution clause — tightening the pattern
The original dilution clause pattern fired on single words: 'resources', 'staffing', 'capacity', 'availability'. This produced false positives on legitimate provision language — 'sensory resources', 'staffing arrangements', 'building capacity'. A finding that fires on 'sensory resources' is not just unhelpful — it actively undermines trust in the tool.
The rewritten pattern only fires on conditional constructions — phrases where a word like 'resources' or 'staffing' appears inside a clause that makes provision conditional: 'subject to staffing', 'where resources allow', 'capacity permitting'. This is the legally relevant pattern. An EHCP that says 'sensory resources will be provided' is compliant. An EHCP that says 'sensory resources will be provided subject to staffing' is not.
The legal basis: the duty to deliver what is specified in Section F is absolute under s42 Children and Families Act 2014. It does not depend on the school's budget. If additional funding is required, the obligation passes to the LA — not to the school to decide whether it can afford it.
---
3. The provision library — why it exists
PROVISION_LIBRARY is a structured dictionary of known, legitimate SEND provision terms organised into four categories that match the four need areas in Section B of an EHCP: Communication and Interaction, Sensory and Physical, SEMH, and Cognition and Learning.
It serves three purposes. First, it powers the B-to-F cross-reference check — FRED can look for provision terms in Section F that correspond to needs identified in Section B, and flag gaps where a need is described but no provision is named. Second, it powers the provision inventory — a plain English summary of what the EHCP commits to. Third, it powers the correspondence green card — when a school uses correct SEND language in an email, FRED can recognise it and prompt the parent to check whether that language is backed by a delivery log.
The library also establishes a standard. When a school says 'emotional support will be provided', that is not a provision term — it is a category. When a school says 'Zones of Regulation check-in three times per week with the SEMH TA', that is a provision term. The library contains the latter, not the former.
---
4. The correspondence intelligence layer
4.1 What the correspondence module does
The correspondence module reads emails and letters from schools and LAs and identifies recognised patterns — ways of writing that have consistent meaning in the context of SEND disputes. It is not sentiment analysis. It is pattern recognition based on documented experience of how institutions communicate when they are managing risk rather than meeting need.
4.2 The six new patterns — reasoning
The Process Substitution — RED
When a school shifts from SEND provision language to operational and policy management language — 'graduated response', 'quality first teaching', 'universal provision' — it is describing what it does for all pupils, not what it is doing for this child. A child with an EHCP is entitled to provision above and beyond universal classroom practice. Substituting process language for provision language is a way of appearing to respond without committing to anything specific.
The Escalation Sequence — RED
Escalation language — referrals, panels, risk assessments, managed moves — creates a paper trail that justifies later formal action. Each step has statutory implications and requires parental involvement. When this language appears in correspondence it signals that the school is preparing a formal process, not resolving a provision question. The parent needs to know this immediately.
The Collaborative Mask — AMBER
Warm collaborative language — partnership, co-production, open dialogue — is not itself a problem. The concern is when it appears alongside escalation steps or provision substitution. A letter that opens with 'we are committed to working in partnership' and closes with a referral to a panel is not a collaborative document. The warm tone is amber, not red, because the language alone is not harmful — it is the combination that matters, and the tone scoring handles that.
The Retrospective Justification — RED
When a school presents a later reconstruction of events as contemporaneous evidence — 'records indicate prior patterns', 'evidence gathered over time' — it is assembling a case after the fact. Contemporaneous records written at the time carry significantly more evidential weight. If a pattern was identified over time, the question is: when was it first recorded and what action was taken? If no contemporaneous record exists, that is itself significant.
The Provision Substitution — RED
When a school describes provision in terms of what is operationally convenient — 'within available resources', 'shared teaching assistant', 'flexible delivery model' — it is substituting its staffing model for the named, quantified provision required by Section F. The duty is absolute. If the school cannot deliver what Section F specifies, the obligation passes to the LA, not to the parent to accept a lesser version.
Correct SEND Language Used — GREEN
When a school uses specific SEND provision language in correspondence — 'sensory diet', 'trusted adult', 'Zones of Regulation', 'co-regulation' — it is a positive signal. The green card acknowledges this while immediately redirecting the parent to the delivery question: is this language in Section F as a specified commitment, and is there a delivery log? Language in an email is not a lawful commitment. Language plus a delivery log is the standard.
The decision to make this green rather than amber was deliberate. FRED should reward correct practice. If the tool only ever produces red and amber findings, it becomes an adversarial instrument. Green cards make FRED a fair witness.
---
5. The tone scoring — what changed and why
The original tone scoring counted two signals: formal language and collaborative language. The new scoring adds two additional signal sets:
Escalation signals — each match adds two points to the formal score. Escalation language always pushes toward a formal response regardless of tone.
Operational SEND signals — correct provision language adds one point to the collaborative score.
The collaborative mask is now explicitly detected: if collaborative language count is two or more AND escalation count is one or more, the recommendation is formal with a specific explanation — 'collaborative mask detected'.
---
6. The core philosophy — never break this
Every decision in this session was made against four principles:
Explicit failure over silent failure. Wrong output with no warning is worse than an honest error.
Grounded in the legal standard. Every RED finding must be citable. SOS!SEN, CoP para 9.69, Children and Families Act 2014.
Plain English output. The analysis is deterministic. The output is for a parent who may never have read the legislation.
Stability before sophistication. No change to the parser without flagging for review. One change at a time, tested before the next.
---
7. What comes next
Section E timeframe regex — 'by end of Key Stage X' not currently recognised
Two-column PDF extraction — Warwickshire format scrambles provision text
EP recommendation laundering check — find 'recommended' in appendix, check it appears as 'will' in Section F
B-to-F matching enhancement — use PROVISION_LIBRARY categories for richer cross-reference
---
FRED — Families' Rights and Entitlements Directory · Session Debrief 01 · Built by a parent. Not a legal service. Plain English findings.