FRED — Session Debrief 02
Families' Rights and Entitlements Directory
Section E rewrite — outcomes, achievement language, and vague aspiration
---
What this session covered
This session completed the Section E rewrite — the outcomes section of the EHCP. The original analyse_section_e function had three problems: it was checking for baselines (which don't belong in the EHCP), its timeframe regex missed common EHCP phrasings like 'by the end of the Key Stage 3', and it had no check for the most common Section E failure — outcomes written as activities or vague aspirations rather than measurable achievements.
The session also produced a key insight about who benefits from vague language in an EHCP — an insight that is now embedded in the commentary of FRED's vague aspiration finding and forms part of the tool's philosophical foundation.
---
1. The baseline check — why it was removed
The original Section E analysis fired an AMBER finding when no baseline was present in the outcomes section. This was incorrect on two levels.
First, the legal standard. Baselines — where the child is now, before intervention — belong in the assessment reports that form the appendices to the EHCP: the EP report, the SALT report, the OT report. The EHCP itself is not an assessment document. It is a commitment document. Section E should contain outcomes — where the child will be — not starting points.
Second, the practical effect. Firing an AMBER for a missing baseline was penalising EHCPs for being correctly structured. A parent reading that finding might ask their LA to add baselines to Section E — which would make the document less legally precise, not more. The finding was actively misleading.
The baseline check was removed entirely. The FRED session starter already documented this correctly: 'Do NOT check for baselines — baselines live in reports, not in the EHCP.' The code had drifted from the specification.
---
2. The timeframe regex — what it needed to recognise
The original timeframe check only recognised two patterns: 'by [month] 20XX' and 'within X months/weeks/terms/years'. A Warwickshire EHCP uses 'by the end of the Key Stage 3' — which failed to match because of the second 'the' before 'Key Stage'. The finding fired incorrectly, telling a parent their EHCP lacked a timeframe when it was clearly present in the document.
The extended regex now recognises all EHCP-relevant timeframe formats:
Key Stage references — 'by the end of Key Stage 2', 'by the end of the Key Stage 3'
Year group references — 'by the end of Year 6', 'by Year 9'
Annual review references — 'by the next annual review', 'at the annual review'
Age references — 'by the age of 11', 'by age 16'
Transition references — 'before transition to secondary', 'before starting college'
Phase references — 'by the end of primary', 'before leaving secondary'
Dated timeframes — 'by July 2026' — already working, retained
The scope is deliberately limited to EHCP-relevant timeframes. APDR targets, half-termly school targets, and local authority planning documents use different timeframe language that should not be in Section E of the EHCP.
---
3. Three categories of outcome language
The session established that Section E outcome language falls into three distinct categories, each with a different legal and practical implication. FRED now detects all three.
Achievement language — correct
Outcomes written as capabilities: 'will be able to', 'will independently', 'will demonstrate', 'will initiate'. These describe what the child will be able to do at the end of a period of provision. They can be assessed at annual review against a clear standard.
Example of a good outcome:
> "Freddie will be able to spend 20 minutes in a group of three peers and initiate conversation on two occasions without adult prompting."
That outcome has a capability, a quantity, a context, and a quality indicator. It can be passed or failed at annual review.
Activity language — RED if no achievement language present, AMBER if mixed
Outcomes written as services received: 'will receive', 'will attend', 'will be supported', 'will be provided with'. These describe provision — what the child will get — not what they will be able to do as a result. Activity language in Section E means provision has been written into the wrong section. It belongs in Section F.
Vague aspiration language — always RED
Outcomes written as directions of travel: 'will have developed', 'will have improved', 'will have made progress', 'will have gained'. These are the most damaging category because they appear child-centred while being entirely unaccountable.
Example of vague aspiration language:
> "By the end of Key Stage 3, Freddie will have developed his social understanding, self-awareness, interaction and communication skills so that he is able to participate in school activities."
This outcome cannot be failed at annual review. The school can always point to some development in social understanding. The outcome protects the institution from accountability while appearing ambitious. It is always RED in FRED.
---
4. Who benefits from vagueness — the key insight
This session produced the clearest articulation of FRED's philosophical foundation. When asked about vague aspiration language, the response from the builder was direct:
> "I think lack of specificity is red — it allows the school to fail but claim vague progress. This does not help the child, just the school. It's like a set of scales — they should always tip to child and family. Who gains from vagueness?"
The answer is: the school. Every time.
Vague outcome language is a risk management tool, not a child development tool. 'Will have developed his social understanding' can never be failed at annual review. The school can always say some progress has been made.
A specific outcome — 'will be able to spend 20 minutes in a group of three peers and initiate conversation on two occasions without adult prompting' — can be failed. The school knows at the start of the year exactly what they are being held to. That specificity is uncomfortable for schools and essential for children.
This insight is now embedded in the commentary of FRED's vague aspiration finding. Every parent who reads that finding will see: 'This language benefits the institution, not the child.'
The scales principle — every decision in FRED should tip toward the child and family — is now a named part of FRED's design philosophy. Applied as a test to every future finding: if this finding fires, who does it protect? If the answer is the institution, the finding is correct.
---
5. The measurability check — why it's AMBER not RED
The Section E analysis includes a check for measurable success criteria — specific numbers, independence indicators, setting criteria. This fires as AMBER, not RED.
The absence of a measurable criterion in an outcome is a best practice gap, not a lawful failure. The CoP requires outcomes to be SMART — including measurable — but an outcome without a specific number is not automatically unenforceable in the same way that a vague aspiration is. A timeframe is required by law. A measurability criterion is required by good practice.
The distinction matters because FRED must calibrate its severity correctly. If everything is RED, nothing is RED.
---
6. Test results — Warwickshire EHCP
The Warwickshire EHCP produced the following findings after the full Section E rewrite and timeframe fix:
RED — Section F uses conditional language instead of a lawful commitment (actual EHCP sentences shown)
RED — Section F contains vague qualifier language that removes enforceability (actual EHCP sentences shown)
RED — Section E outcomes use vague aspiration language — unmeasurable at annual review (actual EHCP sentence shown)
AMBER — No delivery log mechanism referenced
3 Red, 1 Amber, 0 Green. This is an accurate picture of the Warwickshire EHCP. The timeframe false positive — which was incorrectly firing on a document that did contain 'by the end of the Key Stage 3' — has been resolved.
---
7. What comes next
Two-column PDF extraction — Warwickshire format scrambles provision text in two-column tables. Needs pdfplumber x-coordinate column separation.
EP recommendation laundering check — find 'recommended' in appendix reports, check whether it appears as 'will' in Section F. Not yet built.
B-to-F matching enhancement — use PROVISION_LIBRARY categories for richer cross-reference between Section B needs and Section F provision.
The scales principle applies to all three. The two-column fix ensures provision text isn't scrambled — a scrambled extraction could miss a vague language finding, which benefits the LA not the parent. The laundering check catches the most common way LAs avoid committing to provision. The B-to-F matching catches needs that have no provision at all — the most serious gap of all.
---
FRED — Families' Rights and Entitlements Directory · Session Debrief 02 · Built by a parent. Not a legal service. Plain English findings.