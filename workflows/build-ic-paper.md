# Workflow: Build Investment Committee (IC) Paper

Use this routine to generate a consistent, institutional-grade IC paper for deal screening or investment appraisal.

## Pre-flight

- [ ] Deal folder exists: `deals/YYYY-MM-ProjectName/`
- [ ] Key inputs gathered: term sheet, LoI, preliminary financials, sector comparable companies
- [ ] `.xlsx` model underway or completed (if applicable)
- [ ] Confidentiality: no real client names or terms in draft (use anonymized labels for version control)

## Structure

Follow `frameworks/templates/ic-paper-template.md` as your outline. Key sections:

1. **Executive Summary** (1 page)
   - Investment thesis (2–3 sentences)
   - Recommended decision (Approve/Reject/Further Diligence)
   - Key value drivers and risks (bulleted)
   - Expected IRR / MOIC (if applicable)

2. **Investment Opportunity** (1–2 pages)
   - Market context and sector outlook
   - Target company / project overview
   - Strategic fit rationale

3. **Financial Appraisal** (2–3 pages)
   - Model assumptions (summarized; link to full Excel if external model)
   - Key outputs: Base Case IRR, NPV, MOIC, leverage metrics
   - Stress scenarios: upside/downside case results
   - Sensitivity: which drivers move IRR most?

4. **Risk Assessment** (1 page)
   - Deal risks: execution, market, offtake, regulatory, counterparty
   - Mitigants (contractual, structural, or otherwise)

5. **Recommendation & Next Steps** (0.5 page)
   - Recommended decision
   - Conditions for approval (if applicable)
   - Immediate next steps

## Execution

1. **Draft the summary & thesis** (1 hour)
   - Distill the investment idea to 1 sentence; what makes this a "yes"?
   - Sketch your recommended decision before building financials.

2. **Build/review the financial model** (2–4 hours, depending on complexity)
   - If using Python/openpyxl: generate from `projects/Project-123/src/` or similar.
   - If using existing Excel model: ensure formulas are auditable, color-coded per convention.
   - Extract summary tables: Base Case, Stress, Sensitivity.

3. **Write the appraisal narrative** (2–3 hours)
   - Fill sections 2–4 first; these contextualize the financials.
   - Market data from `market_intelligence/research/` or latest Bloomberg/CapIQ.

4. **QA & polish** (1 hour)
   - Verify all figures in the IC paper reconcile to the model (spot-check 5–10 key numbers).
   - Check formatting: headings, tables, font consistency, logos (if client-facing).
   - Proof for typos and logical flow.

5. **Archive & hand off**
   - Save final PDF to `deals/YYYY-MM-ProjectName/ic-paper-final.pdf`
   - Export Excel model (anonymized if needed) to same folder.
   - Note decision date and approver in a `README.md` within the deal folder.

## Time Budget
- Thesis to finished IC paper: 6–10 hours (straightforward deal) to 20+ hours (complex, multi-scenario).
- Fast-path (screening decision): 4 hours (summary + financials only).

## Reusable Assets
- Template: `frameworks/templates/ic-paper-template.md`
- Screening checklist: `frameworks/checklists/deal-screening.md` (run before committing to IC paper time)
- Financial model engines (if in `projects/`): reuse debt-sculpting, DSCR, 3-statement logic rather than rebuilding.

---

**See also:** `workflows/weekly-market-review.md` for ongoing market tracking to feed IC paper context.
