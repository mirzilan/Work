# Workflow: Weekly Market Review

End-of-week synthesis: track market moves, sector developments, and competitive intel. Feeds long-term IC paper context and market-intelligence repo.

**Cadence:** Every Friday, 1–2 hours. Archive in `daily/YYYY-WXX-review.md` and cross-file insights into `market_intelligence/`.

## Sections

### 1. Macro & Policy (15 min)
- **Interest rates, FX, commodity moves:** Check Bloomberg or Reuters for major movers.
- **Policy/regulatory news:** Any changes affecting your sectors (energy transition, renewables, infrastructure permitting)?
- **Economic data:** PMI, employment, inflation — any surprises?
- **Action:** 3–5 bullet points max. Link to articles or Bloomberg terminal screenshots if noteworthy.

### 2. Sector Spotlights (30 min)
- Pick 2–3 sectors you track (hydrogen, renewables, utilities, real estate, etc.).
- **What moved:** Stock prices, deal announcements, M&A, capacity additions, policy shifts.
- **Why it matters:** 1–2 sentences on knock-on effects for your pipeline deals or framework assumptions.
- **Action:** Add to `market_intelligence/research/sector-[name].md` if it's a deep observation; otherwise, just log it in the weekly review.

### 3. Competitive & Deal Flow (20 min)
- Scan S&P CapIQ, Bloomberg M&A, Refinitiv for news on peers/competitors.
- Any new entrants, restructuring, or competitive threats?
- Anything that changes your deal-screening assumptions?
- **Action:** Update `market_intelligence/benchmarks/` if multiples or peer list changes.

### 4. Portfolio / Pipeline (15 min)
- **Active deals:** Any material changes to assumptions, timelines, counterparty risk?
- **Prospects:** Did this week's market moves strengthen or weaken your pipeline theses?
- **Exits/completions:** Any deals that closed or were withdrawn?
- **Action:** Note changes to relevant deal folders; flag if IC paper assumptions need refresh.

### 5. Learning & CFA (10 min)
- **What you learned this week:** Any surprising data, misses in your models, or new frameworks?
- **CFA study progress:** If active, any Level III topics that the week's news illuminated?
- **Action:** Brief note for career/cfa-study.md and daily review archive.

## Template

Use this format in your Friday review file (e.g., `daily/2026-W31-review.md`):

```markdown
# Weekly Review: Week of 2026-07-28

## Macro & Policy
- [3–5 bullets on rates, FX, policy, macro data]

## Sectors: [Sector 1], [Sector 2]
- [Sector 1]: [moves + why it matters]
- [Sector 2]: [moves + why it matters]

## Competitive & Deal Flow
- [Key news, peer movements, deal announcements]

## Portfolio & Pipeline
- **Active:** [changes to deal timelines, assumptions, risks]
- **Prospects:** [pipeline shifts from this week's market moves]
- **Exits:** [deals closed or withdrawn]

## Learning & CFA
- [Key insights, CFA study notes]

## Files Updated
- market_intelligence/research/sector-hydrogen.md
- deals/2026-07-HydrogenCorp/assumptions-refresh.md
- career/cfa-study.md
```

## QA Checklist
- [ ] All 5 sections filled (even if briefly)
- [ ] Links to `market_intelligence/` or deal folders if relevant
- [ ] Any assumption changes flagged to affected deal folders
- [ ] CFA connection noted (if studying)
- [ ] Saved to `daily/YYYY-WXX-review.md` before EOD Friday

## Reuse & Archive
- Weekly reviews auto-populate `daily/` for historical tracking.
- Insights → `market_intelligence/` for permanent reference.
- Deal-specific changes → relevant `deals/` folders immediately.

---

**See also:** `workflows/build-ic-paper.md` (uses this market context); `market_intelligence/` folder structure.
