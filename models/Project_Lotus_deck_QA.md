# Project Lotus — IPO advisory pitch deck (fictional)

Canva design ID: `DAHViAsasX0`
Open: https://www.canva.com/design/DAHViAsasX0/edit
Model: `models/NovaPay_IPO_Model_Illustrative.xlsx`

Everything in the deck and the model is invented for an illustrative pitch exercise.
No real issuer, bank, exchange or transaction is referenced. Regulatory points are
described generically as "per prevailing listing regulations". HOSE is named only as
the listing venue.

## Unit convention

Canvas is 1920x1080 px. Canva exports a 16:9 presentation to a 13.333 in slide,
so 1920 px / 13.333 in = 144 px/in and **1 pt = 2 px**. The Kawasaki 30 pt floor is
therefore **60 px** on this canvas. Every main-slide text element is set to >= 60 px.

## Type scale

| Role | px | pt |
|---|---|---|
| Action title | 76 | 38 |
| Column heading | 68 | 34 |
| Hero figure | 78 | 39 |
| Body / chart label / footer (main slides) | 60 | 30 |
| Appendix table body | 42-44 | 21-22 |
| Appendix caption / footer | 38 | 19 |

## Palette

- Primary navy `#1F3864`
- Accent terracotta `#C3542B`
- Ink `#1A202C`, muted `#5A6478`
- Background `#FFF8F6` (inherited from the generated template)

Navy denotes actuals, terracotta denotes forecasts, consistently across all charts.

## 10/20/30 compliance

- 10 main slides + 4 appendix slides.
- Minimum type on every main slide: 60 px = 30 pt. No exceptions.
- Speaker notes total 2,510 words -> ~17.9 min at 140 wpm, ~19.3 min at 130 wpm.

## Known deviations

1. Slide 5 was simplified from a two-series grouped chart (TPV + MAU) to a single
   TPV series plus three callouts. Two series at 30 pt labels could not be read.
   The brief permits simplifying the chart rather than shrinking type.
2. Axis year labels on slides 5 and 6 are abbreviated ("23A" ... "28E") with the
   "fiscal year" qualifier moved into the axis caption. "FY23A" at 30 pt collided
   between adjacent bars.
3. The footer stays at 60 px (30 pt) on main slides because the rule requires it.
   It is set in a light grey (#A8B0BF) so it recedes without breaking compliance.
   Dropping it to a conventional 12 pt footer would be the only way to make it
   visually typical, and that would be a 10/20/30 violation.

## Cosmetic issues found in QA and since fixed

1. Slide 1 mixed two typefaces and carried a brand-style colour override that
   `format_text` could not beat. The cover text was rebuilt with `add_text` so the
   whole deck now uses one typeface and the intended navy/terracotta.
2. Slide 8 use-of-proceeds bars sat flush under each label and read as underlines.
   Bars were moved to a 16 px gap above and below, thickened to 26 px and recoloured
   terracotta so they read as a bar chart.
3. Slide 10 accent rules sat against the second line of the title. Rules moved to
   y=330 and both columns re-spaced on a 160 px row pitch.
4. Footers on all ten main slides lightened from #5A6478 to #A8B0BF.

## Slide-to-model reconciliation

All figures trace to the workbook. Two presentational roundings:

| Slide | On slide | Model | Note |
|---|---|---|---|
| 2, 8 | USD 103m net primary | 103.4 (Offer_Structure!B21) | rounded to nearest USD m |
| 6 | "quintuples" | 263.2 / 53.6 = 4.91x | rounded up in prose |
| 4 | USD 15bn | 15.1 (Operating_KPI!E5) | rounded in the title |

Everything else matches to the stated precision.
