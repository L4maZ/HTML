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

## Known deviations and cosmetic issues

1. Slide 5 was simplified from a two-series grouped chart (TPV + MAU) to a single
   TPV series plus three callouts. Two series at 30 pt labels could not be read.
   The brief permits simplifying the chart rather than shrinking type.
2. Axis year labels on slides 5 and 6 are abbreviated ("23A" ... "28E") with the
   "fiscal year" qualifier moved into the axis caption. "FY23A" at 30 pt collided
   between adjacent bars.
3. The footer "Illustrative - fictional data" is 60 pt on main slides. Strictly
   compliant, visually heavier than a normal pitchbook footer.
4. Slide 1 mixes typefaces: the template's display serif on "Project Lotus" and a
   sans on the lines added via the API. The API cannot set a font family.
5. Slide 8 use-of-proceeds bars sit directly under each label and read as underlines.

## Slide-to-model reconciliation

All figures trace to the workbook. Two presentational roundings:

| Slide | On slide | Model | Note |
|---|---|---|---|
| 2, 8 | USD 103m net primary | 103.4 (Offer_Structure!B21) | rounded to nearest USD m |
| 6 | "quintuples" | 263.2 / 53.6 = 4.91x | rounded up in prose |
| 4 | USD 15bn | 15.1 (Operating_KPI!E5) | rounded in the title |

Everything else matches to the stated precision.
