# SereneHealth Graphics Engine

A local bulk graphics generator designed to reproduce the supplied SereneHealth square graphic system.

## What it does

Feed it a CSV or XLSX content calendar and it creates one finished JPG per row.

The layout is intentionally fixed to the reference system:

- 1254 × 1254 square canvas
- SereneHealth navy/yellow brand palette
- top-left logo and service descriptor
- weekday pill
- large two-line headline with white/yellow emphasis
- hero photo on the right
- supporting slogan / handwritten callout
- three-column service support card
- outcome strip
- dark CTA footer with website, phone and email
- automatic text fitting and wrapping
- automatic hero-image cropping
- bulk generation from a single calendar

## Quick start

1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Put your photos in `assets/`.
5. Put your custom brand fonts in `fonts/` if you have them. The engine will fall back to installed fonts (Lato/DejaVu Sans on many systems).
6. Edit `content_calendar.csv`.
7. Run:

```bash
python serene_engine.py --calendar content_calendar.csv --output output
```

For Excel:

```bash
python serene_engine.py --calendar content_calendar.xlsx --output output
```

## Calendar columns

Required/important columns:

- `date`
- `day`
- `filename`
- `image`
- `headline_white`
- `headline_yellow`
- `subheadline`
- `body`
- `support_title`
- `item1_title`, `item1_desc`, `item1_icon`
- `item2_title`, `item2_desc`, `item2_icon`
- `item3_title`, `item3_desc`, `item3_icon`
- `outcome`
- `handwritten`
- `outcome_script`
- `cta`
- `website`
- `phone`
- `email`
- `right_slogan`
- `badge`

Supported icon names in the three service cards:

`calendar`, `bell`, `people`, `heart`, `growth`, `check`

The `image` column can contain a filename from `assets/` or an absolute path.

## Scaling the operation

For a 30-day calendar, put 30 rows in the spreadsheet and run one command. The engine produces 30 files.

For a monthly batch:

```bash
python serene_engine.py --calendar September.xlsx --output output/september
```

## Important design note

This is a deterministic template engine, not an AI image generator. That is deliberate: deterministic rendering is what keeps the logo, typography, spacing, footer, CTA, and brand layout consistent across hundreds of graphics.

If you want AI-assisted copy generation later, keep the calendar as the contract: AI fills the calendar; this engine renders the approved content.

## Customization

The main design constants are near the top of `serene_engine.py`:

```python
NAVY = "#090A34"
YELLOW = "#FFB82E"
CREAM = "#FFF4D8"
```

The fixed coordinates inside `generate()` control the layout.

For the closest possible match to the supplied design, use the original SereneHealth logo as `assets/serenehealth_logo.png` and use the same brand font files that were used in the original design if available.

## Included

- `serene_engine.py` — generator
- `content_calendar.csv` — working example
- `assets/serenehealth_logo.png` — extracted logo from the supplied reference
- `assets/hero_default.jpg` — extracted sample hero image from the supplied reference
- `requirements.txt`
- `README.md`
