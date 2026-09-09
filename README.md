# SereneHealth Graphics Engine

A local bulk graphics generator that reproduces the supplied SereneHealth square graphic system — now in **v2**, hardened for bulk production.

## What it does

Feed it a CSV or XLSX content calendar and it creates one finished JPG/PNG per row.

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
- bulk generation from a single calendar

## What v2 adds

- **2× supersampled rendering** — every curve, pill, chip and icon is anti-aliased; the navy panel edge is smooth, icons and small UI elements are crisp.
- **Geometry-aware text fitting** — the navy panel curve is modeled as a function, so the headline, subheadline and body copy can *never* spill past the curve. Text shrinks first; if it still can't fit, it wraps (balanced line-breaking) and finally ellipsizes with "…" instead of ever overflowing.
- **Unified headline scale** — both headline lines always share one type size, matching the reference design.
- **Auto-fit everywhere** — day pill, right slogan, handwritten callout, floating badge, outcome strip, outcome script, CTA button and footer contacts all downscale to fit their containers.
- **More icons** — `calendar, bell, people, heart, growth, check, map, clipboard, chat, star, clock, doc, pin, shield` plus footer glyphs `globe, phone, mail`, with friendly aliases (`workflow→map`, `alarm→bell`, `tick→check`, `todo→clipboard`, `hipaa→shield`, …). Unknown names fall back to `check` with a warning; blank cells use the column default silently.
- **Blank-cell intelligence** — required fields (headline, day, CTA, contacts, …) fall back to brand defaults when left empty; optional decorative fields (right slogan, handwritten, badge, outcome script) are skipped cleanly (no orphaned divider bars) when blank.
- **PNG output** — name the file `*.png` in the calendar, or force it with `--format png` (`--format jpg` forces JPEG).
- **QA tooling** — `--validate` fits every row without saving files and prints a warning report; `--only <substring>` renders a subset; per-row errors are isolated so one bad row can't kill a batch; missing hero images and logos degrade gracefully with warnings.
- **Deterministic** — still a template engine, not an AI generator: same calendar + assets + fonts ⇒ pixel-identical output.

## Quick start

1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Put your photos in `assets/`.
5. Brand fonts are already bundled in `fonts/` (Poppins + Dancing Script, SIL OFL — see license files). Drop in your own files there to override.
6. Edit `content_calendar.csv`.
7. Run:

```bash
python serene_engine.py --calendar content_calendar.csv --output output
```

For Excel:

```bash
python serene_engine.py --calendar content_calendar.xlsx --output output
```

Handy flags:

```bash
python serene_engine.py --calendar content_calendar.csv --validate          # QA without rendering
python serene_engine.py --calendar content_calendar.csv --format png        # force PNG output
python serene_engine.py --calendar content_calendar.csv --only workflow     # render a subset of rows
python serene_engine.py --calendar content_calendar_stress_test.csv --validate   # see the warning system
```

## Calendar columns

Required/important columns:

- `date`, `day`, `filename`, `image`
- `headline_white`, `headline_yellow`, `subheadline`, `body`
- `support_title`
- `item1_title`, `item1_desc`, `item1_icon`
- `item2_title`, `item2_desc`, `item2_icon`
- `item3_title`, `item3_desc`, `item3_icon`
- `outcome`, `handwritten`, `outcome_script`
- `cta`, `website`, `phone`, `email`
- `right_slogan`, `badge`
- optional `brand_subtitle` (defaults to the standard descriptor line)

The `filename` column sets the output name; a `.png` extension produces PNG. If `image` is blank the engine uses `assets/hero_default.jpg`.

## Scaling the operation

Put 30 rows in a spreadsheet and run one command to produce 30 files:

```bash
python serene_engine.py --calendar September.xlsx --output output/september
```

Run `--validate` first to catch long copy, missing images or unknown icons before you render the batch.

## Important design note

This is a deterministic template engine, not an AI image generator. That is deliberate: deterministic rendering is what keeps the logo, typography, spacing, footer, CTA and brand layout consistent across hundreds of graphics.

If you want AI-assisted copy later, keep the calendar as the contract: AI fills the calendar; this engine renders the approved content — and v2's fit system guarantees the layout survives whatever copy it is fed.

## Customization

The main design constants are near the top of `serene_engine.py`:

```python
NAVY = "#090A34"
YELLOW = "#FFB82E"
CREAM = "#FFF4D8"
```

The navy panel's curve is defined by `PANEL_PTS`; text fitting automatically respects whatever shape you set there. Fixed layout coordinates live inside `generate()`.

## Included

- `serene_engine.py` — generator (v2)
- `content_calendar.csv` — working example
- `content_calendar_stress_test.csv` — edge-case fixture demonstrating shrink/warn behavior
- `assets/serenehealth_logo.png` — extracted logo from the supplied reference
- `assets/hero_default.jpg` — extracted sample hero image from the supplied reference
- `fonts/` — Poppins + Dancing Script (SIL OFL, license files included)
- `requirements.txt`, `README.md`
