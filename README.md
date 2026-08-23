# Trade Flow Report Generator

Automatically produces KEPROBA-style "Kenya – <Country> Trade Flow Analysis"
Word reports (`.docx`) from a standard set of seven ITC (International Trade
Centre) Excel files. The generator replicates the layout, styles and narrative
phrasing of the reference report
`KENYA-SAUDI ARABIA TRADE FLOW_Reviewed 1.docx` for any partner country.

The pipeline has two stages:

1. **`make_tables.py`** – builds the six `Table 1 .. Table 6` workbooks from
   the raw ITC downloads (values in USD Thousand are converted to USD
   Billion/Million, rows ranked by the latest year, top-25 exporters for the
   market tables and top-20 products kept, Kenya row inserted/highlighted,
   and "All other …" + total rows added). It also writes
   `Figure 1 Trade Balance.xlsx` derived from the Table 5 / Table 6 totals
   (Kenya's exports minus imports with the partner).
2. **`generate_report.py`** – reads those tables plus the trade-balance file
   and writes the Word report.

The generated tables are live Excel workbooks: every derived figure is an
**Excel formula with a cached result**. The `Share` column is `value / total`,
the `All other …` rows are `total - SUM(top rows)`, the `World`/`All products`
totals are `=SUM(...)`, and the Figure 1 `Balance of Trade` is
`= Exports - Imports`. Because each formula also carries its computed result,
editing a number in Excel recalculates the whole sheet, while the report
generator still reads the current values without needing Excel.

The workbooks also include usability extras: header rows are **frozen**, each
table has Excel **AutoFilter** dropdowns, workbooks **recalc on open**, and
Figure 1 embeds a **line chart** of exports/imports/balance by year. In
addition to the seven individual files, `make_tables.py` writes a combined
**`All Tables.xlsx`** with all seven sheets in one workbook.

## Requirements

- Python 3.9+
- `python-docx`, `openpyxl`, `matplotlib` (install with `pip install -r requirements.txt`)
- Chart font is auto-detected: **Times New Roman on Windows**, Liberation Serif
  (metrically identical) on Linux. No font configuration needed.

```bash
pip install -r requirements.txt
```

**On Windows you don't need to touch a terminal:** the launcher scripts
(`run.bat`, `run_tables.bat`, `run_webapp.bat`) call `install_deps.bat`, which
finds Python (plain `python` or the `py` launcher) and installs anything in
`requirements.txt` that is missing — then runs the task. Just double-click
the launcher; no command prompt needed.

## Input Excel files

### Option A: build the tables from raw ITC downloads

Put the six raw downloads in one folder (the `sourcefiles` folder is
pre-loaded with the Saudi Arabia files). Files are located by keywords:

| Raw file keyword                     | Builds                        | Units       |
|--------------------------------------|-------------------------------|-------------|
| `imports-from-world-by-exporter`     | Table 1 import source markets | USD billion |
| `imports-from-world-by-product`      | Table 2 import products       | USD billion |
| `exports-to-world-by-importer`       | Table 3 export destinations   | USD billion |
| `exports-to-world-by-product`        | Table 4 export products       | USD billion |
| `kenyas-exports-to-<country>-by-product` | Table 5 Kenya exports    | USD million |
| `kenyas-imports-from-<country>-by-product` | Table 6 Kenya imports  | USD million |

```bash
python3 make_tables.py --excel-dir sourcefiles --out-dir output/tables
python3 make_tables.py --excel-dir uae --out-dir output/tables_uae   # UAE files (2021-2024)
```

This also writes `Figure 1 Trade Balance.xlsx` into the output folder, derived
from the Table 5 / Table 6 totals (Kenya's exports minus imports with the
partner country, USD Million). Use that folder as `--excel-dir` for
`generate_report.py`.

### Option B: use ready-made tables

Put the seven Excel files in one folder (the `sample_data` folder is
pre-loaded with the Saudi Arabia files). Files are located by keywords in the
filename:

| File keyword       | Contents                                          | Units         |
|--------------------|---------------------------------------------------|---------------|
| `Table 1`          | Country import source markets (ranking)           | USD billion   |
| `Table 2`          | Country import products (ranking)                 | USD billion   |
| `Table 3`          | Country export destination markets (ranking)      | USD billion   |
| `Table 4`          | Country export products (ranking)                 | USD billion   |
| `Table 5`          | Kenya exports to the country (ranking)            | USD million   |
| `Table 6`          | Kenya imports from the country (ranking)          | USD million   |
| `Figure 1` / `Balance` | Kenya–country bilateral trade balance (annual series) | USD million |

Expected layout inside each workbook: one worksheet containing a header row
with a run of 4 to 10 consecutive years (2021–2025 for Saudi Arabia,
2021–2024 for the UAE files — any consecutive range from 4 to 10 years
works), a code column (header contains "code"), rank rows followed by
trailing `All other …` and `All products`/`World` rows, and a
`Share in <latest year> %` column. The balance sheet has an
`Exports`/`Imports`/`Balance of Trade` label block on the row above the year
series. `generate_report.py` uses the latest available year as the report
year (falling back from `report.year` in the config if it is not present in
the data).

## Kenya Quarterly Performance Report

A second, self-contained pipeline reproduces KEPROBA's quarterly
*Kenya Export Performance* report (`KENYA EXPORT PERFORMANCE IN APRIL-JUNE
2025 -2026.docx`) straight from raw KRA customs extracts:

1. **`make_quarterly_tables.py`** – parses the raw KRA `.xls` extracts
   (`…Exports by HS Destination…` and `…Imports by HS Origin…`; columns
   Year/Month/Destination/HS/SITC/SHORT_DESC/CIF Value in KES), converts
   values to **Ksh. Billion**, drops the `AIRCRAFT & SHIPSTORES` export rows,
   labels blank product descriptions `All others (non-defined)` /
   `All others`, and writes live-formula workbooks `Exports.xlsx`
   (Data, By Partner/Product per year, Table 1 markets, Table 2 products,
   Balance) and `Imports.xlsx` (Data, pivots, Annex 1 partners,
   Annex 2 products).
2. **`generate_quarterly_report.py`** – re-reads those workbooks, derives every
   narrative figure from the data and builds the Word report: letterhead title
   page, TOC + lists of tables/figures (unnumbered), body pages restarting at 1
   with sections Overview → Comparative Perspectives → Lead Destination Markets
   → Composition of Export Commodities → Deduction → Annex.

Every report table (**Tables 1–6**) compares the current quarter with the
same quarter of the previous year:

    Rank | Item | <Y-1> Apr May Jun Total | <Y> Apr May Jun Total |
    Change in <Y-1>-<Y>: Total change | %

The change columns carry green ▲ / red ▼ arrow markers for increases and
decreases. After Table 1 and Table 3 there are two-column exhibits listing
the markets / products that registered growth of over Ksh. 1 billion against
those that declined by over Ksh. 1 billion. Figures 1 and 2 each show **two**
doughnuts stacked vertically on a full A4 page — "SHARE IN <Y>" above
"SHARE IN <Y-1>".

**Single-year mode.** When no prior-year rows are present the comparison is
skipped entirely (like the original sample report): tables collapse to
`Rank | Item | months | Total | Share`, exhibits and year-on-year narrative
clauses disappear, each figure shows a single doughnut, and numbering shifts
down to Tables 1–4.

Command line:

```bash
# raw KRA extracts -> Exports.xlsx / Imports.xlsx (+ report in one go)
python make_quarterly_tables.py --excel-dir "EXPORT PERFORMANCE FOR Q2" --out-dir output/quarterly
python generate_quarterly_report.py --excel-dir output/quarterly --output output/report.docx
```

Several files may be supplied per side — e.g. one extract per year, or a
single file covering both years; they are merged automatically (duplicate or
conflicting rows abort with an error). Add `--synthesize-prev-year` to fake
prior-year data (fixed seed) for testing the comparison layout when only the
current year's extracts exist.

In the web app choose **Quarterly Performance**, then upload the raw KRA
extracts for both years (up to four files: exports + imports for each), a zip
of them, or the generated `Exports.xlsx` + `Imports.xlsx`. With only the
current year's files the report is produced in single-year mode; no synthetic
data is ever generated by the web app.

## Configuration

Each country has a JSON config file in `config/`. Copy
`config_template.json` and fill in the country-specific text. Pre-loaded:
`config/saudi_arabia.json` (complete), `config/uae.json` and
`config/uganda.json` (structure + country identifiers; narrative paragraphs
and quick-facts values are placeholders to fill in). Key fields:

- `country.name`, `country.short`, `country.possessive`, `country.title` –
  used throughout headings, narrative and captions.
- `report.title_line1/2`, `report.year`, `report.month_year` – title page.
- `world_trade.exports/imports` – world totals (USD billion) used to compute
  the "share of world exports/imports" sentences. Omit to skip those clauses.
- `export_potential.image` (assets folder), `export_potential.paragraphs`.
- `map.image` and `map.source`.
- `quick_facts` – list of `[label, value]` pairs for Annex II (fill in
  manually; the `...` placeholders appear in the report until you do).
- `references` – list for the References section.

The Section 1 `background` narrative is **auto-generated** – it is no longer a
config field. The generator follows the reference report's four-paragraph flow
(overview → economy → outlook → trade & policy) for every country so reports
read consistently. Whatever can be computed from the Excel data is written
directly; facts that must be researched externally (GDP, population, IMF
forecasts, WTO membership, etc.) are emitted as `[RESEARCH NEEDED: ...]`
segments highlighted in yellow so they stand out for manual editing.

Assets referenced by config (map image, export-potential image) are resolved
relative to the config file location.

## Usage

### One-word launch

The launcher is a single command, **`tradeflow`**, on both operating systems.
It auto-installs any missing packages, opens your browser and starts the web
app at http://127.0.0.1:8000.

- **Windows:** double-click **`tradeflow.bat`**, or add the Trade Flow folder
  to your `PATH` and type `tradeflow` in any Command Prompt.
- **Linux / macOS:** make it executable and put it on your `PATH`, then run
  `tradeflow` from anywhere:

  ```bash
  chmod +x tradeflow
  echo 'export PATH="$HOME/tradeflow:$PATH"' >> ~/.bashrc   # adjust the path
  source ~/.bashrc
  tradeflow
  ```

  (or symlink it into an existing PATH folder: `ln -s "$PWD/tradeflow" ~/.local/bin/`)

### Web app (manual start)

Start the app and open http://127.0.0.1:8000 in your browser:

- **Windows:** double-click `run_webapp.bat` (it auto-installs any missing
  packages first)
- **Linux / macOS:**

  ```bash
  pip install -r requirements.txt
  python3 -m uvicorn webapp.main:app --host 127.0.0.1 --port 8000
  ```

Pick a country config, choose your files, click **Generate**. The upload can
be the six raw ITC downloads, the seven ready-made tables, or a zip of either
set. The app runs `make_tables.py` (when raw files are given), builds the
report, and offers the `.docx` for download plus a zip of the generated
tables. Generated jobs are stored under `webapp/jobs/` and cleaned up after
24 hours.

**Auto-detecting the country:** the dropdown's first option is *Auto-detect
country from data*. The server reads the country name and report years
straight from Table 1 of the uploaded set: if a matching `config/<country>.json`
exists it is reused, otherwise `make_config.py` generates it automatically
(with `...` placeholders for the facts that still need research). The
config file is written to `config/` so you can edit the placeholders
afterwards.

### Auto-generating a config from the command line

```bash
python3 make_config.py output/tables_uae                # config/united_arab_emirates.json
python3 make_config.py sourcefiles                       # raw ITC downloads work too
python3 make_config.py sourcefiles --output config/mine.json
```

### Windows
Install Python (tick "Add Python to PATH") — no pip step is needed, the
launchers auto-install `requirements.txt`. To build the tables first,
double-click `run_tables.bat` (or run `python make_tables.py --excel-dir
sourcefiles --out-dir output\tables`). Then double-click `run.bat` to
generate the Saudi Arabia report, or run:

```bat
python generate_report.py --excel-dir sample_data --config config\saudi_arabia.json --output "output\KENYA-SAUDI ARABIA TRADE FLOW.docx"
```

### Linux / macOS

```bash
python3 generate_report.py \
    --excel-dir sample_data \
    --config config/saudi_arabia.json \
    --output "output/KENYA-SAUDI ARABIA TRADE FLOW.docx"
```

For the United Arab Emirates (tables built from the `uae` folder into
`output/tables_uae`):

```bash
python3 generate_report.py \
    --excel-dir output/tables_uae \
    --config config/uae.json \
    --output "output/KENYA-UNITED ARAB EMIRATES TRADE FLOW.docx"
```

Options: `--excel-dir` (default `sample_data`), `--config` (default
`config/saudi_arabia.json`), `--output` (default
`output/KENYA-<Name> TRADE FLOW.docx`), `--tmp` (chart scratch directory,
default `output/.tmp`).

`make_tables.py` options: `--excel-dir` (default `sourcefiles`), `--out-dir`
(default `output/tables`), `--top` (rows kept per table, default 20).

## Free hosting (Render.com)

The web app runs on Render's free tier (it sleeps after ~15 min idle and
wakes on the first request, so the first load can take ~1 minute).

1. Make a git repo and push to GitHub:

       git init && git add -A && git commit -m "Initial"
       git remote add origin https://github.com/<you>/tradeflow.git
       git push -u origin main

2. Sign in at https://render.com and choose **New + > Blueprint**.
3. Pick the repo — Render reads `render.yaml` and creates the web service
   automatically, then deploys. Your app is live at
   `https://tradeflow.onrender.com`.

Notes:

- No files are stored permanently: uploads and generated reports live in
  `webapp/jobs/` and are cleaned after 24 h, so nothing sensitive lingers.
- There is no login on the app, so only share the URL with people you trust.
- Charts use Liberation Serif (installed during the build) so figures match
  your local output. If that install is skipped, matplotlib falls back to a
  built-in font and only the chart font changes.

## What is computed automatically

All figures and narrative sentences are derived from the Excel data:

- latest-year totals, year-on-year growth, start-to-end average growth,
  average/max for bilateral exports (over whatever year range the data covers);
- Section 1 "Backgrounds" in the reference four-paragraph flow, with
  data-derived trade facts computed and research-only facts highlighted;
- bilateral imports narrative ("highest/lowest value … before rising by X%")
  from the full balance series — with the 10-year reference balance file this
  reproduces the Reviewed report's "highest in 2018" sentence;
- top-N market/product rankings (top 3 markets, top 4–5 products);
- share-of-world sentences from `world_trade` config values;
- three matplotlib charts: bilateral trade balance (clustered column) and
  doughnut shares of Kenya's top export/import products;
- Table 1–6 rows, Kenya row highlighted, table captions and source lines.

Notes on fidelity to the reference report:

- Layout and wording follow `KENYA-SAUDI ARABIA TRADE FLOW_Reviewed 1.docx`:
  "1.1. Backgrounds", "Compiled by KEPROBA" source lines, the imports
  bullets in section 3.1, "Source: Google map", and split References
  entries.
- `make_tables.py` preserves older balance history: when a `Figure 1 Trade
  Balance.xlsx` already exists in the output folder, its earlier years are
  kept and only the overlapping years are refreshed (so the 2016–2020
  bilateral series survive regenerations).
- Product labels keep their trailing `…` inside tables but are stripped for
  the prose.
- The "Top 10 products accounted for X%" figure is computed as the true
  cumulative share of the top 10 (94.5% for Saudi Arabia, not the 93.7%
  printed in the Reviewed original, which was the top-9 total).
- The lead export destination narrative follows the data (rank 1, currently
  "Area Nes" for Saudi Arabia); the Reviewed report's Section-2 intro instead
  names the UAE, inconsistent with its own Table-3 list.
- In Microsoft Word, the Table of Contents must be refreshed once
  (Right-click > Update Field) after opening the generated document.
# tradeflow
