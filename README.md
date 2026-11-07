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

- Python 3.10+
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
pre-loaded with the Saudi Arabia files). Files are located by keywords —
both Trade Map download styles are accepted, the **beta** interface slugs
and the **previous** interface names (`Trade_Map_-_List_of_supplying_
markets_for_a_product_imported_by_<Partner>.xls`,
`Trade_Map_-_List_of_products_imported_by_<Partner>.xls`,
`Trade_Map_-_List_of_importing_markets_for_a_product_exported_by_<Partner>.xls`,
`Trade_Map_-_List_of_products_exported_by_<Partner>.xls` and the two
`Trade_Map_-_Bilateral_trade_between_Kenya_and_<Partner>` workbooks, which
are told apart by their content):

| Raw file keyword (beta)              | Classic "previous" download            | Builds                        | Units       |
|--------------------------------------|----------------------------------------|-------------------------------|-------------|
| `imports-from-world-by-exporter`     | List of supplying markets … imported   | Table 1 import source markets | USD billion |
| `imports-from-world-by-product`      | List of products imported              | Table 2 import products       | USD billion |
| `exports-to-world-by-importer`       | List of importing markets … exported   | Table 3 export destinations   | USD billion |
| `exports-to-world-by-product`        | List of products exported              | Table 4 export products       | USD billion |
| `kenyas-exports-to-<country>-by-product` | Bilateral trade (Kenya's exports)  | Table 5 Kenya exports         | USD million |
| `kenyas-imports-from-<country>-by-product` | Bilateral trade (Kenya's imports)| Table 6 Kenya imports         | USD million |

All tables in one report share a common year set: with mixed coverage the
overlap is used (e.g. classic bilateral files covering 2023–2025 yield
three-year tables).

```bash
python3 make_tables.py --excel-dir sourcefiles --out-dir output/tables
python3 make_tables.py --excel-dir uae --out-dir output/tables_uae   # UAE files (2021-2024)
```

This also writes `Figure 1 Trade Balance.xlsx` into the output folder, derived
from the Table 5 / Table 6 totals (Kenya's exports minus imports with the
partner country, USD Million). Use that folder as `--excel-dir` for
`generate_report.py`.

### Web app: specify the partner & auto-prepare the ITC download checklist

In the **Goods Trade Flow** page there is a **Partner** box (e.g. `Saudi
Arabia`, `Malaysia`). Pressing **Prep checklist** calls
`GET /api/goods/partner-setup?partner=<name>` which:

- **reuses** the partner's `config/<slug>.json` if one exists, otherwise
  **auto-creates it from the name** (country/report titles) using
  `make_config.build_config_from_name`;
- **pre-fills the Quick Facts** as `...` placeholders, to be researched or
  edited in the generated config before the report is built;
- returns the exact **six ITC Trade Map queries** to run plus the file-name
  patterns (`table1`–`table6`) that upload auto-detection expects.

Note: ITC Trade Map has **no public API** (it is an Angular SPA), so the
system cannot download the files itself. After you download the six workbooks
from trademap.org, upload them together — the partner country is auto-detected
from Table 1 and the matching config reused automatically.

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
with a run of 2 to 10 consecutive years (2021–2025 for Saudi Arabia,
2021–2024 for the UAE files — any consecutive range from 2 to 10 years
works, including the short 2023–2025 windows the previous Trade Map
interface produces), a code column (header contains "code"), rank rows followed by
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

**Background processing:** the heavy build runs server-side in a background
job — the browser receives an immediate acknowledgement and follows live
progress by polling, so even multi-minute quarterly builds (large KRA
extracts) or slow hosts no longer kill the request mid-way. If the server
does restart during a job, the page says so instead of hanging.

**Preview, quality, sharing and comments.** A completed job shows workbook
previews, a data-quality report, and an optional comment thread. The
**Recent jobs** panel lists jobs created by the current browser; each job's
owner token is kept in `localStorage` under `tf_job_store_v1`. A share link
opens a read-only view of the job, expires after seven days by default, and is
signed with the server secret.

New jobs are owner-token protected by default: API, status, download, preview
and quality requests need either `X-Job-Token: <token>` or
`?token=<token>`, while shared links use `?share=<signed-token>`. Legacy job
folders with no registered owner remain readable. Set
`TRADEFLOW_ENFORCE_JOB_OWNERSHIP=0` only on a trusted single-user deployment
that needs bare bookmarkable job URLs.

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

### Open-data sources

The **Or fetch open data** card appears on the Trade Brief, Product Profile and
Services Brief tabs, and lists only the sources that tab can use. Fetched files
are kept for six hours (configurable) and can be downloaded individually or as
a ZIP. A report-producing fetch feeds a new job directly; a reference fetch is
installed instead.

| Source | Produces | Credentials | Notes |
|---|---|---|---|
| **UNCTADstat** | Services Brief source workbooks | `UNCTAD_CLIENT_ID`, `UNCTAD_API_KEY` | UNCTAD-WTO trade in services by EBOPS category. Builds the six Trade Map-style services workbooks. Services tab only. |
| **KNBS reference series** | `knbs_reference.json` + `knbs_observations.csv` | none | Rebuilds the Economic Survey snapshot every report cross-checks against, from the portal's Hub-and-Node API (about 12 s for 14 datasets). The card lists the reports and tables that read it before you fetch. **Install snapshot** replaces `data/knbs_reference.json` and drops the in-process cache, so the next build reads the new vintage. Not a report input. |

No goods source is offered. The UN Comtrade preview and World Bank WITS cannot
return the partner-market HS4 tables the Trade Brief needs, so their files were
unusable as report inputs and those connectors are not registered;
`webapp/connectors/__init__.py` lists what the app exposes. ITC Trade Map has no
public API either, so goods, product and quarterly data must still be
downloaded manually from trademap.org and uploaded. `GET /api/sources` reports
the catalog's parameters, credential readiness and, for a reference source, the
`used_by` list of reports and tables that read it; the browser disables sources
whose server credentials are missing.

### Platform endpoints and operations

Health and operations:

| Endpoint | Purpose |
|---|---|
| `GET /healthz` | Liveness, version, queue depth and job counts. |
| `GET /metrics` | Prometheus text metrics for HTTP, jobs and queue depth. |
| `GET /api/jobs?limit=25` | Recent jobs visible to the caller. |

Job inspection and collaboration:

| Endpoint | Purpose |
|---|---|
| `GET /api/preview/{job_id}` | Table/chart preview data for a completed job. |
| `GET /api/quality/{job_id}` | Input/output validation report for a completed job. |
| `POST /api/share/{job_id}` | Create a signed read-only share URL (owner only). |
| `GET /api/jobs/{job_id}/comments` | List comments. |
| `POST /api/jobs/{job_id}/comments` | Add a comment (`{"body": "...", "author": "..."}`; owner only). |

Open-data fetch flow:

| Endpoint | Purpose |
|---|---|
| `GET /api/sources` | Connector catalogue, parameters, credential readiness and `used_by`. |
| `POST /api/sources/{source_id}/fetch` | Fetch data; body is the connector's parameter object. |
| `GET /api/fetches/{fetch_id}/files/{name}` | Download one fetched file. |
| `GET /api/fetches/{fetch_id}/download` | Download all fetched files as a ZIP. |
| `POST /api/fetches/{fetch_id}/run` | Start a report from a report-producing fetch (`{"report_type": "goods", ...}`). A reference-only source is refused whatever type is asked for. |
| `POST /api/fetches/{fetch_id}/install` | Replace `data/knbs_reference.json` with a fetched KNBS snapshot. Refuses a snapshot with no observations, leaving the installed series untouched. |

If `TRADEFLOW_API_KEY` is set, clients must also send `X-API-Key` (or
`?api_key=`) on job-list and fetch/run endpoints. Owner/share tokens continue
to control access to individual jobs.

### Environment configuration

All settings are optional; the defaults preserve local behaviour. Set them as
normal environment variables or in your hosting provider's dashboard.

| Variable | Default | Meaning |
|---|---|---|
| `TRADEFLOW_DATA_DIR` | `webapp/` | Root for the registry, fetches, cache and generated secret. |
| `TRADEFLOW_DB` | `<data>/platform.db` | SQLite registry path. |
| `TRADEFLOW_JOBS_DIR` | `webapp/jobs/` | Per-job uploads, reports and tables. |
| `TRADEFLOW_FETCH_DIR` | `<data>/fetches/` | Staged source downloads. |
| `TRADEFLOW_CACHE_DIR` | `<data>/cache/` | Connector/reference-data cache. |
| `TRADEFLOW_REPORT_VERSION` | `2026.1` | Version stamp written into job provenance. |
| `TRADEFLOW_JOB_TTL` | `86400` | Job retention in seconds (24 hours). |
| `TRADEFLOW_FETCH_TTL` | `21600` | Fetch retention in seconds (6 hours). |
| `TRADEFLOW_WORKERS` | `1` | Background queue worker count. |
| `TRADEFLOW_JOB_ATTEMPTS` | `2` | Attempts before a job is marked failed. |
| `TRADEFLOW_SHARE_TTL` | `604800` | Share-link lifetime in seconds (7 days). |
| `TRADEFLOW_API_KEY` | unset | Optional API gate for job-list and fetch/run endpoints. |
| `TRADEFLOW_ENFORCE_JOB_OWNERSHIP` | `1` | Require owner tokens for registered jobs. |
| `TRADEFLOW_WEBHOOK_URL` | unset | Optional completion webhook (POST JSON). |
| `TRADEFLOW_WEBHOOK_TIMEOUT` | `5` | Webhook timeout in seconds. |
| `TRADEFLOW_SECRET` | generated | HMAC secret for share tokens. Set it in hosting to keep links valid across restarts. |
| `UNCTAD_CLIENT_ID` / `UNCTAD_API_KEY` | unset | UNCTADstat credentials. |
| `COMTRADE_API_KEY` | unset | UN Comtrade subscription key. |

The generated secret is stored at `<data>/.platform_secret`. If the data
directory is read-only, an ephemeral secret is used and existing share links
stop working after a restart. Job ownership is enforced only for jobs created
by this version; old job folders without an owner hash remain readable.

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

## KNBS reference series (cross-checks, regions, one year ahead)

Kenya National Bureau of Statistics publishes the machine-readable equivalent
of Economic Survey chapter 6 (external sector) through the Hub-and-Node API
behind [opendata.knbs.or.ke](https://opendata.knbs.or.ke/beta/databrowser/).
The values are in Kenyan shillings, so they are an independent check on the US
dollar figures ITC and UNCTAD give us, and they arrive about a year earlier
than the international databases.

| Module | Does |
|---|---|
| `knbs_reference.py` | Portal client, the unit-scale registry and the annual average KSh/USD table. `python knbs_reference.py refresh` rewrites the vendored snapshot in `data/knbs_reference.json`; the web app's KNBS connector does the same and installs it on request. |
| `knbs_check.py` | Cross-checks. Each returns one shape (`label`, `rows` of `{year, ours, knbs, diff_pct, provisional}`, `status`, `source`) that any report can render. |
| `knbs_regions.py` | KNBS's own region tree (continents, sub-regions, the EU and EAC as "of which" blocs) and one year of a cube turned into report rows. |
| `knbs_sector.py` | The industry behind a product: area, output, yield and the farm-gate price, from the Economic Survey's sector chapters, as rows a profile can print. |

What a reader sees:

- **Trade Brief** — `Table 4a` sets Figure 1's bilateral exports and imports
  beside KNBS's customs line for the same partner, and section 3.6 adds
  `Table 7a`, Kenya's exports by region, which the country list cannot show.
- **Services Brief** — Figure 12 continues into KNBS's provisional Balance of
  Payments year (2025) that UNCTAD's EBOPS series lacks, and `Table 13a` at the
  end of section 5 prints the gap between the two sources. Category detail in
  Tables 3 and 4 stays on UNCTAD's EBOPS series, which ends a year earlier.
- **Product Profile** — a profile scoped to one 4-digit heading that KNBS
  publishes separately (cut flowers 0603, coffee 0901, tea 0902, essential
  oils 3301) gets `Table 12a`. KNBS counts *domestic* exports, so a shortfall
  that repeats every year is re-exported volume rather than an error. When the
  portfolio also holds the "List of products exported by Kenya" workbook, the
  same table carries a second row for Kenya's total merchandise exports. A
  profile of a whole family (several headings) prints no check, because KNBS
  has no single line that means the same thing.
- **Product Profile, sectoral context** — `Table 12b` (or `12a` when there is
  no customs check) sets the exports beside KNBS's account of the industry
  itself: hectares under the crop, tonnes picked, yield and the farm-gate price
  for coffee and tea; volumes and earnings for horticulture; output and prices
  for minerals; live weights and prices for meat.  It covers the five years the
  Survey publishes cleanly, 2021-2025, and only for the headings KNBS keeps such
  a series for; a dairy or crafts profile prints nothing.  Coffee rows are crop
  years, so the 2025 column is the 2024/25 season, and a cell KNBS itself
  publishes inconsistently prints `n/a` with a note rather than a wrong number.

House rules the layer keeps to: a report prints nothing rather than a zero, a
year KNBS marks provisional carries a `*`, gaps above `REVIEW_THRESHOLD_PCT`
(2%) are bold red with a note asking for a review, and a missing or stale
snapshot leaves every report exactly as it was before the layer existed.
Partner names are matched to KNBS's own labels, so a market KNBS lumps into
"Other Countries in Africa" (Nigeria, for instance) produces no check at all.

## Competitor analysis (per product or per Kenya export market)

`generate_competitor_analysis.py` reads ITC Trade Map "List of supplying
markets for a product imported by <market>" downloads and writes a
paste-ready `.png`, an editable `.docx` and an `.xlsx`. The layout is detected
from the files:

- **per product** - one importing market, several HS products (e.g. the
  `TEKTAT/` set);
- **per Kenya export market** - one HS product line, several importing markets
  (e.g. the cut-flowers set in `Samplefiles/Competitor Analysis/CA files/`,
  where every file ends `_0603.xlsx`). Only markets where Kenya exported the
  product are kept.

```bash
python3 generate_competitor_analysis.py --data TEKTAT --out TEKTAT
python3 generate_competitor_analysis.py \
    --data "Samplefiles/Competitor Analysis/CA files" --out output/cut-flowers
```

In the per-market layout, the countries shown are the destinations Kenya
supplies, ranked by Kenya's exports to them; each block lists the market's top
5 exporters (Kenya's competitors) with value and share. Kenya is always
included in red. In the per-product layout, products are ranked by import
value. Use `--year`, `--top` or `--mode auto|product|country` to override the
defaults.

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
  `webapp/jobs/`, fetched source files under `webapp/fetches/`, and both are
  cleaned after their TTLs. The SQLite registry and generated secret also sit
  on Render's ephemeral filesystem, so job history is lost on redeploy.
- Set `TRADEFLOW_SECRET` in Render to keep signed share links valid across
  restarts and redeploys; otherwise the generated secret is recreated.
- The whole app -- including the original single-country report generator -- is
  served from one process via `uvicorn webapp.main:app`
  (see `render.yaml`); the launch code works whether the app is imported as a
  script, a module, or by gunicorn/uvicorn `webapp.main:app`.
- The app has no user accounts. Each browser receives an owner token for the
  jobs it creates, and share links are signed, read-only and expiring. Set
  `TRADEFLOW_API_KEY` for a private deployment, or
  `TRADEFLOW_ENFORCE_JOB_OWNERSHIP=0` only when bare job URLs are acceptable.
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
- **Africa & Kenya destination insight** (Section 2.3): the share of the partner
  market's exports destined to Africa (aggregated over the African destination
  rows of Table 3), and the share destined to Kenya specifically — surfaced as a
  narrative bullet and a compact "Focus" table (Table 3a) with Kenya highlighted.
  Powered by the African-country mapping in `country_names.is_africa`.

## Batch / desktop-research deliverables (`batch_briefs.py`)

A general, market-agnostic runner covering the RID desktop-research deliverables
of the FY 2026/27 workplan. It is **data-driven**: every deliverable is a
reusable template applied to any market in the manifest, so new markets need no
code changes.

    python3 batch_briefs.py --manifest briefs_manifest_template.json --out-dir output/briefs

Manifest-driven types:

- `trade_brief` — full Kenya–<Market> trade-flow report (Tasks 2/3)
- `regional_brief` — a brief framed for a regional grouping (EAC / COMESA / AfCFTA)
- `secondary_research` — compact market fact-sheet for time-boxed desk research (e.g. Egypt/Nigeria) (Task 1)
- `performance_brief` — quarterly/annual export-performance summary feat. YoY + CAGR (Tasks 4/5)
- `market_intelligence` — value-chain analysis for a prioritized value chain (Task 7)
- `tariff_ntb` — tariff / non-tariff barrier documentation scaffold (Tasks 8/14)
- `prefeasibility` — market-gravity ranking to justify warehouse siting (Task 10)

Every automatic ITC Trade Map pull must request **4-digit (HS4)** product detail,
not the default coarser level ("`hs_level`" in the manifest, default 4).

The original report generator (`make_tables.find_source_files`) picks the right
raw ITC workbook for each table by filename keyword, and when several versions
of the same product table are uploaded it **prefers the HS4 copy automatically**
(by probing each file's product-code column). `generate_report` then warns on
`stderr` if a product table's codes are not HS4 (e.g. an HS2 or HS6 download),
so you can re-download at 4-digit level. The bundled sample tables are already HS4.

## Notes on fidelity to the reference report:

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
