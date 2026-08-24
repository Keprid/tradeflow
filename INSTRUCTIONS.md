# INSTRUCTIONS - Trade Flow Report Generator

This guide walks you through installing and running the Trade Flow Report
Generator on a Windows PC or a Linux machine.

The pipeline has two steps:

1. **`make_tables.py`** – builds the six "Table 1 .. Table 6" workbooks from the
   raw ITC (International Trade Centre) downloads. The workbooks are live Excel
   files: the Share column, the "All other …" rows, the "World"/"All products"
   totals and the Figure 1 "Balance of Trade" row are Excel **formulas** (with
   cached results), so editing a number recalculates the sheet. Header rows are
   frozen, each table has **AutoFilter** dropdowns, the workbooks **recalc on
   open**, and Figure 1 embeds a **line chart** of exports/imports/balance. A
   combined **`All Tables.xlsx`** with all seven sheets is written alongside
   the individual files.
2. **`generate_report.py`** – reads those tables (plus the trade-balance file)
   and writes the Word report.

Layout and wording follow the reviewed reference
`KENYA-SAUDI ARABIA TRADE FLOW_Reviewed 1.docx` (e.g. "Compiled by KEPROBA",
the imports bullets in section 3.1, "Source: Google map", split References).
The report year range adapts to the data: any run of **2 to 10 consecutive
years** (2021-2025 for Saudi Arabia, 2021-2024 for the UAE, or the short
2023-2025 windows produced by the previous Trade Map interface) is accepted.

If you already have the seven ready-made Excel files you can skip straight to
step 2. For full details on building the tables, see `MAKE_TABLES.md`.

---

## 1. What you need

- Python 3.9 or newer. Get it at https://www.python.org/downloads/
  - On **Windows**, during installation tick **"Add Python to PATH"**.
  - Check it works by opening Command Prompt (or Terminal) and typing:
    ```
    python --version
    ```

---

## 2. Install the required libraries (one time)

Open Command Prompt / Terminal in the folder containing this project
(the folder with `generate_report.py`), then run:

```
pip install -r requirements.txt
```

This installs `python-docx`, `openpyxl`, `matplotlib`, plus the packages used
by the web app (`fastapi`, `uvicorn`, `python-multipart`).

> Windows only: if `pip` is not recognised, try `py -m pip install -r requirements.txt`.

---

## 3. Web app (easiest option)

Instead of the command-line steps below, you can generate a report from your
browser:

1. Launch the app with the one-word command **`tradeflow`**:
   - **Windows:** double-click **`tradeflow.bat`** (or add the Trade Flow
     folder to `PATH` and type `tradeflow` in any Command Prompt). It runs
     `install_deps.bat`, which finds Python and auto-installs any missing
     packages, then opens your browser — no terminal needed.
   - **Linux / macOS:** run
     ```
     chmod +x tradeflow
     ./tradeflow
     ```
     (put it on your PATH, e.g. `ln -s "$PWD/tradeflow" ~/.local/bin/`,
     to run `tradeflow` from anywhere).
   - The older launcher **`run_webapp.bat`** still works the same way.
2. Open http://127.0.0.1:8000 in your browser (the launcher usually does this
   for you).
3. Pick the country config, choose your files and click **Generate**.

The upload can be the six raw ITC downloads, the seven ready-made tables, or
a zip of either set. When raw files are uploaded the app builds the tables for
you automatically. When it finishes you can download the `.docx` report and a
zip of the generated tables. Generated jobs are stored under `webapp/jobs/`
and cleaned up after 24 hours.

Press `Ctrl+C` in the terminal to stop the web app.

---

## 4. Build the tables from raw ITC files (optional step 1)

If you only have the six **raw ITC downloads** (not the ready-made Table 1-6
files), put them in one folder (the `sourcefiles` folder is pre-loaded with
the Saudi Arabia files). Both Trade Map interfaces are accepted — the beta
slugs below and the previous-interface names (`Trade_Map_-_List_of_
supplying_markets_for_a_product_imported_by_<Partner>.xls`, `..._List_of_
products_imported_by_...`, `..._List_of_importing_markets_for_a_product_
exported_by_...`, `..._List_of_products_exported_by_...` and the two
`Trade_Map_-_Bilateral_trade_between_Kenya_and_<Partner>` workbooks, which
are told apart by their content):

| File name contains...        | Table built             | Units      |
|------------------------------|-------------------------|------------|
| `imports-from-world-by-exporter` | Table 1 import source markets | USD billion |
| `imports-from-world-by-product`  | Table 2 import products       | USD billion |
| `exports-to-world-by-importer`   | Table 3 export destinations   | USD billion |
| `exports-to-world-by-product`    | Table 4 export products       | USD billion |
| `kenyas-exports-to-<country>-by-product` | Table 5 Kenya exports | USD million |
| `kenyas-imports-from-<country>-by-product` | Table 6 Kenya imports | USD million |

### On Windows (easiest)

Double-click **`run_tables.bat`**. It first runs `install_deps.bat` (finds
Python and auto-installs any missing packages), then builds all six tables
from `sourcefiles\` into `output\tables\` and also writes the
`Figure 1 Trade Balance.xlsx` (derived from the Table 5 / Table 6 totals).

### Manual command (Windows or Linux)

```
python make_tables.py --excel-dir sourcefiles --out-dir output/tables
```

The balance file is created automatically; nothing else needs to be copied.

> The script ranks rows by the latest year (2025 for Saudi, 2024 for UAE),
> keeps the top 25 exporters (Tables 1 and 3) and top 20 products (Tables
> 2, 4, 5, 6), inserts Kenya's own row if it is outside the kept rows
> (highlighted), and adds the "All other countries/products" and
> World/All-products totals.

---

## 5. Prepare your data files (or use the built tables)

Put the **seven ITC Excel files** in one folder (the `sample_data` folder is
pre-loaded with the Saudi Arabia files):

| File name contains... | Contents                          | Units       |
|-----------------------|-----------------------------------|-------------|
| `Table 1`             | Country import source markets     | USD billion |
| `Table 2`             | Country import products           | USD billion |
| `Table 3`             | Country export destination markets| USD billion |
| `Table 4`             | Country export products           | USD billion |
| `Table 5`             | Kenya exports to the country      | USD million |
| `Table 6`             | Kenya imports from the country    | USD million |
| `Figure 1` (or `Balance`) | Kenya-country trade balance   | USD million |

The script finds files by these keywords in the file name, so the exact
names don't matter as long as each keyword appears.

---

## 6. Run the generator (command line)

### On Windows (easiest)

Double-click **`run.bat`**. It first runs `install_deps.bat` (finds Python
and auto-installs any missing packages), then generates the report using the
Saudi Arabia configuration and the `sample_data` files, and saves it in the
`output` folder. A black window will stay open with the result - press any
key to close it.

### Manual command (Windows or Linux)

```
python generate_report.py --excel-dir sample_data --config config/saudi_arabia.json --output "output/KENYA-SAUDI ARABIA TRADE FLOW.docx"
```

- `--excel-dir` : folder with the 7 Excel files (if you built the tables with
  `make_tables.py`, point this at `output/tables` instead)
- `--config`   : the per-country JSON config (see below)
- `--output`   : where to save the Word document (optional)

On Linux use `python3` instead of `python` if needed.

---

## 6b. Kenya Quarterly Performance Report (KRA data)

This is a separate, simpler report: the quarterly *Kenya Export Performance*
report built from two raw KRA customs extracts (the files named
`…Exports by HS Destination…` and `…Imports by HS Origin…`).

**Web app:** choose **Quarterly Performance**, upload the KRA `.xls` files
(up to four: exports + imports for each of the two years, or a single file
per side covering both years; a zip also works), and click Generate.
Everything else is automatic. You can also upload the already-built
`Exports.xlsx` + `Imports.xlsx` instead.

**Command line:**

```
python make_quarterly_tables.py --excel-dir "EXPORT PERFORMANCE FOR Q2" --out-dir output/quarterly
python generate_quarterly_report.py --excel-dir output/quarterly --output output/report.docx
```

The first command builds `Exports.xlsx` / `Imports.xlsx` (values in Ksh.
Billion, live formulas). The second writes the Word report with title page,
table of contents, lists of tables and figures, the balance-of-trade chart,
Tables 1–6, Figures 1–2, a data-driven deduction section and both annexes.
Title page and contents are not numbered; body pages restart at 1.

**Year-on-year comparison:** every table (1–6) shows the previous year's
quarter next to the current one, plus a change column with green ▲ / red ▼
arrows. Tables 2 and 4 list the markets/products that grew or fell by over
Ksh. 1 billion; Figures 1–2 show "SHARE IN <Y>" above "SHARE IN <Y-1>",
stacked vertically on a full A4 page. Upload extracts covering **both
years** (e.g. Jan–Jun 2026 plus Apr–Jun 2025).

**Single year only?** Then the comparison is skipped automatically — no
synthetic data is created. Tables become `Rank | Item | months | Total |
Share`, the growth/decline exhibits disappear and numbering shifts to
Tables 1–4; each figure shows a single doughnut. For testing the comparison
layout anyway, add `--synthesize-prev-year` on the command line (fixed-seed
fake prior-year rows).

---

## 7. Generating a report for a different country

Two ways to get the country config:

**Option A - auto-generate it (recommended).** The config file can be created
automatically from the data, so you do not have to hand-write one when the
country (and its data) changes frequently:

1. Put that country's tables in a folder: either the 7 ready-made Excel files,
   or the 6 raw ITC downloads (the script builds the tables first).
2. Run:
   ```
   python make_config.py <excel-folder>
   ```
   This detects the country name from Table 1, reads the report years from the
   table columns, and writes `config/<country>.json` (for example
   `config/uganda.json`). In the web app, the country is always
   auto-detected from the uploaded data: it reuses an existing config if the
   country already has one, otherwise it creates the JSON on the spot.

**Option B - write it by hand.** Copy `config_template.json` to
`config/<country>.json`.

Then fill in the country-specific text that the program cannot compute:
   - `country.name`, `country.short`, `country.possessive`, `country.title`,
     `country.capital`
   - `report.title_line1/2` and `report.year`
   - `world_trade.exports/imports` (world totals in USD billion, used for the
     share-of-world sentences)
   - `quick_facts` (Annex II table) - fill every `...` placeholder
   - `references`
   - `export_potential.image` and `map.image` must point to real image files
     (paths are resolved relative to the config file).
   - The Section 1 "Backgrounds" text is **auto-generated** from the data;
     do not add a `background` key to the config.
3. Get that country's tables (skip if you already did this for Option A):
   Either put the 7 ready-made Excel files in a folder (or replace the files
   in `sample_data`), or drop the 6 raw ITC downloads in a folder and build
   them first:
   ```
   python make_tables.py --excel-dir <raw-folder> --out-dir <tables-folder>
   ```
   (the `Figure 1 Trade Balance.xlsx` for that country is created
   automatically from the Table 5 / Table 6 totals).
4. Run:
   ```
   python generate_report.py --excel-dir <excel-folder> --config config/<country>.json --output "output/KENYA-<COUNTRY> TRADE FLOW.docx"
   ```

---

## 8. After generation (important)

- Open the generated `.docx` in Microsoft Word.
- The Table of Contents shows a placeholder until you refresh it:
  right-click the table of contents and choose **"Update Field"**, then
  **"Update entire table"**.
- Any part of the Section 1 background that the program cannot compute from
  the Excel data is printed as **yellow-highlighted** `[RESEARCH NEEDED: ...]`
  text (for example GDP, population, IMF forecasts, WTO membership). Look for
  the yellow highlights, research each item, and replace the highlighted text
  with the researched facts. The same applies to the `...` placeholders in the
  Annex II quick-facts table.
- Review the narrative sentences - the script computes every figure from the
  Excel data automatically, but you can edit the text in Word as usual.

---

## 9. Troubleshooting

| Problem | Solution |
|---------|----------|
| `'python' is not recognized` | Reinstall Python and tick "Add Python to PATH", or use `py` instead of `python`. |
| `ModuleNotFoundError: No module named 'docx'` | Run `pip install -r requirements.txt`. |
| `[ERROR] Could not find all six raw source files` | Your `--excel-dir` folder is missing one of the six raw ITC files, or a file name lacks the required keyword (`imports-from-world-by-exporter`, `imports-from-world-by-product`, `exports-to-world-by-importer`, `exports-to-world-by-product`, `kenyas-exports-to-...-by-product`, `kenyas-imports-from-...-by-product`). |
| `[ERROR] Missing Excel file(s)` | Your data folder is missing one of the 7 files, or a file name lacks the required keyword (`Table 1`..`Table 6`, `Figure 1`/`Balance`). |
| Charts look wrong / font warnings | Not a problem - matplotlib falls back automatically. On Windows it uses Times New Roman. |
| Report looks different from template | Ensure fonts (Times New Roman) are installed on the machine. |

---

## 10. Files in this project

```
make_tables.py             Builds the six Table 1-6 workbooks from raw ITC files
make_config.py             Auto-generates config/<country>.json from the data
generate_report.py         Reads the tables and writes the Word report
webapp/                    FastAPI web app (main.py + static/index.html)
config/saudi_arabia.json   Working configuration (Saudi Arabia)
config/uae.json            Configuration for the United Arab Emirates
config/uganda.json         Configuration for Uganda (or let make_config.py create these)
config_template.json       Copy this to make a new country by hand
sourcefiles/               The 6 raw ITC downloads (for make_tables.py)
sample_data/               The 7 ready-made Saudi Arabia Excel files (for testing)
assets/                    Map + export-potential images used by the report
output/                    Generated tables (.docx reports go here too)
requirements.txt           Python packages to install
install_deps.bat           Finds Python and auto-installs missing packages
tradeflow                  One-word launcher (Linux/macOS): starts the web app
tradeflow.bat              One-word launcher (Windows): starts the web app
run_tables.bat             Windows launcher for make_tables.py
run.bat                    Windows one-click launcher for generate_report.py
run_webapp.bat             Windows launcher for the web app
README.md                  Technical overview
```
