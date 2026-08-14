# Instructions - Building the Table 1-6 Workbooks

`make_tables.py` turns the six raw ITC (International Trade Centre) downloads
into the six "Table 1 .. Table 6" workbooks (plus a `Figure 1 Trade Balance.xlsx`
derived from Tables 5 and 6) that `generate_report.py` expects.

---

## 1. What you need

- Python 3.9 or newer. Get it at https://www.python.org/downloads/
  - On **Windows**, during installation tick **"Add Python to PATH"**.
- The libraries in `requirements.txt` (one time, from the project folder):

  ```
  pip install -r requirements.txt
  ```

  > Windows only: if `pip` is not recognised, try
  > `py -m pip install -r requirements.txt`.

---

## 2. Gather the six raw ITC files

Put the six files in one folder (the `sourcefiles` folder is pre-loaded with
the Saudi Arabia files). The script finds them by keywords in the file name:

| File name contains... | Table built | Units |
|---|---|---|
| `imports-from-world-by-exporter` | Table 1 import source markets | USD billion |
| `imports-from-world-by-product`  | Table 2 import products       | USD billion |
| `exports-to-world-by-importer`   | Table 3 export destinations   | USD billion |
| `exports-to-world-by-product`    | Table 4 export products       | USD billion |
| `kenyas-exports-to-<country>-by-product` | Table 5 Kenya exports | USD million |
| `kenyas-imports-from-<country>-by-product` | Table 6 Kenya imports | USD million |

Example filenames for Saudi Arabia:

```
saudi-arabias-imports-from-world-by-exporter_all.xlsx
saudi-arabias-imports-from-world-by-product_all.xlsx
saudi-arabias-exports-to-world-by-importer_all.xlsx
saudi-arabias-exports-to-world-by-product_all.xlsx
kenyas-exports-to-saudi-arabia-by-product_all.xlsx
kenyas-imports-from-saudi-arabia-by-product_all.xlsx
```

---

## 3. Build the tables

### On Windows (easiest)

Double-click **`run_tables.bat`**. It builds all six tables (plus the
`Figure 1 Trade Balance.xlsx`) from `sourcefiles\` into `output\tables\` and
prints the names of the files it created.

### Manual command (Windows or Linux)

```
python make_tables.py --excel-dir sourcefiles --out-dir output/tables
```

- `--excel-dir` : folder with the 6 raw files (default `sourcefiles`)
- `--out-dir`   : where the Table 1-6 workbooks are written (default
  `output/tables`)
- `--top`       : how many rows to keep per table (default `20`)

On Linux use `python3` instead of `python` if needed.

---

## 4. What the script does

For every table it:

1. reads the values (stored in USD Thousand by ITC) and converts them:
   - Tables 1-4 divide by 1,000,000 -> "Value in USD Billion",
   - Tables 5-6 divide by 1,000 -> "Value in USD Million";
2. keeps the columns for the report years found in the files (2021-2025 for
   Saudi, 2021-2024 for the UAE files);
3. ranks the rows by the latest year, highest first, and writes the rank into
   a new first column;
4. keeps the top `--top` product rows (default 20) for Tables 2, 4, 5 and 6;
   Tables 1 and 3 always keep the **top 25** exporters/partners;
5. for Tables 1 and 3, if **Kenya** is not inside the kept rows it is still
   included, at the end, with its actual rank and a highlighted background;
6. sums the rows that are not shown into an
   "All other countries" / "All other products" row;
7. adds the **World** / **All products** total row;
8. adds a final "Share in <latest year> %" column (row value / total).

Product codes are written as text with a leading apostrophe (e.g. `'8703`) so
Excel does not reformat them. Long product labels are trimmed to read as
clean sentences: the whole `;`-separated clauses are kept while they fit
within ~80 characters, and any trailing dots are removed (no " ..." suffix).

Finally it writes **`Figure 1 Trade Balance.xlsx`** with the years from the
source files, where:

- **Exports** = Kenya's exports to the partner country (Table 5 total),
- **Imports** = Kenya's imports from the partner country (Table 6 total),
- **Balance of Trade** = Exports minus Imports.

All three are in USD Million. No separate balance file is needed.

### Live formulas and usability extras

Every derived cell is an **Excel formula with a cached result** (stored
exactly as Excel stores a recalculated formula), so:

- the report generator reads the correct numbers (it never needs Excel), and
- if you edit a value in Excel, the sheet recalculates live
  (workbooks also recalc on open via `fullCalcOnLoad`).

Header rows are **frozen** and each table has Excel **AutoFilter** column
dropdowns. Figure 1 embeds a **line chart** of exports/imports/balance by
year below the data. A combined **`All Tables.xlsx`** workbook with all seven
sheets in one file is written alongside the individual ones.

---

## 5. After building

The output folder contains:

```
All Tables.xlsx
Figure 1 Trade Balance.xlsx
Table 1 <Country> Import Source Markets.xlsx
Table 2 <Country> Lead Imports.xlsx
Table 3 <Country> Lead Export Destinations.xlsx
Table 4 <Country> Lead Export Products.xlsx
Table 5 Kenya Exports to <Country>.xlsx
Table 6 Kenya Top Imports from <Country>.xlsx
```

---

## 6. Next step

Feed the output folder to the report generator:

```
python generate_report.py --excel-dir output/tables --config config/saudi_arabia.json --output "output/KENYA-SAUDI ARABIA TRADE FLOW.docx"
```

---

## 7. Troubleshooting

| Problem | Solution |
|---|---|
| `[ERROR] Could not find all six raw source files` | The `--excel-dir` folder is missing one of the six files, or a file name lacks the required keyword (see the table in section 2). |
| `'python' is not recognized` | Reinstall Python and tick "Add Python to PATH", or use `py` instead of `python`. |
| `ModuleNotFoundError: No module named 'openpyxl'` | Run `pip install -r requirements.txt`. |
| Wrong years or columns in the output | Make sure the raw files contain the year columns (e.g. 2021-2025 for Saudi, 2021-2024 for the UAE files); the script keeps the last year columns found. |
