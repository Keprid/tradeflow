# Pending Improvements - Kenya Services Trade Analysis

## Completed (This Session)

### Excel Formula Integration
- [x] **Live Excel formulas in all table writers** - Tables 1-4, Figure 1
  - Share columns use `=C{row}/C{total_row}` formulas
  - "All other" rows use `=C{world}-SUM(C{first}:C{last})` formulas
  - Balance of trade uses `=C{exports}-C{imports}` formulas
  - Added `_put_formula()` helper matching `make_tables.py` pattern
  - Added `_save_workbook()` with inject+verify+fallback pattern
  - All formulas have cached values for `data_only=True` reads

### Error Handling Improvements
- [x] **Missing file identification by name** - All pipelines
  - `find_service_files()`: Lists each missing file with expected keyword
  - `find_source_files()`: Lists each missing file with expected pattern
  - `find_service_excel_files()`: Lists expected keywords per missing file
  - `find_excel_files()`: Lists expected keywords per missing file
  - `_detect_mode()`: Shows exact uploaded filenames and expected keywords

### Technical Improvements
- [x] **CAGR calculation** - Compound Annual Growth Rate
  - `calc_cagr(vals, years)` function added
  - Computed for all items in `calc_growth_rates()`
  - Displayed in Tables 5-8 (dev status/region) and Table 13 (peers)

- [x] **Growth stability coefficient** - Coefficient of variation
  - `calc_growth_stability(vals)` function added
  - Measures consistency of year-over-year growth rates

### Competitive Position Analysis
- [x] **Export Concentration Index (HHI)** - Table 10 + Figure 7
  - Computes Herfindahl-Hirschman Index for Kenya vs World
  - Includes Top-3 and Top-5 concentration ratios
  - Effective number of categories metric

- [x] **Regional Classification Update** - Tables 7-8
  - Updated ITC_REGIONS to match tradebriefs.intracen.org
  - New regions: Africa, Asia, Americas, Pacific, Europe

### Market Opportunity Analysis
- [x] **Diversification Potential** - Table 11 + Figure 8
  - Opportunity score = global_growth × (1 - kenya_share) × import_penetration
  - Identifies fast-growing categories where Kenya is underrepresented

### Structural Transformation
- [x] **Value Composition Trajectory** - Table 12 + Figure 9
  - Groups: High-Value vs Traditional vs Other
  - Shows share evolution over time (2020-2024)

### Policy Benchmarking
- [x] **Kenya vs Peer Comparison** - Table 13
  - African Peers: South Africa, Egypt, Mauritius, Rwanda
  - Aspirational Peers: Singapore, Malaysia

---

## Pending Improvements (Requires New Data Downloads)

### Data Gaps Requiring ITC Trade Map Downloads

#### 1. Bilateral Partner Analysis
**Status:** Cannot implement without partner-country data

**Files needed:**
```
Trade_Map_-_List_of_partners_for_the_selected_service_(All_services)_exported_by_Kenya_*.xls
Trade_Map_-_List_of_partners_for_the_selected_service_(All_services)_imported_by_Kenya_*.xls
```

#### 2. Regional Integration Analysis
**Status:** Cannot implement without bilateral data

#### 3. Peer Country Composition Comparison
**Status:** Cannot implement without per-category data for peer countries

#### 4. Services Exports as Share of GDP
**Status:** Cannot implement without GDP time series data

#### 5. Barriers Analysis
**Status:** Cannot implement without regulatory data

### Remaining Technical Improvements

#### 6. Enhanced Diversification Metrics
- **HHI by development status** - Compare concentration for Developed vs Developing vs LDC groups
- **Export sophistication index** - Weighted average of partner income levels

#### 7. Enhanced Peer Comparison
- **Per-category peer comparison** - When peer category data is available
- **Market share convergence analysis** - Is Kenya catching up to peers?

#### 8. Enhanced Regional Analysis
- **Intra-regional trade intensity** - Kenya's trade intensity index within EAC/COMESA
- **Regional hub potential score** - Composite score of Kenya's regional position

#### 9. Report Generation Improvements
- **Dynamic figure numbering** - Currently hardcoded (Figure 1-11)
- **Narrative auto-generation** - Generate analytical text from table data
- **Interactive charts** - Add Plotly/Altair for web-based reports

---

## Implementation Priority

### High Priority (Next Session)
1. Bilateral partner data download and analysis
2. Regional integration (EAC/COMESA) analysis
3. Peer country composition comparison

### Medium Priority
4. Services/GDP ratio analysis
5. Barriers analysis
6. Enhanced diversification metrics

### Low Priority
7. Enhanced peer comparison
8. Enhanced regional analysis
9. Report generation improvements

---

## Notes

- All values from ITC Trade Map are in USD Thousand
- Tables 1-6, Global services: values ÷ 1e6 = USD Billion
- Tables 3-4, Kenya services: values ÷ 1e3 = USD Million
- RCA auto-detects latest year with actual Kenya category data (2024 is None for categories)
- The `parse_kenya_commercialized()` function is defined but never called - balance data comes from export-import difference
