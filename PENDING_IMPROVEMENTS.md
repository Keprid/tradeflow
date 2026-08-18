# Pending Improvements - Kenya Services Trade Analysis

## Completed (This Session)

### Competitive Position Analysis
- [x] **Export Concentration Index (HHI)** - Table 10 + Figure 7
  - Computes Herfindahl-Hirschman Index for Kenya vs World
  - Includes Top-3 and Top-5 concentration ratios
  - Effective number of categories metric

- [x] **Regional Classification Update** - Tables 7-8
  - Updated ITC_REGIONS to match tradebriefs.intracen.org
  - New regions: Africa, Asia, Americas, Pacific, Europe
  - Replaces old: Africa, Eastern Europe/Central Asia, Middle East/North Africa, Asia/Pacific, Latin America/Caribbean

### Market Opportunity Analysis
- [x] **Diversification Potential** - Table 11 + Figure 8
  - Opportunity score = global_growth × (1 - kenya_share) × import_penetration
  - Identifies fast-growing categories where Kenya is underrepresented
  - Considers import substitution potential

### Structural Transformation
- [x] **Value Composition Trajectory** - Table 12 + Figure 9
  - Groups: High-Value (Finance, ICT, IP, Other business) vs Traditional (Transport, Travel, Construction) vs Other
  - Shows share evolution over time (2020-2024)
  - Absolute values in USD Millions

### Policy Benchmarking
- [x] **Kenya vs Peer Comparison** - Table 13
  - African Peers: South Africa, Egypt, Mauritius, Rwanda
  - Aspirational Peers: Singapore, Malaysia
  - Shows total exports, global share, growth, and ranking

---

## Pending Improvements (Requires New Data Downloads)

### Data Gaps Requiring ITC Trade Map Downloads

#### 1. Bilateral Partner Analysis
**Status:** Cannot implement without partner-country data

**What's needed:**
- Download Kenya's services exports by partner country from ITC Trade Map
- Download Kenya's services imports by partner country from ITC Trade Map

**Files needed:**
```
Trade_Map_-_List_of_partners_for_the_selected_service_(All_services)_exported_by_Kenya_*.xls
Trade_Map_-_List_of_partners_for_the_selected_service_(All_services)_imported_by_Kenya_*.xls
```

**Analyses enabled:**
- Kenya's service export gap with key partners
- Services trade balance by partner
- Which partners does Kenya run a surplus/deficit with?
- Import substitution opportunities

#### 2. Regional Integration Analysis
**Status:** Cannot implement without bilateral data

**What's needed:**
- Same bilateral partner data as above
- EAC member state identification
- COMESA member state identification

**Analyses enabled:**
- EAC/COMESA services trade flows
- Kenya's position as regional services hub
- Which categories dominate intra-regional trade?
- Regional services trade balance

#### 3. Peer Country Composition Comparison
**Status:** Cannot implement without per-category data for peer countries

**What's needed:**
- Download service exports by category for: Singapore, Malaysia, South Africa, Egypt, Mauritius, Rwanda
- Files needed per country:
```
Trade_Map_-_List_of_services_exported_by_{Country}_(All_services).xls
```

**Analyses enabled:**
- Kenya vs Singapore: services export composition comparison
- Kenya vs Malaysia: services export composition comparison
- Peer benchmarking on high-value vs traditional services share

#### 4. Services Exports as Share of GDP
**Status:** Cannot implement without GDP time series data

**What's needed:**
- World Bank WDI data for Kenya GDP (2020-2024)
- Or IMF World Economic Outlook data
- Download from: https://data.worldbank.org/indicator/NY.GDP.MKTP.CD

**Analyses enabled:**
- Kenya's services exports as % of GDP over time
- Comparison with global trend
- Services sector structural transformation metrics

#### 5. Barriers Analysis
**Status:** Cannot implement without regulatory data

**What's needed:**
- WTO Trade Policy Review data
- OECD Services Trade Restrictiveness Index (STRI)
- Or UNCTAD services trade barriers database

**Analyses enabled:**
- Which service categories face most regulatory barriers
- Mode 1-4 restrictions analysis
- Policy recommendations for liberalization

### Technical Improvements

#### 6. Enhanced Diversification Metrics
- **Herfindahl-Hirschman Index (HHI) by development status** - Compare concentration for Developed vs Developing vs LDC groups
- **Export sophistication index** - Weighted average of partner income levels
- **Growth stability coefficient** - Standard deviation of growth rates across categories

#### 7. Enhanced Peer Comparison
- **Per-category peer comparison** - When peer category data is available
- **Growth trajectory comparison** - How peer growth rates compare to Kenya's
- **Market share convergence analysis** - Is Kenya catching up to peers?

#### 8. Enhanced Regional Analysis
- **Intra-regional trade intensity** - Kenya's trade intensity index within EAC/COMESA
- **Regional hub potential score** - Composite score of Kenya's regional position
- **Spillover effects analysis** - How regional growth affects Kenya's services

#### 9. Data Quality Improvements
- **Handle missing 2024 category data** - Currently auto-detects latest year with data
- **CAGR calculation** - Add compound annual growth rate to all tables
- **5-year trend analysis** - Instead of just YoY growth

#### 10. Report Generation Improvements
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
9. Data quality improvements
10. Report generation improvements

---

## Notes

- All values from ITC Trade Map are in USD Thousand
- Tables 1-6, Global services: values ÷ 1e6 = USD Billion
- Tables 3-4, Kenya services: values ÷ 1e3 = USD Million
- RCA auto-detects latest year with actual Kenya category data (2024 is None for categories)
- The `parse_kenya_commercialized()` function is defined but never called - balance data comes from export-import difference
