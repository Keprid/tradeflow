# Pakistan National Accounts Dashboard — Feature Audit & Replication Blueprint

Source website: https://na.data.gov.pk/

## 1. Website Architecture Overview

The website can be understood as a collection of interconnected analytical dashboards:

```text
National Accounts Data Platform
│
├── GDP Dashboard
│   ├── Annual GDP
│   ├── Sector Analysis
│   ├── GDP Growth
│   ├── Economic Indicators
│   └── Heat Tables
│
├── Quarterly National Accounts
│   ├── Quarterly GDP
│   ├── Sector Shares
│   ├── Growth Trends
│   └── Quarterly Heat Tables
│
├── Agriculture Statistics
│   ├── Crop Explorer
│   ├── Geographic Filters
│   ├── Production Statistics
│   ├── Yield Analysis
│   ├── Maps
│   └── Bar Chart Race
│
├── Tax Collection
│   ├── Annual Collection
│   ├── Growth Trends
│   └── Tax Heat Tables
│
└── Global GDP Comparison
```

---

## 2. Navigation Structure

### Recommended navigation model

```text
LOGO
│
├── Home / GDP
├── Quarterly GDP
├── Agriculture
├── Tax
├── Global Comparison
└── About / Data Source
```

### Current functional modules

| Module | Purpose |
|---|---|
| GDP | Annual national accounts |
| QNA | Quarterly national accounts |
| Agriculture | Crop statistics |
| Tax | Tax collection analytics |
| Global | International GDP comparison |

---

## 3. Annual GDP Dashboard

This is the platform's core analytical page.

### 3.1 Sector Navigation Tabs

```text
[ GDP ]
[ Agriculture ]
[ Industry ]
[ Service ]
[ Subsidies ]
[ Taxes ]
```

These are data dimension filters rather than ordinary navigation buttons.

### Functional behavior

When a user clicks a sector such as Agriculture, the dashboard should dynamically update:

- Main GDP chart
- Pie chart
- Historical trend
- Growth charts
- Heatmap/table

This creates a linked visualization system.

---

## 4. GDP Combination Chart

### Chart type

A dual-axis combination chart.

### Data dimensions

```text
X-axis
└── Financial Year

Primary Y-axis
└── GDP Value

Secondary Y-axis
└── GDP Growth Rate (%)
```

### Visualization structure

```text
GDP VALUE
   │
   │       █
   │   █   █
   │ █ █   █
   └──────────────── Year

                    ╱╲
Growth Rate ───────╱  ╲────
```

### Important interaction

The chart can act as a dashboard controller.

Conceptually:

```javascript
onGDPYearClick(year) {
    updateSectorPieChart(year)
    updateSectorAreaChart(year)
    updateIndicators(year)
}
```

---

## 5. Sectoral GDP Share Pie Chart

### Purpose

Shows how GDP is distributed across major sectors:

```text
        GDP

   Agriculture
       23%

Industry          Services
  19%               58%
```

### Interaction

Clicking a sector should update related charts.

### Recommended logic

```javascript
onSectorClick(sector) {
    selectedSector = sector
    updateHistoricalSectorChart(sector)
    updateGrowthChart(sector)
    updateHeatTable(sector)
}
```

This represents a cross-filtering dashboard architecture.

---

## 6. Historical Sector Share Chart

Tracks the percentage contribution of major sectors over time.

```text
60% ───────────────── Services
          ╲
50%        ╲
            ╲

30% ─ Agriculture
      ╲
       ╲

20% ─────── Industry

     2000 → 2026
```

### Features

- Multi-series line chart
- Historical time series
- Sector comparison
- Percentage contribution
- Data-table accessibility

---

## 7. Economic Indicator Dashboard

The GDP page integrates macroeconomic indicators.

### Exchange Rate

```text
Year → Exchange Rate
        Percentage Change
```

### NPI

Tracks:

- Historical values
- Percentage changes
- Time-series trends

### Per Capita Income

Tracks:

- Income level
- Historical trend
- Percentage change

Each indicator can be represented using interactive charts and tabular views.

---

## 8. Growth Analysis Module

### Section

```text
Historical Trend of Sectoral / GDP Growth
```

Possible controls:

```text
[ Growth Bands ]

[ Govt Bands ]
```

### Core chart

```text
Sector Growth Over the Years

          GDP
Growth ─── Agriculture
          Industry
          Services

          ↓

      Historical Years
```

---

## 9. GDP Heat Table

One of the strongest analytical features is a matrix view of sectors across time.

Example:

| Sector | 2021 | 2022 | 2023 | 2024 | 2025 |
|---|---:|---:|---:|---:|---:|
| Agriculture | 23% | 22% | 23% | 24% | 23% |
| Industry | 19% | 18% | 18% | 18% | 19% |
| Services | 58% | 59% | 58% | 58% | 58% |

The detailed structure can extend into subsectors.

### Possible subsectors

#### Agriculture

- Crops
- Fishing
- Forestry
- Livestock

#### Industry

- Manufacturing
- Construction
- Mining
- Electricity

#### Services

- Education
- Finance & Insurance
- Government
- Hotels & Restaurants
- Health
- Information & Communication
- Other Private Services

### Required features

#### Filter controls

```text
[ Select Metric ▼ ]

[ Select Category ▼ ]
```

#### Table interactions

- Horizontal scrolling
- Sticky sector column
- Sticky header
- Conditional colors
- Tooltip on cells
- Sorting
- Export CSV
- Search

### Improved heatmap concept

```text
Low Value     🟦
Medium Value  🟨
High Value    🟥
```

---

## 10. Quarterly National Accounts (QNA)

A second GDP dashboard operating at quarterly frequency.

### Core features

#### Quarterly GDP chart

```text
2016Q1
2016Q2
2016Q3
...
2025Q4
```

#### Metrics

- GDP value
- Year-over-year growth

#### Quarterly sector share

```text
GDP Share FY2025 Q4

Agriculture
Industry
Services
```

#### Historical quarterly sector share

Allows analysis of sector composition across quarters.

#### Quarterly heat table

| Sector | 2016Q1 | 2016Q2 | ... | 2025Q4 |
|---|---:|---:|---|---:|
| Agriculture | | | | |
| Crops | | | | |
| Livestock | | | | |
| Industry | | | | |
| Manufacturing | | | | |
| Services | | | | |

---

## 11. Agriculture Statistics Module

The agriculture module is one of the most sophisticated data exploration components.

Potential coverage:

- Crop production
- Yield
- Cultivation area
- Geographic disaggregation
- Historical trends
- District-level analysis

---

## 12. Cascading Geographic Filters

A key UI pattern is hierarchical filtering.

```text
Province
   ↓
Division
   ↓
District
   ↓
Year
   ↓
Crop
```

### Cascading dropdown logic

```javascript
selectProvince(province)

↓

loadDivisions(province)

↓

selectDivision(division)

↓

loadDistricts(division)

↓

selectDistrict(district)

↓

updateCropStatistics()
```

API-driven loading is preferable to loading all geographic data at once.

---

## 13. Crop Selection System

Users can select:

```text
Province
Division
District
Year
Crop
```

Example:

```text
Province: Punjab
Division: Lahore
District: Kasur
Year: 2024
Crop: Wheat
```

The dashboard can then display:

- Crop name
- Cultivation area
- Production
- Yield

---

## 14. Agricultural KPI Cards

Example KPI card:

```text
┌──────────────────────┐
│ CULTIVATION AREA     │
│                      │
│  1,240,000 Acres     │
│                      │
│  ↑ 4.2%              │
└──────────────────────┘
```

Recommended cards:

### Card 1

```text
Crop
Wheat
```

### Card 2

```text
Cultivation Area
1.2 Million Acres
```

### Card 3

```text
Production
3.8 Million Tons
```

### Card 4

```text
Yield
3.1 Tons / Acre
```

---

## 15. Yearly Crop Trend Analysis

Recommended charts:

### Cultivation Area

```text
Area
 │       ╭──
 │    ╭──╯
 │ ╭──╯
 └────────────────
  2018  2020  2022  2024
```

### Production

```text
Production
 │          ╭───
 │      ╭───╯
 │  ╭───╯
 └──────────────────
```

### Yield

```text
Yield per Acre
 │
 │  ╭──────
 │──╯
 └──────────────
```

---

## 16. Best District Feature

Example:

```text
🏆 Best District

District: Faisalabad

Highest Production:
1,500,000 Tons
```

Potential ranking metrics:

- Largest cultivation area
- Highest production
- Highest yield

---

## 17. Best Year Feature

Example:

```text
🏆 Best Year

Year: 2022

Production:
2,400,000 Tons
```

Possible metrics:

```text
Best Year by Area
Best Year by Production
Best Year by Yield
```

---

## 18. Yearly Crop Data Table

| Fiscal Year | Area | Production | Yield |
|---|---:|---:|---:|
| 2020 | 200 | 500 | 2.5 |
| 2021 | 220 | 540 | 2.45 |
| 2022 | 250 | 650 | 2.60 |

### Recommended improvements

- Search
- Sorting
- Pagination
- CSV download
- Excel download
- Copy table
- Print
- Column visibility controls

---

## 19. Geographic Visualization

A modern implementation should support:

```text
Pakistan Map
│
├── Province
│
├── Division
│
└── District
```

### Map interactions

Hover:

```text
District: XYZ
Production: 240,000 Tons
Yield: 3.4
```

Click:

```text
Open district analytics
```

### Recommended technologies

- Leaflet
- Mapbox
- GeoJSON
- React Leaflet

---

## 20. Bar Chart Race

Example:

```text
2020

Punjab       ███████████
Sindh        ████████
KPK          █████
Balochistan  ███
```

Then:

```text
2021

Punjab       █████████████
Sindh        █████████
...
```

Useful dimensions:

- Crop production by district
- Yield by province
- Cultivation area by district
- Top crops over time

---

## 21. Tax Collection Dashboard

Three major analytical components can be identified.

### A. Annual Tax Collection Overview

Time range controls:

```text
[10 Years]

[15 Years]

[20 Years]

[MAX]
```

Recommended behavior:

```javascript
setTimeRange(10)
setTimeRange(15)
setTimeRange(20)
setTimeRange("ALL")
```

### B. Tax Growth Analysis

Controls:

```text
[ Sector ]

[ SubSector ]

[ Dropdown ▼ ]
```

Potential hierarchy:

```text
Total Tax
   │
   ├── Income Tax
   │
   ├── Sales Tax
   │
   ├── Customs
   │
   └── Federal Excise
```

### C. Tax Heat Table

| Tax Category | 2020 | 2021 | 2022 | 2023 |
|---|---:|---:|---:|---:|
| Income Tax | | | | |
| Sales Tax | | | | |
| Customs | | | | |

---

## 22. Global GDP Comparison

A global comparison page can support:

```text
Country Selector

[ Pakistan ▼ ]
[ India ]
[ Bangladesh ]
[ China ]
[ USA ]
```

### GDP comparison

```text
GDP
│
│           █████ China
│
│      ███ India
│
│   ██ Pakistan
│
└────────────────
```

### Useful indicators

- Nominal GDP
- GDP growth
- GDP per capita
- PPP GDP
- Inflation
- Population

---

# 23. Core Interaction Architecture

The most important technical concept behind this website is linked dashboard state.

```text
User Interaction
       │
       ▼
Dashboard State
       │
 ┌─────┼──────┐
 ▼     ▼      ▼
Chart  Table  KPI
```

Example state:

```text
Selected Year = 2024
Selected Sector = Agriculture
Selected Metric = GDP Share
```

Every visualization reads from shared dashboard state.

---

# 24. Recommended React Architecture

```text
src/
│
├── components/
│   ├── Navbar.jsx
│   ├── Sidebar.jsx
│   ├── KPIcard.jsx
│   ├── FilterPanel.jsx
│   ├── ChartContainer.jsx
│   ├── HeatTable.jsx
│   └── DataTable.jsx
│
├── charts/
│   ├── GDPChart.jsx
│   ├── SectorPie.jsx
│   ├── GrowthChart.jsx
│   ├── ExchangeRateChart.jsx
│   ├── CropProductionChart.jsx
│   └── BarChartRace.jsx
│
├── pages/
│   ├── GDP.jsx
│   ├── QuarterlyGDP.jsx
│   ├── Agriculture.jsx
│   ├── Tax.jsx
│   └── Global.jsx
│
├── hooks/
│   └── useDashboardData.js
│
├── services/
│   └── api.js
│
└── App.jsx
```

---

# 25. Recommended Backend Architecture

```text
                    FRONTEND
                   React / Next.js
                         │
                         │ REST / JSON
                         ▼
                    BACKEND API
                 Django / FastAPI
                         │
            ┌────────────┼────────────┐
            ▼            ▼            ▼
         PostgreSQL    Redis       Celery
            │            │            │
            ▼            ▼            ▼
         Data        Caching      ETL Jobs
```

---

# 26. Suggested API Design

## GDP

```text
GET /api/gdp
```

Parameters:

```text
?year=2024
&sector=agriculture
```

## GDP historical series

```text
GET /api/gdp/trends
```

Example:

```json
{
  "years": [2020, 2021, 2022],
  "agriculture": [23.5, 23.0, 22.6],
  "industry": [18.8, 19.2, 19.0],
  "services": [57.7, 57.8, 58.4]
}
```

## Quarterly data

```text
GET /api/qna?q=2024Q4&sector=industry
```

## Agriculture

```text
GET /api/agriculture/crops
```

Parameters:

```text
?province=Punjab
&district=Kasur
&crop=Wheat
&year=2024
```

## Geographic hierarchy

```text
GET /api/provinces

GET /api/divisions?province=Punjab

GET /api/districts?division=Lahore
```

---

# 27. Recommended Database Design

## GDP Table

```text
gdp_data
```

| Column | Type |
|---|---|
| id | Integer |
| year | Integer |
| sector | String |
| subsector | String |
| value | Decimal |
| growth_rate | Decimal |

## Quarterly GDP

```text
quarterly_gdp
```

| Column | Type |
|---|---|
| id | Integer |
| year | Integer |
| quarter | Integer |
| sector | String |
| value | Decimal |
| yoy_growth | Decimal |

## Agriculture

```text
crop_statistics
```

| Column | Type |
|---|---|
| id | Integer |
| province_id | FK |
| division_id | FK |
| district_id | FK |
| crop_id | FK |
| year | Integer |
| area | Decimal |
| production | Decimal |
| yield | Decimal |

## Geography hierarchy

```text
provinces
    │
    └── divisions
           │
           └── districts
```

---

# 28. Recommended Visualization Stack

| Feature | Technology |
|---|---|
| Line charts | Apache ECharts |
| Combination charts | Apache ECharts |
| Pie charts | Apache ECharts |
| Heatmaps | Apache ECharts |
| Data tables | AG Grid / TanStack Table |
| Maps | Leaflet |
| Bar chart race | D3.js |
| Animations | Framer Motion |
| Dashboard layout | Tailwind CSS |

## Recommended stack

```text
React
+
TypeScript
+
Apache ECharts
+
TanStack Query
+
FastAPI
+
PostgreSQL
```

Apache ECharts is particularly suitable because it supports:

- Combination charts
- Dual Y-axis charts
- Large datasets
- Interactive tooltips
- Heatmaps
- Data zoom
- Chart linking

---

# 29. Features Missing That Could Improve the Platform

## Data export

```text
⬇ CSV
⬇ Excel
⬇ JSON
⬇ PDF
```

## API access

```text
Developers
│
├── API Documentation
├── API Key
├── Endpoints
└── Code Examples
```

## Data source metadata

Every chart should show:

```text
Source: Pakistan Bureau of Statistics

Dataset: National Accounts

Last Updated: August 2026

Frequency: Annual
```

## Chart download

```text
[ PNG ]
[ SVG ]
[ PDF ]
```

## Advanced filters

```text
Year Range
2000 ───────── 2026
```

## Comparison mode

```text
Compare:

Pakistan
vs
Bangladesh
vs
India
```

## Saved dashboards

Users could save:

```text
My Agriculture Dashboard

My GDP Comparison

My Tax Analysis
```

---

# 30. Best Overall System Architecture

```text
                         USERS
                           │
                           ▼
                    Next.js Frontend
                           │
                     API Gateway
                           │
                           ▼
                    FastAPI Backend
                           │
              ┌────────────┼────────────┐
              ▼            ▼            ▼
          PostgreSQL      Redis      Celery
              │            │            │
              │            │            ▼
              │            │        Data ETL
              ▼            ▼
          Data Warehouse  Cache
```

### Data pipeline

```text
Excel / CSV / Government Data
              │
              ▼
         ETL Pipeline
              │
              ▼
        Data Validation
              │
              ▼
         PostgreSQL
              │
              ▼
            API
              │
              ▼
      Interactive Dashboard
```

---

# Overall Design Pattern

The platform's core strength is the combination of the following patterns.

## 1. Hierarchical data navigation

```text
Country
→ Sector
→ Subsector
→ Indicator
→ Time
```

## 2. Cross-filtered charts

```text
Click Chart A
      ↓
Updates Chart B
      ↓
Updates Chart C
      ↓
Updates Table
```

## 3. Multiple analytical views of the same dataset

```text
Same Data
   │
   ├── KPI
   ├── Line Chart
   ├── Pie Chart
   ├── Heatmap
   ├── Table
   └── Geographic Map
```

## 4. Drill-down analysis

```text
National
   ↓
Province
   ↓
Division
   ↓
District
   ↓
Individual Indicator
```

---

# Recommended Development Roadmap

## Phase 1 — Foundation

- Database
- Data import system
- REST API
- Authentication (optional)

## Phase 2 — Core Dashboard

- Navigation
- KPI cards
- GDP chart
- Pie chart
- Trend charts

## Phase 3 — Advanced Analytics

- Cross-filtering
- Heat tables
- Time range selection
- Drill-down functionality

## Phase 4 — Geographic Analytics

- Maps
- Province/district filters
- GeoJSON integration

## Phase 5 — Data Portal Features

- Downloads
- API access
- Metadata
- Dataset catalog
- Automated updates

---

# Final Recommendation

Rather than reproducing each dashboard page independently, build a generic metadata-driven dashboard engine.

The architecture should treat the following as configurable entities:

- Datasets
- Dimensions
- Measures
- Filters
- Chart configurations
- Geographic hierarchies
- Time frequencies

This allows GDP, agriculture, taxation, trade, population, employment, health, education, and other domains to become separate modules powered by the same underlying analytical infrastructure.

## Recommended Technology Stack

```text
Frontend
├── React / Next.js
├── TypeScript
├── Tailwind CSS
├── Apache ECharts
├── TanStack Query
└── Leaflet

Backend
├── FastAPI
├── PostgreSQL
├── Redis
├── Celery
└── SQLAlchemy

Data Engineering
├── Python
├── Pandas / Polars
├── ETL pipelines
├── Data validation
└── Scheduled updates

Infrastructure
├── Docker
├── Nginx
├── CI/CD
└── Cloud deployment
```

The recommended end product is a reusable national statistics and economic intelligence platform rather than a collection of isolated dashboards.
