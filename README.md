# Dynamic Three-Statement Financial Model & Corporate Credit Analysis (Case Study)





## Project Overview

This repository contains a dynamically linked, 10-year **Three-Statement Financial Model (`2020–2029F`)** built in Microsoft Excel following the [Corporate Finance Institute (CFI) Three-Statement Financial Modeling framework](https://www.google.com/search?q=http://www.youtube.com/watch%3Fv%3DUMYDxmiVin4) ([corporatefinanceinstitute.com](https://corporatefinanceinstitute.com/)).

Starting from 5 years of historical financial statements (`2020–2024`), the model constructs a fully integrated **5-year Forecast Period (`2025F–2029F`)** connecting the **Income Statement**, **Balance Sheet**, and **Cash Flow Statement** through dynamic **Working Capital**, **Depreciation / PP&E**, and **Debt & Interest ("Corkscrew") Supporting Schedules** with a real-time **Balance Sheet Integrity Check (`Assets = Liabilities + Shareholder's Equity`, Check = `0` / `OK`)**.

To align with **Investment Grade Private Credit, Asset Management, and Corporate Debt Monitoring** workflows, this project also evaluates the borrower's **EBITDA generation**, **Cash Flow Conversion**, **Scheduled Debt Paydown (`$20M` in `2027F`)**, **Gross Debt / EBITDA Leverage**, and **Interest Coverage (`EBITDA / Interest Expense`)**.

---

## 1. Integrated Three-Statement Model (`2020–2029F`)

The model follows CFI's single-sheet vertically stacked architecture, separating **Inputs/Assumptions**, **Processing (3 Financial Statements & Supporting Schedules)**, and **Outputs (Financial Charts & Credit Metrics)**.



### Key Highlights Across the 5-Year Forecast (`2025F–2029F`)

| Financial Metric (`$000's`) | FY2024 (Actual) | FY2025F | FY2026F | FY2027F | FY2028F | FY2029F | 5-Yr Forecast Trend & Credit Takeaway |
| :--- | ---: | ---: | ---: | ---: | ---: | ---: | :--- |
| **Revenue** | `$150,772` | `$165,849` | `$182,434` | `$200,678` | `$220,745` | `$242,820` | Steady `10.0%` YoY top-line expansion across the forecast horizon |
| **Gross Profit** | `$94,062` | `$96,193` | `$96,690` | `$100,339` | `$141,277` | `$157,833` | Margin dip in `2025F–2027F` (`50%` floor) followed by expansion to `65%` |
| **EBITDA** | `$57,405` | `$52,998` | `$50,676` | `$51,224` | `$88,750` | `$101,554` | Strong operating cash earnings cushioning fixed charges |
| **Net Earnings (Net Income)** | `$28,227` | `$26,543` | `$24,401` | `$25,209` | `$52,749` | `$61,839` | Profit more than doubles between `2024` and `2029F` |
| **Operating Cash Flow (CFO)** | `$43,479` | `$36,501` | `$33,125` | `$33,652` | `$69,849` | `$74,464` | Consistently funds `$15,000` annual maintenance CapEx internally |
| **Closing Cash Balance** | `$139,549` | `$161,050` | `$179,174` | `$177,827` | `$232,676` | `$292,140` | Liquidity grows even after `$20,000` debt repayment in `2027F` |
| **Total Debt** | `$30,000` | `$30,000` | `$30,000` | `$10,000` | `$10,000` | `$10,000` | `$20,000` principal paydown executed in `2027F` from internal cash |
| **Balance Sheet Check** | `OK (0)` | `OK (0)` | `OK (0)` | `OK (0)` | `OK (0)` | `OK (0)` | `Total Assets = Total Liabilities + Shareholder's Equity` across all years |

---

## 2. Supporting Schedules & Corkscrew Roll-Forwards

To avoid hardcoding Balance Sheet and Cash Flow items, three dedicated supporting schedules drive the forecast period:



### How the Supporting Schedules Connect the Statements

* **Working Capital Schedule (`Rows 88–93`):**
  * Forecasts **Accounts Receivable** (`Revenue × AR Days / 365` at `18 days`), **Inventory** (`COGS × Inventory Days / 365` scaling from `80` to `100 days`), and **Accounts Payable** (`COGS × AP Days / 365` at `37 days`).
  * Computes **Net Working Capital (`NWC = AR + Inventory – AP`)** and links **Change in NWC (`Current NWC – Prior NWC`)** directly into Operating Cash Flow (`Row 68`).
* **Depreciation & PP&E Corkscrew Schedule (`Rows 95–99`):**
  * Rolls forward fixed assets using a corkscrew structure: `Closing PP&E = Opening PP&E + CapEx ($15,000/yr) – Depreciation (35% of Opening PP&E)`.
  * Feeds **Depreciation & Amortization** to the Income Statement (`Row 32`) and Cash Flow Statement add-back (`Row 67`), and **Closing PP&E** to the Balance Sheet (`Row 47`).
* **Debt, Interest & Credit Monitoring Schedule (`Rows 101–105`):**
  * Tracks `Opening Debt + Issuance / (Repayment) = Closing Debt`, modeling a `$20,000` (`$20M`) debt paydown in `2027F` (`Row 103`).
  * Computes **Interest Expense** dynamically on average debt (`AVERAGE(Opening Debt, Closing Debt) × 10% Interest Rate`), reducing annual interest expense from `$3,000` (`2025F–2026F`) to `$1,000` (`2028F–2029F`).
  * Demonstrates an **Investment Grade credit profile**, with **Gross Leverage (`Debt / EBITDA`)** deleveraging from `1.95x` (`2020`) to `0.52x` (`2024`) and `0.10x` (`2029F`), while **Interest Coverage (`EBITDA / Interest`)** remains above `16.9x` throughout the forecast period.

---

## 3. Financial Performance & Cash Flow Visualizations

The model outputs dynamic combo and grouped bar charts tracking top-line growth, margin cyclicality, and cash flow composition across all 10 years:



* **Revenue & Gross Margin Trajectory (Left Chart):** Illustrates how COGS inflation in `2025F–2027F` (`42% → 50%` of revenue) temporarily compresses Gross Profit Margin to `50.0%` before operational efficiencies in `2028F–2029F` (`COGS` at `36%–35%`) expand Gross Margin to `65.0%`.
* **Cash Flow Waterfall & Debt Service Capacity (Right Chart):** Shows that **Operating Cash Flow** (`$33.1M–$74.5M` per year in the forecast period) comfortably covers annual **Capital Expenditures (`-$15.0M`)** every year, generating substantial **Free Cash Flow (`FCF = CFO – CapEx`)** to self-fund the `-$20.0M` debt repayment in `2027F` without external refinancing.

---

## 4. Core Excel Formulas & Modeling Architecture

| Model Section | Excel Formula Used in Forecast (`Col I = 2025F`) | Accounting & Financial Logic |
| :--- | :--- | :--- |
| **Revenue Forecast** | `=H26*(1+I8)` | Prior Year Revenue × `(1 + YoY Revenue Growth Assumption)` |
| **Accounts Receivable** | `=I26*I15/I21` | `Revenue × AR Days (18) / Days in Period (365)` |
| **Inventory & AP** | `=I27*I16/I21` and `=I27*I17/I21` | `COGS × Inventory Days / 365` and `COGS × AP Days / 365` |
| **Operating Cash Flow** | `=I66+I67-I68` | `Net Earnings + Depreciation & Amortization – Change in NWC` |
| **Interest Expense** | `=AVERAGE(I102,I104)*I13` | `Average(Opening Debt, Closing Debt) × Interest Rate (10%)` |
| **Retained Earnings** | `=H56+I38` | `Prior Retained Earnings + Current Net Earnings (– Dividends)` |
| **Balance Sheet Check** | `=IFERROR(IF(ABS(I60)>1,"ERROR","OK"),"OK")` | Verifies `Total Liabilities & Equity – Total Assets = 0` |

---

## 5. Repository Structure & How to Use

```text
├── CFI-Case-Study-Three-Statement-Model.xlsx   # Complete 10-Year Dynamic Excel Financial Model
├── images/
│   ├── 1_three_statement_model_overview.png    # Snapshot of Income Statement, Balance Sheet & Cash Flow
│   ├── 2_supporting_schedules_corkscrew.png    # Working Capital, PP&E Corkscrew & Debt/Credit Schedule
│   └── 3_financial_performance_charts.png      # Revenue/Margin & Cash Flow Breakdown Charts
└── README.md                                   # Project documentation & credit analysis summary
```

1. **Open the Workbook:** Download and open `CFI-Case-Study-Three-Statement-Model.xlsx` in Microsoft Excel.
2. **Run Scenario & Stress Tests:** Modify any driver in the **Key Assumptions block (`Rows 8–20`, Columns `I:M`)**—such as setting **Revenue Growth (`Row 8`)** to `0%` (flatline stress test) or increasing **Interest Rate (`Row 13`)**—to observe real-time impacts on EBITDA, Free Cash Flow, and Closing Cash while the Balance Sheet remains dynamically balanced (`Check = OK`).

---

## Credits & Attribution

* **Case Study & Modeling Framework:** Built following the [Corporate Finance Institute (CFI) Three-Statement Modeling Tutorial by CFI CEO Tim Vipond ([corporatefinanceinstitute.com](https://corporatefinanceinstitute.com/)).
* **Educational Use Notice:** The baseline template and historical figures (`2020–2024`) are copyright material of CFI Education Inc., used here strictly for non-commercial educational and portfolio demonstration purposes.
