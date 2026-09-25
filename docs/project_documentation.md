# Loan Portfolio & Financial Decision Engine

## Project Documentation

> **Advanced Excel | Banking & Financial Analytics**

## 1. Project Overview

The **Loan Portfolio & Financial Decision Engine** is an advanced
Microsoft Excel analytics project designed to simulate how a financial
institution can monitor a lending portfolio, analyze credit and
delinquency risk, perform financial modelling, forecast portfolio
activity, and support borrower-level financial decisions.

The project uses **10,000 synthetic loan records** covering
**2023--2026**.

### Project Objectives

-   Monitor overall loan portfolio exposure
-   Analyze outstanding principal
-   Identify credit and delinquency risk
-   Compare performance across loan types and regions
-   Analyze default exposure
-   Build borrower-level loan calculations
-   Generate amortization schedules
-   Evaluate interest-rate scenarios
-   Forecast future loan origination
-   Support prepayment and repayment decisions
-   Present management KPIs through an interactive dashboard
-   Improve dashboard usability through VBA automation

## 2. Business Problem

Financial institutions manage loans with different borrower profiles,
products, interest rates, repayment schedules, credit scores,
delinquency levels, and outstanding balances. This project combines
descriptive analytics, risk analysis, financial modelling, forecasting,
scenario analysis, decision support, and executive reporting in a single
Excel solution.

## 3. Dataset

  Attribute               Description
  ----------------------- -------------------------------
  Dataset Type            Synthetic loan portfolio
  Number of Records       10,000
  Analysis Period         2023--2026
  Raw File                `data/loan_portfolio_raw.csv`
  Primary Analysis Tool   Microsoft Excel

The raw dataset is maintained separately from the analytical workbook so
the source data remains independently available.

## 4. Analytical Workflow

``` text
Raw Data
    ↓
Data Cleaning & Validation
    ↓
Loan-Level Financial Calculations
    ↓
Amortization Analysis
    ↓
Portfolio Analysis
    ↓
Risk & Delinquency Analysis
    ↓
Forecasting
    ↓
Scenario Analysis
    ↓
Financial Decision Engine
    ↓
Executive Dashboard
```

## 5. Workbook Architecture

  -----------------------------------------------------------------------
  Worksheet                           Purpose
  ----------------------------------- -----------------------------------
  `01_README`                         Project overview, instructions and
                                      workbook navigation

  `02_RAW_DATA`                       Original loan-level dataset

  `03_DATA_CLEANING`                  Data quality checks and validation

  `04_LOAN_CALCULATOR`                EMI and borrower-level loan
                                      calculations

  `05_AMORTIZATION`                   Loan repayment and amortization
                                      schedule

  `06_PORTFOLIO_ANALYSIS`             Portfolio KPIs and exposure
                                      analysis

  `07_RISK_ANALYSIS`                  Credit risk, default and
                                      delinquency analysis

  `08_FORECASTING`                    Portfolio forecasting

  `09_SCENARIO_ANALYSIS`              Interest-rate and financial what-if
                                      analysis

  `10_DECISION_ENGINE`                Borrower financial decision support

  `11_DASHBOARD`                      Interactive executive portfolio
                                      dashboard
  -----------------------------------------------------------------------

## 6. Data Preparation & Validation

The data preparation layer focuses on data type validation,
missing-value checks, duplicate detection, logical validation, numerical
range checks, category validation, date validation, and financial-field
consistency. Excel Tables and structured references support scalable
formulas and consistent analytical references.

## 7. Portfolio Analysis

Core metrics include Total Loan Amount, Outstanding Principal, Number of
Loans, Average Interest Rate, Average Credit Score, Default Rate, Loan
Exposure by Product, and Regional Portfolio Exposure.

Primary segmentation dimensions are **Loan Type**, **Region**, and
**Risk Category**.

### Baseline Portfolio KPIs

  KPI                              Value
  ----------------------- --------------
  Total Loan Amount         ₹1,673.46 Cr
  Outstanding Principal     ₹1,266.04 Cr
  Default Rate                     2.02%
  Average Credit Score            709.96
  Average Interest Rate           11.67%

These values represent the complete portfolio when all dashboard filters
are set to `All`.

## 8. Credit Risk & Delinquency Analysis

Risk analysis uses Risk Category, Default Flag, Credit Score, Days Past
Due, Outstanding Principal, Loan Type, and Region.

### Delinquency Buckets

-   Current
-   1--30 Days Past Due
-   31--60 Days Past Due
-   61--90 Days Past Due
-   90+ Days Past Due

This supports analysis of both account volume and financial exposure
across delinquency stages.

### Default Exposure

Default exposure measures outstanding principal associated with loans
marked as defaulted, providing a financial measure of portfolio risk in
addition to default counts and rates.

## 9. Loan Calculator

The Loan Calculator provides borrower-level calculations using inputs
such as Loan Amount, Interest Rate, and Loan Tenure.

Financial functions demonstrated include `PMT`, `IPMT`, `PPMT`, `PV`,
and `FV`.

## 10. Amortization Analysis

The amortization model breaks scheduled repayment into Opening
Principal, EMI, Interest Component, Principal Component, and Closing
Principal, showing how balances and repayment composition change through
the loan term.

## 11. Forecasting

The forecasting component estimates future loan origination activity
from historical portfolio patterns. The dashboard includes a **12-month
loan origination forecast** at portfolio level.

Forecast results are analytical estimates and are not guaranteed future
outcomes.

## 12. Scenario Analysis

The scenario-analysis layer evaluates how changes in assumptions such as
interest rate, loan tenure, and repayment conditions can affect loan
economics.

Techniques demonstrated include **What-If Analysis**, **Goal Seek**, and
**Scenario Analysis**.

## 13. Financial Decision Engine

The Financial Decision Engine extends the project beyond descriptive
analytics by using borrower and loan information to support evaluation
of choices such as continuing an existing repayment plan versus
considering prepayment.

The component is intended for educational and portfolio demonstration
and is not financial advice.

## 14. Executive Dashboard

### KPI Cards

-   Total Loan Amount
-   Outstanding Principal
-   Default Rate
-   Average Credit Score
-   Average Interest Rate

### Visualizations

-   Outstanding Principal by Loan Type
-   Regional Default Rate
-   Risk Category Distribution
-   Loans by Delinquency / DPD
-   12-Month Loan Origination Forecast
-   Default Exposure by Loan Type

## 15. Dashboard Interactivity

The dashboard uses three primary filters: **Loan Type**, **Region**, and
**Risk Category**. Applicable KPIs and charts dynamically recalculate as
selections change.

Some visualizations intentionally retain their own comparison dimension.
For example, a Loan Type comparison can continue showing all loan types
while responding to Region and Risk Category selections. This preserves
meaningful category comparison.

## 16. VBA Automation --- Reset Filters

The dashboard includes a VBA-powered **RESET FILTERS** button that
restores:

-   Loan Type = `All`
-   Region = `All`
-   Risk Category = `All`

### VBA Code

``` vb
Sub ResetDashboardFilters()

    Application.ScreenUpdating = False

    With Worksheets("11_DASHBOARD")
        .Range("L3").Value = "All"
        .Range("M3").Value = "All"
        .Range("N3").Value = "All"
    End With

    Application.Calculate
    Application.ScreenUpdating = True

End Sub
```

### Why the Macro Is Used

Instead of manually returning three filter selections to `All`, the
macro restores the complete portfolio view with one click and
recalculates the workbook.

The workbook is distributed as
`Loan_Portfolio_Financial_Decision_Engine.xlsm`. Users may need to
enable macros in Microsoft Excel for the Reset Filters button to work.

## 17. Excel Techniques Demonstrated

### Data Management

Excel Tables, structured references, data validation, conditional
formatting, and data quality checks.

### Analytical Functions

`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `SUMPRODUCT`, `IF`, `IFERROR`,
`XLOOKUP`, and `FILTER`.

### Financial Functions

`PMT`, `IPMT`, `PPMT`, `PV`, and `FV`.

### Risk Analytics

Default Rate, Days Past Due analysis, risk segmentation, default
exposure, and regional risk analysis.

### Advanced Analysis

What-If Analysis, Goal Seek, Scenario Analysis, and forecasting.

### Visualization & Automation

KPI cards, dynamic charts, executive dashboard, VBA macro, and automated
filter reset.

## 18. Dashboard Calculation Design

The dashboard uses structured Excel formulas and helper analytical
ranges. Dynamic calculations use functions such as `SUMPRODUCT`,
`FILTER`, `AVERAGE`, `SUMIFS`, `COUNTIFS`, and `IFERROR`.

Filter selections drive applicable KPI and chart calculations, while
comparison-oriented charts can intentionally ignore the filter
corresponding to their own comparison dimension.

## 19. Business Value

The project demonstrates the workflow:

**Data → Analysis → Risk Measurement → Financial Modelling → Forecasting
→ Decision Support → Executive Reporting**

The solution provides a consolidated view of portfolio size, outstanding
exposure, product concentration, regional performance, credit quality,
delinquency, default exposure, future origination trends, loan
economics, financial scenarios, and borrower-level decision support.

## 20. Repository Structure

``` text
Loan-Portfolio-Financial-Decision-Engine/
│
├── README.md
├── Loan_Portfolio_Financial_Decision_Engine.xlsm
├── .gitignore
│
├── data/
│   └── loan_portfolio_raw.csv
│
├── screenshots/
│   ├── 01_executive_dashboard.png
│   ├── 02_portfolio_analysis.png
│   ├── 03_risk_analysis.png
│   ├── 04_loan_calculator.png
│   ├── 05_amortization.png
│   ├── 06_forecasting.png
│   ├── 07_scenario_analysis.png
│   └── 08_decision_engine.png
│
└── docs/
    └── project_documentation.md
```

## 21. Opening the Project

1.  Download `Loan_Portfolio_Financial_Decision_Engine.xlsm`.
2.  Open it in Microsoft Excel Desktop.
3.  Enable macros if you want to use Reset Filters.
4.  Review `01_README` for workbook navigation.
5.  Open `11_DASHBOARD`.
6.  Use Loan Type, Region, and Risk Category filters.
7.  Click **RESET FILTERS** to return to the complete portfolio view.

## 22. Limitations & Disclaimer

-   The dataset is synthetic.
-   Results do not represent actual banking performance.
-   Forecasts are analytical estimates rather than guaranteed outcomes.
-   The Financial Decision Engine is for educational and portfolio
    demonstration only.
-   VBA functionality requires a compatible Microsoft Excel environment
    with macros enabled.
-   No real customer, borrower, bank, or financial institution data is
    used.
