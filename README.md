# Sales Forecast Accuracy Analysis

A Python pipeline that ingests Walmart sales forecast data, computes forecast accuracy metrics, and exports a professionally formatted Excel report with **live formulas**.

## Metrics Calculated

| Metric | Formula | Interpretation |
|--------|---------|----------------|
| **MAD** (Mean Absolute Deviation) | mean(\|actual − forecast\|) | Average magnitude of forecast errors |
| **MAPE** (Mean Absolute Percentage Error) | mean(\|actual − forecast\| / \|actual\|) × 100 | Percentage accuracy of forecasts |
| **Sales-to-Forecast Ratio** | Σ actual / Σ forecast | Values < 1 indicate over-forecasting |
| **Forecast Bias** | mean(actual − forecast) | Positive = under-forecast, Negative = over-forecast |

## Project Structure

```
digital_operations/
├── ingestion_etl.py          # ETL: reads Excel, cleans & standardizes data
├── forecast_metrics.py       # Computes MAD, MAPE, Ratio, Bias (console report)
├── export_metrics.py         # Exports multi-sheet Excel report with formulas
├── sales_forecast_data.xlsx  # Source data (Walmart sales & forecasts)
├── .gitignore
└── README.md
```

### File Descriptions

- **`ingestion_etl.py`** — Reads the "Data" sheet from the source Excel file, standardizes column names, coerces types, and drops empty rows. Exposes `ingest_and_etl()` for use by other modules.
- **`forecast_metrics.py`** — Calculates the four forecast metrics at both the overall and grouped (store × department) levels. Exposes `calculate_metrics()` and prints a formatted console report when run directly.
- **`export_metrics.py`** — Builds a three-sheet Excel workbook using live `SUMPRODUCT`-based formulas (not static values), with formatted headers, number formats, borders, and frozen panes.

## Prerequisites

- Python 3.9+
- Dependencies (install in a virtual environment):

```bash
python3 -m venv venv
source venv/bin/activate
pip install pandas openpyxl
```

## Usage

**Generate the Excel report:**

```bash
python3 export_metrics.py
```

This creates `forecast_metrics_report.xlsx` with three sheets:

| Sheet | Contents |
|-------|----------|
| **By Store_Dept** | Metrics per store/department combination (formulas) |
| **Overall** | Aggregate metrics across all data (formulas) |
| **Raw Data** | Cleaned source data + formula columns (Error, \|Error\|, APE) |

**Print a console report instead:**

```bash
python3 forecast_metrics.py
```

**Run the ETL step alone (verbose output):**

```bash
python3 ingestion_etl.py
```

## Notes

- All metric cells in the Excel report use live formulas — the spreadsheet recalculates if source data changes.
- Rows with zero actual sales are excluded from MAPE calculations (division by zero).
- Each module is quiet when imported; verbose output only appears when run directly via `__main__`.
