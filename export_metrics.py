"""
Export Forecast Metrics to Excel
Reads sales forecast data, computes grouped and overall metrics,
and writes a formatted multi-sheet Excel report.
"""

import pandas as pd
from ingestion_etl import ingest_and_etl
from forecast_metrics import calculate_metrics
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, numbers


def format_worksheet(ws, num_formats=None):
    """Apply formatting to a worksheet: bold headers, auto-width, number formats."""
    # Bold header row
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Auto-fit column widths
    for col_idx, col in enumerate(ws.columns, 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for cell in col:
            val = cell.value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Apply number formats where specified
    if num_formats:
        for cell in ws[1]:
            fmt = num_formats.get(cell.value)
            if fmt:
                for row_cell in ws.iter_rows(
                    min_row=2, max_row=ws.max_row,
                    min_col=cell.column, max_col=cell.column
                ):
                    row_cell[0].number_format = fmt


def main():
    # 1. Ingest data
    df = ingest_and_etl("sales_forecast_data.xlsx")
    if df is None or df.empty:
        print("❌ Failed to load data. Exiting.")
        return

    # 2. Compute metrics
    grouped_metrics = calculate_metrics(df, group_cols=["store", "dept"])
    overall_metrics = calculate_metrics(df)

    # 3. Write to Excel
    output_file = "forecast_metrics_report.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Sheet 1: Grouped metrics
        grouped_metrics.to_excel(writer, sheet_name="By Store_Dept", index=False)

        # Sheet 2: Overall metrics (single row)
        overall_df = overall_metrics.to_frame().T  # Series → single-row DataFrame
        overall_df.to_excel(writer, sheet_name="Overall", index=False)

        # Sheet 3: Raw data
        df.to_excel(writer, sheet_name="Raw Data", index=False)

    # 4. Format all sheets
    from openpyxl import load_workbook
    wb = load_workbook(output_file)

    # Common number format map
    metrics_num_fmt = {
        "MAD": "#,##0.00",
        "MAPE": "0.00\"%\"",
        "Sales_to_Forecast_Ratio": "0.0000",
        "Forecast_Bias": "#,##0.00",
    }

    # Format Sheet 1: By Store_Dept
    format_worksheet(wb["By Store_Dept"], num_formats=metrics_num_fmt)

    # Format Sheet 2: Overall
    format_worksheet(wb["Overall"], num_formats=metrics_num_fmt)

    # Format Sheet 3: Raw Data
    raw_num_fmt = {
        "store": "0",
        "dept": "0",
        "year": "0",
        "week": "0",
        "forecast": "#,##0.00",
        "weekly_sales": "#,##0.00",
    }
    format_worksheet(wb["Raw Data"], num_formats=raw_num_fmt)

    wb.save(output_file)
    print(f"\n✅ Report saved to '{output_file}'")
    print(f"   Sheets: By Store_Dept | Overall | Raw Data")
    print(f"   Grouped rows: {len(grouped_metrics)}")
    print(f"   Raw data rows: {len(df)}")


if __name__ == "__main__":
    main()
