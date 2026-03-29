"""
Forecast Metrics Calculator
Computes MAD, MAPE, Sales-to-Forecast Ratio, and Forecast Bias
from sales forecast data processed by ingestion_etl.
"""

import numpy as np
import pandas as pd

from ingestion_etl import ingest_and_etl


def calculate_metrics(df, group_cols=None):
    """
    Calculate forecast accuracy metrics.

    Metrics
    -------
    MAD   – Mean Absolute Deviation  = mean(|actual − forecast|)
    MAPE  – Mean Absolute Percentage Error  = mean(|actual − forecast| / |actual|) × 100
    Sales-to-Forecast Ratio  = Σ actual / Σ forecast
    Forecast Bias  = mean(actual − forecast)

    Parameters
    ----------
    df : pd.DataFrame
        Must contain 'weekly_sales' and 'forecast' columns.
    group_cols : list of str or None
        Columns to group by (e.g. ['store', 'dept']).

    Returns
    -------
    pd.DataFrame (grouped) or pd.Series (overall)
    """
    df = df.copy()
    df["error"] = df["weekly_sales"] - df["forecast"]
    df["abs_error"] = df["error"].abs()

    # Exclude rows where actual sales are zero (MAPE is undefined)
    df = df[df["weekly_sales"] != 0].copy()

    def _compute(group):
        return pd.Series({
            "MAD": group["abs_error"].mean(),
            "MAPE": (group["abs_error"] / group["weekly_sales"].abs()).mean() * 100,
            "Sales_to_Forecast_Ratio": (
                group["weekly_sales"].sum() / group["forecast"].sum()
                if group["forecast"].sum() != 0
                else np.nan
            ),
            "Forecast_Bias": group["error"].mean(),
        })

    if group_cols:
        grouped = df.groupby(group_cols, observed=True)
        return grouped.apply(_compute, include_groups=False).reset_index()

    return _compute(df)


def main():
    """Load data, compute metrics, and print a console report."""
    df = ingest_and_etl("sales_forecast_data.xlsx")
    if df is None or df.empty:
        print("Failed to load data. Exiting.")
        return

    sep = "=" * 60
    print(f"\n{sep}")
    print("  FORECAST METRICS REPORT")
    print(f"{sep}")
    print(f"  Records loaded: {len(df)}")
    print(sep)

    # Grouped metrics
    grouped = calculate_metrics(df, group_cols=["store", "dept"])
    if grouped is not None and not grouped.empty:
        print("\n  Metrics by Store & Department:")
        print(f"  {'-' * 56}")
        print(grouped.round(4).to_string(index=False))

    # Overall metrics
    print(f"\n{sep}")
    overall = calculate_metrics(df)
    print("  Overall Metrics:")
    print(f"  {'-' * 56}")
    for metric, value in overall.items():
        print(f"    {metric:<30} {value:>12.4f}")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
