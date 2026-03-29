"""
Forecast Metrics Calculator
Computes MAD, MAPE, Sales-to-Forecast Ratio, and Forecast Bias
from sales forecast data processed by ingestion_etl.py
"""

import numpy as np
import pandas as pd
from ingestion_etl import ingest_and_etl


def calculate_metrics(df, group_cols=None):
    """
    Calculate forecast accuracy metrics.

    Metrics:
      - MAD  (Mean Absolute Deviation)
      - MAPE (Mean Absolute Percentage Error)
      - Sales-to-Forecast Ratio
      - Forecast Bias

    Parameters
    ----------
    df : pd.DataFrame  Must contain 'weekly_sales' and 'forecast' columns.
    group_cols : list or None  Columns to group by (e.g. ['store', 'dept']).

    Returns
    -------
    pd.DataFrame (grouped) or pd.Series (overall)
    """
    df = df.copy()
    df["error"] = df["weekly_sales"] - df["forecast"]
    df["abs_error"] = df["error"].abs()

    # Filter out rows where actual sales are zero (MAPE undefined)
    df = df[df["weekly_sales"] != 0].copy()

    if group_cols:
        grouped = df.groupby(group_cols, observed=True)

        def _agg(group):
            return pd.Series(
                {
                    "MAD": group["abs_error"].mean(),
                    "MAPE": (group["abs_error"] / group["weekly_sales"].abs()).mean()
                    * 100,
                    "Sales_to_Forecast_Ratio": (
                        group["weekly_sales"].sum() / group["forecast"].sum()
                        if group["forecast"].sum() != 0
                        else np.nan
                    ),
                    "Forecast_Bias": group["error"].mean(),
                }
            )

        return grouped.apply(_agg, include_groups=False).reset_index()

    # Overall
    return pd.Series(
        {
            "MAD": df["abs_error"].mean(),
            "MAPE": (df["abs_error"] / df["weekly_sales"].abs()).mean() * 100,
            "Sales_to_Forecast_Ratio": (
                df["weekly_sales"].sum() / df["forecast"].sum()
                if df["forecast"].sum() != 0
                else np.nan
            ),
            "Forecast_Bias": df["error"].mean(),
        }
    )


def main():
    """Load data, compute metrics, print report."""
    df = ingest_and_etl("sales_forecast_data.xlsx")

    if df is None or df.empty:
        print("Failed to load data. Exiting.")
        return

    print(f"\n{'=' * 60}")
    print("  FORECAST METRICS REPORT")
    print(f"{'=' * 60}")
    print(f"  Records loaded: {len(df)}")
    print(f"{'=' * 60}")

    # Grouped metrics
    grouped = calculate_metrics(df, group_cols=["store", "dept"])
    if grouped is not None and not grouped.empty:
        print("\n  Metrics by Store & Department:")
        print(f"  {'-' * 56}")
        print(grouped.round(4).to_string(index=False))

    # Overall metrics
    print(f"\n{'=' * 60}")
    overall = calculate_metrics(df)
    print("  Overall Metrics:")
    print(f"  {'-' * 56}")
    for metric, value in overall.items():
        print(f"    {metric:<30} {value:>12.4f}")
    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    main()
