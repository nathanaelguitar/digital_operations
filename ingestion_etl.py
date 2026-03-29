"""
Ingestion & ETL Module
Reads a Walmart sales forecast Excel file, cleans the data,
and returns a tidy Pandas DataFrame ready for analysis.
"""

import pandas as pd


def ingest_and_etl(file_path, verbose=False):
    """
    Ingest the Walmart sales forecast Excel file and perform ETL.

    Parameters
    ----------
    file_path : str
        Path to the Excel file (must contain a "Data" sheet).
    verbose : bool, optional
        If True, print progress and summary info (default False).

    Returns
    -------
    pd.DataFrame or None
        Cleaned DataFrame with columns:
        store, dept, year, week, forecast, weekly_sales.
        Returns None on failure.
    """
    try:
        # --- INGESTION ---
        # "Data" sheet: headers at row index 3 (0-based), columns B–G.
        df = pd.read_excel(
            file_path,
            sheet_name="Data",
            header=3,
            usecols=range(1, 7),
        )
        if verbose:
            print(f"Ingested '{file_path}' → sheet 'Data'")
            print(f"Raw shape: {df.shape}")
            print(f"Raw columns: {df.columns.tolist()}")
            print(df.head(3).to_string())

        # --- TRANSFORM ---
        df.columns = ["store", "dept", "year", "week", "forecast", "weekly_sales"]

        # Drop fully-empty rows
        df.dropna(how="all", inplace=True)

        # Coerce to numeric
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        # Integer columns where appropriate
        for col in ["store", "dept", "year", "week"]:
            df[col] = df[col].astype("Int64")

        df.reset_index(drop=True, inplace=True)

        # --- SUMMARY ---
        if verbose:
            print(f"\n--- ETL Complete ---")
            print(f"Final shape: {df.shape}")
            print(f"Dtypes:\n{df.dtypes}")
            print(f"\nNull counts:\n{df.isnull().sum()}")
            print(f"\nFirst 5 rows:\n{df.head().to_string()}")
            print(f"\nUnique stores: {sorted(df['store'].dropna().unique())}")
            print(f"Unique depts:  {sorted(df['dept'].dropna().unique())}")

        return df

    except FileNotFoundError:
        print(f"Error: File not found → {file_path}")
        return None
    except Exception as e:
        print(f"Error during ETL: {e}")
        return None


if __name__ == "__main__":
    df = ingest_and_etl("sales_forecast_data.xlsx", verbose=True)
    if df is not None:
        print("\n✅ DataFrame ready for analysis.")
    else:
        print("\n❌ ETL failed.")
