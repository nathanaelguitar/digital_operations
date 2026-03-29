import pandas as pd

def ingest_and_etl(file_path):
    """
    Ingests the Walmart sales forecast Excel file, performs ETL,
    and loads it into a clean Pandas DataFrame.

    Args:
        file_path (str): Path to the Excel file.

    Returns:
        pd.DataFrame: The processed DataFrame with columns:
            store, dept, year, week, forecast, weekly_sales
    """
    try:
        # --- INGESTION ---
        # The "Data" sheet has headers at row index 2 (0-indexed),
        # with actual data starting at row 3.
        df = pd.read_excel(
            file_path,
            sheet_name="Data",
            header=3,          # Row 3 is the header row (0-indexed)
            usecols=range(1, 7) # Columns B through G
        )
        print(f"Ingested '{file_path}' → sheet 'Data'")
        print(f"Raw shape: {df.shape}")
        print(f"Raw columns: {df.columns.tolist()}")
        print(df.head(3).to_string())

        # --- TRANSFORM ---
        # 1. Standardize column names
        df.columns = ["store", "dept", "year", "week", "forecast", "weekly_sales"]

        # 2. Drop rows where all values are NaN (empty rows)
        df.dropna(how="all", inplace=True)

        # 3. Convert numeric columns
        numeric_cols = ["store", "dept", "year", "week", "forecast", "weekly_sales"]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        # 4. Convert store/dept/year/week to integers where possible
        int_cols = ["store", "dept", "year", "week"]
        for col in int_cols:
            df[col] = df[col].astype("Int64")

        # 5. Reset index
        df.reset_index(drop=True, inplace=True)

        # --- SUMMARY ---
        print(f"\n--- ETL Complete ---")
        print(f"Final shape: {df.shape}")
        print(f"Dtypes:\n{df.dtypes}")
        print(f"\nNull counts:\n{df.isnull().sum()}")
        print(f"\nFirst 5 rows:")
        print(df.head().to_string())
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
    df = ingest_and_etl("sales_forecast_data.xlsx")
    if df is not None:
        print("\n✅ DataFrame ready for analysis.")
    else:
        print("\n❌ ETL failed.")
