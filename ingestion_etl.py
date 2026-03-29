import pandas as pd

def ingest_and_etl(file_path):
    """
    Ingests an Excel file, performs basic ETL, and loads it into a Pandas DataFrame.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame: The processed DataFrame.
    """
    try:
        # Ingestion: Read the Excel file
        df = pd.read_excel(file_path)
        print(f"Successfully ingested data from {file_path}. Initial shape: {df.shape}")
        print("Initial DataFrame head:")
        print(df.head())

        # Basic ETL steps (can be expanded based on requirements):
        # 1. Rename columns for easier access (e.g., lowercase, replace spaces)
        df.columns = df.columns.str.lower().str.replace(' ', '_')
        print("\nDataFrame columns after renaming:")
        print(df.columns.tolist())

        # 2. Handle missing values (e.g., fill with 0 or mean, or drop rows)
        # For now, let's assume no critical missing values that prevent calculation
        # df = df.fillna(0) # Example: fill NaNs with 0

        # 3. Convert data types if necessary
        # Pandas often infers correctly, but explicit conversion can be done here
        # Example: df['date_column'] = pd.to_datetime(df['date_column'])

        # 4. Remove duplicates (if applicable)
        # df.drop_duplicates(inplace=True)

        print(f"\nETL completed. Final DataFrame shape: {df.shape}")
        print("Final DataFrame head:")
        print(df.head())

        return df

    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
        return None
    except Exception as e:
        print(f"An error occurred during ingestion and ETL: {e}")
        return None

if __name__ == "__main__":
    excel_file_path = 'sales_forecast_data.xlsx' # Assuming the Excel file is in the same directory
    processed_df = ingest_and_etl(excel_file_path)

    if processed_df is not None:
        print("\nData loaded into DataFrame successfully and ready for further analysis.")
        # You can now proceed with MAD, MAPE, etc., using 'processed_df'
        # For example:
        # mad, mape, sales_to_forecast_ratio, forecast_bias = calculate_metrics(processed_df)
    else:
        print("\nFailed to process data.")
