import pandas as pd
import sys
import os

try:
    from ahlib import create_extended_logger, settings_import
except ImportError:
    print("Error: ahlib library not found. Please ensure the virtual environment is set up and ahlib is installed.")
    sys.exit(1)

def instruments_import_for_check(filename, logger):
    """
    Replicates the core logic of instruments_import from depot.py to load and preprocess
    the instruments data, specifically handling WKNs.
    """
    try:
        if not filename.endswith(('.xlsx', '.xls')):
            raise ValueError(f"The file '{filename}' is not an Excel file.")
        
        # Read the first four columns and set the first column (WKN) as index
        df = pd.read_excel(filename, usecols=[0, 1, 2, 3], header=None, names=['wkn_raw', 'ticker', 'instrument_name', 'default_value'])
        
        # Drop rows where 'wkn_raw' is entirely NaN (empty rows at the end of Excel)
        df.dropna(subset=['wkn_raw'], inplace=True)

        # Convert wkn to string and then to lowercase, handling potential non-string types
        df['wkn'] = df['wkn_raw'].astype(str).str.strip().str.lower()
        
        # Set the processed 'wkn' as the index
        df.set_index('wkn', inplace=True)
        
        # Convert ticker to lowercase
        df['ticker'] = df['ticker'].astype(str).str.lower()
        
        # Set the column names (excluding the raw wkn column)
        df = df[['ticker', 'instrument_name', 'default_value']]
        
        return df
    except FileNotFoundError:
        logger.error(f"The file '{filename}' was not found.")
        return None
    except ValueError as ve:
        logger.error(f"{ve}")
        return None
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
        return None

def main():
    logger = create_extended_logger("check_instruments.log", script_name="check_instruments")
    logger.info("Starting check for instruments.xlsx issues.")

    settings_file = "depot.ini"
    settings = settings_import(settings_file, logger)
    if settings is None:
        logger.error(f"Could not load settings from {settings_file}. Exiting.")
        sys.exit(1)

    instruments_file = settings.get('Files', {}).get('instruments', '')
    if not instruments_file:
        logger.error("Instrument file path not found in depot.ini. Exiting.")
        sys.exit(1)

    instruments_df = instruments_import_for_check(instruments_file, logger)

    if instruments_df is None:
        logger.error("Failed to import instruments DataFrame. Check log for details. Exiting.")
        sys.exit(1)

    # Check for NaN (empty) WKNs after processing
    nan_wkn_in_index = instruments_df.index[instruments_df.index.isna()]
    if not nan_wkn_in_index.empty:
        logger.error(f"Found NaN (empty) WKNs in the instruments file index. These rows will cause issues.")
        print(f"NaN WKNs: {nan_wkn_in_index.tolist()}")
        # Show rows that originally had NaN WKNs
        original_df = pd.read_excel(instruments_file, header=None, names=['wkn_raw', 'ticker', 'instrument_name', 'default_value'])
        problem_rows = original_df[original_df['wkn_raw'].isna()]
        if not problem_rows.empty:
            logger.error("Original rows from Excel with empty WKNs:")
            print(problem_rows)
        
        sys.exit(1)

    # Check for empty string WKNs after stripping whitespace
    empty_string_wkn_in_index = instruments_df.index[instruments_df.index == '']
    if not empty_string_wkn_in_index.empty:
        logger.error(f"Found empty string WKNs in the instruments file index after stripping whitespace. These rows will cause issues.")
        print(f"Empty string WKNs: {empty_string_wkn_in_index.tolist()}")
        sys.exit(1)

    # Check for duplicate WKNs (after lowercasing and stripping)
    duplicate_wkns = instruments_df.index[instruments_df.index.duplicated(keep=False)]
    if not duplicate_wkns.empty:
        logger.error(f"Found duplicate WKNs in the instruments file index: {duplicate_wkns.unique().tolist()}. Duplicate WKNs will cause 'non-unique multi-index' errors.")
        print("Rows with duplicate WKNs:")
        print(instruments_df.loc[duplicate_wkns.unique()])
        sys.exit(1)

    logger.info("No NaN, empty, or duplicate WKNs found in instruments.xlsx after processing. File structure seems valid for WKNs.")
    
    # Check for NaN tickers, as this also appeared in the logs
    nan_tickers = instruments_df[instruments_df['ticker'].isna()]
    if not nan_tickers.empty:
        logger.warning(f"Found NaN tickers for WKNs: {nan_tickers.index.tolist()}. Yfinance will not be able to fetch prices for these.")
        print(f"WKNs with NaN tickers: {nan_tickers.index.tolist()}")

if __name__ == "__main__":
    main()
