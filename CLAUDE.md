# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
A comprehensive financial portfolio analysis system that processes stock prices, bookings, and generates various financial reports. The system imports data from Yahoo Finance, processes transactions, and exports detailed Excel reports for portfolio analysis.

## Development Commands

### Running the Application
```bash
# Direct execution
python depot.py

# Via batch file (production/scheduled execution)
start_depot.bat
```

### Dependencies
```bash
# Install all dependencies
pip install -r requirements.txt

# Key dependency: ahlib (custom library)
pip install ahlib
```

### Virtual Environment
The batch file expects a virtual environment at `.\.venv\Scripts\python.exe`

## Architecture

### Core Data Processing Pipeline
The application follows a structured 8-step data processing pipeline:

1. **Initialization**: Load settings from `depot.ini`
2. **Instruments Import**: Load and process financial instruments from Excel files
3. **Prices Processing**: Import/update price data using yfinance, store in `prices.parquet`
4. **Bookings Processing**: Process transaction bookings from Excel files
5. **Shares Calculation**: Calculate daily share positions by bank from bookings
6. **Values Calculation**: Calculate portfolio values (shares × prices)
7. **Provisions Processing**: Handle broker fees and provisions
8. **Export Generation**: Generate 30+ different Excel reports

### Key Data Structures
- **MultiIndex DataFrames**: Primary data structure using (date, wkn) or (date, bank) indexes
- **Time Series Processing**: Daily, monthly, and yearly aggregations
- **Portfolio Tracking**: Positions, values, gains/losses by instrument and bank

### Configuration System
`depot.ini` controls:
- File paths for input/output
- Export settings for 30+ different report types
- Each export can be enabled/disabled with custom formatting
- Network paths for shared data access

### Key Functions from ahlib
- `screen_and_log()`: Unified logging to console and file
- `export_df_to_excel()`: Excel export with formatting
- `import_parquet()`: Parquet file handling
- `settings_import()`: Configuration file parsing

### Data Transformation Functions
- `df_to_eom()`: Reduce to month-end data points
- `df_to_eoy()`: Reduce to year-end data points
- `df_2D_sum_per_period()`: Aggregate data by time periods
- `df_transform_each_line_to_percentage()`: Convert to percentage distributions

## Production Deployment
- Uses Windows Task Scheduler via `start_depot.bat`
- Automatic git pull before execution
- Comprehensive logging with monthly log files
- Automatic cleanup of logs older than 120 days
- UTF-8 encoding handling for Windows environments

## File Locations
- **Input files**: Network paths on `\\WIN-H7BKO5H0RMC\Dataserver\`
- **Output files**: Excel reports to shared network location
- **Local data**: `prices.parquet` for price caching
- **Logs**: Both `depot.log` and monthly batch logs in `logs/` directory