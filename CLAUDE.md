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

## Recent Changes

### 2025-12-07 - Yearly Profitability Analysis Table with Cash Exclusion

**Feature Implemented:**
- New yearly profitability table showing holding periods and annual yields per instrument
- Cash excluded from all profitability calculations to avoid data quality issues
- Custom Excel export with formatted table, column-specific number formats, and 90° rotated headers

**New Components:**

1. **Lines 1329-1428: `profitability_year_table()` function**
   - Creates profitability table with years as rows
   - Two columns per WKN: `{wkn}_days` (holding period) and `{wkn}_yield` (annual return)
   - Uses Time-Weighted Return (TWR) formula: `∏(1 + daily_yield) - 1`
   - Provides transparency for partial-year holdings (e.g., 101 days vs 365 days)

2. **Lines 1142-1173: Cash filtering in `yield_components_day()`**
   - Excludes 'cash' from all yield calculations
   - Filters applied to: values_day, gains_losses, interest_dividends, fees, taxes, transaction_value
   - Prevents issues from missing transaction_value_at_price in cash bookings
   - CM (call money) and FTD (fixed-term deposit) still included with profitability tracking

3. **Lines 1358-1360: Cash filtering in `profitability_year_table()`**
   - Additional filter to exclude cash from yearly aggregation
   - Ensures cash never appears in profitability reports

4. **Lines 2554-2634: Custom openpyxl export for profitability_year.xlsx**
   - Dynamic column formatting based on column name patterns:
     - `_days` columns → Integer format `#,##0`
     - `_yield` columns → Percentage format `0.00%`
   - Excel Table with TableStyleMedium9 style
   - 90° text rotation for all headers (compact display)
   - Custom column widths: year=8, days=4.57, yield=7.71
   - No depot.ini maintenance required when adding new WKNs

**Why Custom Export Instead of export_2D_df_to_excel_format:**
- Cannot use standard 2D export function due to structure incompatibility
- Transformation to (year, wkn) MultiIndex creates duplicate column names after unstack()
- Custom export provides dynamic formatting without ini file maintenance
- As portfolio grows, formatting automatically applies to new WKNs

**Example Output Structure:**
```
year | a0f5uf_days | a0f5uf_yield | 843002_days | 843002_yield | ...
2022 |         101 |       15.39% |         365 |       22.18% | ...
2023 |         250 |       20.15% |         365 |       19.42% | ...
2024 |         365 |       18.73% |         365 |       20.89% | ...
```

**Benefits:**
- Transparent partial-year holdings (days column shows actual holding period)
- Automatic formatting for unlimited number of WKNs
- Clean separation of cash (excluded) vs investment instruments
- Compact Excel table with professional formatting

**Configuration:**
- depot.ini line 57: `profitability_year_to_excel`
- File: `profitability_year.xlsx`
- Note: column_formats in ini are not used (custom formatting in code)

### 2025-12-06 - Intelligent Denominator Selection for Yield Calculations

**Problem Solved:**
- Extreme yield percentages (100%+) occurred on first trading days and during depot transfers
- Traditional formula `yield = gains_losses / portfolio_value` produced misleading results when portfolio_value was near zero or when large transactions occurred

**Solution Implemented:**
- Intelligent denominator selection in `yield_components_day()` function
- For large transactions (|transaction| > 50% of portfolio value), uses transaction value as denominator
- Prevents false extreme yields while maintaining accurate percentage calculations

**Key Changes in depot.py:**

1. **Lines 1098-1211: `yield_components_day()` function**
   - Added parameter: `transaction_value_at_price_df`
   - Intelligent denominator logic:
     ```python
     # Calculate absolute transaction value
     abs_transaction = transaction_value_at_price.abs()

     # Start with portfolio value as denominator
     denominator = portfolio_value

     # For large transactions (>50% of portfolio), use transaction value
     large_transaction_mask = (abs_transaction > portfolio_value * 0.5) & (abs_transaction > 0)
     denominator[large_transaction_mask] = abs_transaction[large_transaction_mask]

     # Apply to all yield components
     yield_price = gains_losses / denominator
     yield_dividends = interest_dividends / denominator
     yield_fees = fees / denominator
     yield_taxes = taxes / denominator
     ```

2. **Line 2366: Function call updated**
   ```python
   yield_components_day_df = yield_components_day(
       gains_losses_before_fees_taxes_day_df,
       fees_df,
       taxes_df,
       interest_dividends_df,
       values_day_df,
       transaction_value_at_price_day_df,  # NEW PARAMETER
       logger
   )
   ```

**Impact:**
- Reduced extreme values from 141 to 10 (real market volatility only)
- Average total return: -0.0004 (effectively 0%, correct for balanced portfolio)
- 89 large transactions now correctly handled
- First trading days show 0% yield (correct, as no prior position exists)
- Depot transfers properly neutralized (balanced bookings)

**Data Quality Requirements:**
- `transaction_value_at_price` must be filled for all buy/sell transactions
- Formula: `transaction_value_at_price = -(delta × price)`
- For depot transfers: Paired bookings must sum to zero
- Interest/dividends go in `interest_dividends` column, NOT transaction_value

## Production Deployment
- Uses Windows Task Scheduler via `start_depot.bat`
- Automatic git pull before execution
- Comprehensive logging with monthly log files
- Automatic cleanup of logs older than 120 days
- UTF-8 encoding handling for Windows environments

## File Locations
- **Input files**: Network paths on `\WIN-H7BKO5H0RMC\Dataserver\`
- **Output files**: Excel reports to shared network location
- **Local data**: `prices.parquet` for price caching
- **Logs**: Both `depot.log` and monthly batch logs in `logs/` directory
