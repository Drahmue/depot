# Depot Project

A comprehensive financial portfolio analysis system that processes stock prices, bookings, and generates detailed financial reports. The system integrates with Yahoo Finance for price data and produces 30+ different Excel reports for portfolio analysis.

## Features

- **Multi-bank portfolio tracking** with position and value calculations
- **Automated price data updates** using Yahoo Finance API
- **Transaction processing** from Excel booking files
- **Comprehensive reporting** with 30+ configurable Excel exports
- **Time series analysis** with daily, monthly, and yearly aggregations
- **Intelligent yield calculations** with accurate handling of large transactions
- **Production deployment** via Windows Task Scheduler

## Recent Updates

### December 2025 - Yearly Profitability Analysis
- New profitability analysis table showing annual yields with holding period transparency
- Displays days held and annual yield percentage for each instrument per year
- Cash excluded from profitability calculations to maintain data quality
- Custom Excel export with dynamic formatting (scales automatically with portfolio growth)
- Time-Weighted Return (TWR) formula for accurate multi-period performance measurement

### December 2025 - Intelligent Yield Calculation
- Implemented intelligent denominator selection for yield calculations
- Prevents extreme yield percentages (100%+) on first trading days and depot transfers
- Handles large transactions (>50% of portfolio) with transaction-based denominators
- Improved accuracy of total return analysis with proper component breakdown

## Dependencies

Install required packages:

```bash
pip install -r requirements.txt
pip install ahlib
```

## Key Files

- `depot.py` - Main application with 8-step data processing pipeline
- `depot.ini` - Configuration file controlling all input/output settings
- `prices.parquet` - Local price data cache
- `start_depot.bat` - Production batch file for scheduled execution
- `requirements.txt` - Python dependencies
- `CLAUDE.md` - Technical documentation for developers

## Usage

### Development
```bash
python depot.py
```

### Production (Scheduled)
```bash
start_depot.bat
```

## Architecture

The application follows an 8-step data processing pipeline:
1. Configuration loading
2. Financial instruments import
3. Price data processing
4. Transaction bookings processing
5. Share position calculations
6. Portfolio value calculations
7. Provisions and fees processing
8. Excel report generation

Data is processed using pandas MultiIndex DataFrames with time series operations for financial analysis.

## Data Quality

For accurate yield calculations, ensure:
- `transaction_value_at_price` is filled for all buy/sell transactions
- Formula: `transaction_value_at_price = -(delta × price)`
- Depot transfers have balanced bookings (sum to zero)
- Interest/dividends use `interest_dividends` column, not transaction_value
