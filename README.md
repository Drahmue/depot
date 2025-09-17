# Depot Project

A comprehensive financial portfolio analysis system that processes stock prices, bookings, and generates detailed financial reports. The system integrates with Yahoo Finance for price data and produces 30+ different Excel reports for portfolio analysis.

## Features

- **Multi-bank portfolio tracking** with position and value calculations
- **Automated price data updates** using Yahoo Finance API
- **Transaction processing** from Excel booking files
- **Comprehensive reporting** with 30+ configurable Excel exports
- **Time series analysis** with daily, monthly, and yearly aggregations
- **Production deployment** via Windows Task Scheduler

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