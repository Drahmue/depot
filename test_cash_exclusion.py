"""
Quick test to verify cash is excluded from profitability calculations
"""
import pandas as pd
import sys

print("Testing cash exclusion from profitability calculations...")
print("=" * 70)

# Test yield components
try:
    # Check if yield component files exist and what WKNs they contain
    test_files = [
        'TEST_yield_price_day.xlsx',
        'TEST_yield_total_day.xlsx',
    ]

    for filename in test_files:
        try:
            # Try to read from network path
            filepath = f'\\\\WIN-H7BKO5H0RMC\\Dataserver\\Dummy\\Finance_Output\\{filename}'
            df = pd.read_excel(filepath, index_col=[0, 1])

            wkns = df.index.get_level_values('wkn').unique()
            wkns_lower = [str(w).lower() for w in wkns]

            print(f"\n{filename}:")
            print(f"  Total WKNs: {len(wkns)}")
            print(f"  Contains 'cash': {'cash' in wkns_lower}")
            print(f"  Contains 'cm': {'cm' in wkns_lower}")
            print(f"  Contains 'ftd': {'ftd' in wkns_lower}")

            if 'cash' in wkns_lower:
                print("  ❌ FAILED: Cash should be excluded!")
            else:
                print("  ✓ PASSED: Cash is excluded")

        except FileNotFoundError:
            print(f"\n{filename}: File not found (skipping)")
        except Exception as e:
            print(f"\n{filename}: Error - {e}")

    # Check profitability_year table
    try:
        filepath = '\\\\WIN-H7BKO5H0RMC\\Dataserver\\Dummy\\Finance_Output\\profitability_year.xlsx'
        df = pd.read_excel(filepath, index_col=0)

        columns = df.columns.tolist()
        cash_columns = [col for col in columns if 'cash' in col.lower()]
        cm_columns = [col for col in columns if col.startswith('cm_')]
        ftd_columns = [col for col in columns if col.startswith('ftd_')]

        print(f"\nprofitability_year.xlsx:")
        print(f"  Total columns: {len(columns)}")
        print(f"  Cash columns: {len(cash_columns)}")
        print(f"  CM columns: {len(cm_columns)}")
        print(f"  FTD columns: {len(ftd_columns)}")

        if cash_columns:
            print(f"  Cash columns found: {cash_columns}")
            print("  ❌ FAILED: Cash columns should not exist!")
        else:
            print("  ✓ PASSED: No cash columns")

        if cm_columns:
            print(f"  ✓ CM columns exist: {cm_columns[:2]}...")
        if ftd_columns:
            print(f"  ✓ FTD columns exist: {ftd_columns[:2]}...")

    except FileNotFoundError:
        print("\nprofitability_year.xlsx: File not found (skipping)")
    except Exception as e:
        print(f"\nprofitability_year.xlsx: Error - {e}")

    print("\n" + "=" * 70)
    print("Test complete!")

except Exception as e:
    print(f"Unexpected error: {e}")
    sys.exit(1)
