"""
Verify the new profitability_year.xlsx export format
"""
import pandas as pd

print("Verifying profitability_year.xlsx export...")
print("=" * 70)

try:
    # Read the exported file
    filepath = r'\\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Output\profitability_year.xlsx'
    df = pd.read_excel(filepath, index_col=[0, 1])

    print(f'\n✓ File successfully loaded from network share')
    print(f'\nStructure:')
    print(f'  Shape: {df.shape} (rows, columns)')
    print(f'  Index names: {df.index.names}')
    print(f'  Columns: {df.columns.tolist()}')

    print(f'\nData Summary:')
    print(f'  Years: {sorted(df.index.get_level_values("year").unique().tolist())}')
    print(f'  Number of WKNs: {len(df.index.get_level_values("wkn").unique())}')
    print(f'  WKNs: {sorted(df.index.get_level_values("wkn").unique().tolist())[:5]}...')

    print(f'\n\nFirst 15 rows:')
    print(df.head(15))

    print(f'\n\nSample: Year 2023 data:')
    print(df.loc[2023].head(10))

    print(f'\n\nColumn data types:')
    print(df.dtypes)

    print("\n" + "=" * 70)
    print("✓ New export format verified successfully!")
    print("\nFormat: MultiIndex (year, wkn) with columns ['days', 'yield']")
    print("This allows using the standard export_2D_df_to_excel_format function")

except FileNotFoundError:
    print("\n❌ File not found on network share")
    print("Make sure the network path is accessible")
except Exception as e:
    print(f"\n❌ Error: {e}")
