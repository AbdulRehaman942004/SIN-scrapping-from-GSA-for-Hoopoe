import pandas as pd

# Read Excel file
excel_file = "../../ScrappedProducts.xlsx"
df = pd.read_excel(excel_file)

print("="*80)
print("EXCEL FILE DIAGNOSIS")
print("="*80)
print(f"\nFile: {excel_file}")
print(f"Total rows: {len(df)}")
print(f"Total columns: {len(df.columns)}")

print(f"\nAll columns:")
for i, col in enumerate(df.columns, 1):
    print(f"  {i}. {col}")

print(f"\nSIN Column Check:")
print(f"  Has SIN1: {'SIN1' in df.columns}")
print(f"  Has SIN2: {'SIN2' in df.columns}")
print(f"  Has SIN3: {'SIN3' in df.columns}")

if 'SIN1' in df.columns:
    sin1_count = df['SIN1'].notna().sum()
    sin1_not_found = (df['SIN1'] == "SIN not found").sum()
    print(f"\nSIN1 Statistics:")
    print(f"  Non-empty cells: {sin1_count}")
    print(f"  'SIN not found' entries: {sin1_not_found}")
    print(f"  Empty cells: {len(df) - sin1_count}")
    print(f"\nFirst 10 SIN1 values:")
    for i, val in enumerate(df['SIN1'].head(10), 1):
        print(f"    Row {i}: {val}")

if 'SIN2' in df.columns:
    sin2_count = df['SIN2'].notna().sum()
    print(f"\nSIN2 Statistics:")
    print(f"  Non-empty cells: {sin2_count}")

if 'SIN3' in df.columns:
    sin3_count = df['SIN3'].notna().sum()
    print(f"\nSIN3 Statistics:")
    print(f"  Non-empty cells: {sin3_count}")

print("\n" + "="*80)
