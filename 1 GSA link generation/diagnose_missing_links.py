import pandas as pd
import re

def extract_item_number_from_link(link):
    """Extract item number from GSA search link."""
    if pd.isna(link) or not link:
        return None
    
    pattern = r'q=7:1(\d+)'
    match = re.search(pattern, str(link))
    
    if match:
        return match.group(1)
    
    return None

def diagnose_data():
    """Diagnose why direct links are failing"""
    
    print("="*80)
    print("DIAGNOSTIC - Analyzing Failed Direct Links")
    print("="*80)
    
    excel_file = "ScrappedProducts.xlsx"
    
    print(f"\nReading file: {excel_file}")
    df = pd.read_excel(excel_file)
    print(f"Total rows: {len(df)}")
    
    # Check each required column
    print("\n" + "="*80)
    print("COLUMN ANALYSIS")
    print("="*80)
    
    # 1. Links column
    print("\n1. LINKS COLUMN:")
    links_empty = df['Links'].isna().sum()
    links_not_empty = df['Links'].notna().sum()
    print(f"   Empty: {links_empty}")
    print(f"   Not Empty: {links_not_empty}")
    
    # Test extraction
    extracted_items = df['Links'].apply(extract_item_number_from_link)
    extracted_success = extracted_items.notna().sum()
    extracted_failed = extracted_items.isna().sum()
    print(f"   Item Number Extracted Successfully: {extracted_success}")
    print(f"   Item Number Extraction Failed: {extracted_failed}")
    
    # Show some examples of failed extractions
    if extracted_failed > 0:
        print("\n   Examples of links where extraction FAILED:")
        failed_mask = extracted_items.isna() & df['Links'].notna()
        failed_links = df[failed_mask]['Links'].head(5)
        for i, link in enumerate(failed_links, 1):
            print(f"      {i}. {link}")
    
    # 2. Manufacturer Long Name column
    print("\n2. MANUFACTURER LONG NAME COLUMN:")
    mfr_empty = df['Manufacturer Long Name'].isna().sum()
    mfr_not_empty = df['Manufacturer Long Name'].notna().sum()
    print(f"   Empty: {mfr_empty}")
    print(f"   Not Empty: {mfr_not_empty}")
    
    # Check for empty strings (not NA but empty)
    mfr_empty_strings = (df['Manufacturer Long Name'].astype(str).str.strip() == '').sum()
    print(f"   Empty Strings (not NA): {mfr_empty_strings}")
    
    # 3. contract#: column
    print("\n3. CONTRACT#: COLUMN:")
    contract_empty = df['contract#:'].isna().sum()
    contract_not_empty = df['contract#:'].notna().sum()
    print(f"   Empty (NA): {contract_empty}")
    print(f"   Not Empty: {contract_not_empty}")
    
    # Check for empty strings
    contract_empty_strings = (df['contract#:'].astype(str).str.strip() == '').sum()
    contract_nan_strings = (df['contract#:'].astype(str).str.strip().str.lower() == 'nan').sum()
    print(f"   Empty Strings (not NA): {contract_empty_strings}")
    print(f"   'nan' Strings: {contract_nan_strings}")
    
    # Show some examples
    print("\n   Sample of contract#: values (first 10 non-empty):")
    sample_contracts = df[df['contract#:'].notna()]['contract#:'].head(10)
    for i, val in enumerate(sample_contracts, 1):
        print(f"      {i}. '{val}' (type: {type(val).__name__})")
    
    # 4. Combined Analysis - Why are links failing?
    print("\n" + "="*80)
    print("COMBINED FAILURE ANALYSIS")
    print("="*80)
    
    # Create conditions
    has_extracted_item = extracted_items.notna()
    has_manufacturer = df['Manufacturer Long Name'].notna() & (df['Manufacturer Long Name'].astype(str).str.strip() != '')
    has_contract = df['contract#:'].notna() & (df['contract#:'].astype(str).str.strip() != '') & (df['contract#:'].astype(str).str.strip().str.lower() != 'nan')
    
    # Count combinations
    all_three = has_extracted_item & has_manufacturer & has_contract
    print(f"\nRows with ALL THREE required fields: {all_three.sum()}")
    
    print(f"\nBreakdown of missing data:")
    print(f"   Missing Item Number (extraction failed): {(~has_extracted_item).sum()}")
    print(f"   Missing Manufacturer Name: {(~has_manufacturer).sum()}")
    print(f"   Missing Contract Number: {(~has_contract).sum()}")
    
    # More detailed breakdown
    print(f"\nDetailed combinations:")
    print(f"   Has Item + Manufacturer, Missing Contract: {(has_extracted_item & has_manufacturer & ~has_contract).sum()}")
    print(f"   Has Item + Contract, Missing Manufacturer: {(has_extracted_item & has_contract & ~has_manufacturer).sum()}")
    print(f"   Has Manufacturer + Contract, Missing Item: {(~has_extracted_item & has_manufacturer & has_contract).sum()}")
    print(f"   Missing All Three: {(~has_extracted_item & ~has_manufacturer & ~has_contract).sum()}")
    
    # Check GSA Direct Product Link column
    print("\n" + "="*80)
    print("CURRENT GSA DIRECT PRODUCT LINK COLUMN")
    print("="*80)
    
    if 'GSA Direct Product Link' in df.columns:
        direct_links = df['GSA Direct Product Link']
        links_generated = (direct_links.notna() & (direct_links.astype(str).str.strip() != '')).sum()
        links_missing = (direct_links.isna() | (direct_links.astype(str).str.strip() == '')).sum()
        print(f"   Links Generated: {links_generated}")
        print(f"   Links Missing: {links_missing}")
        
        # Show some examples
        print("\n   Sample of generated links (first 3):")
        valid_links = df[direct_links.notna() & (direct_links.astype(str).str.strip() != '')]['GSA Direct Product Link'].head(3)
        for i, link in enumerate(valid_links, 1):
            print(f"      {i}. {link[:100]}...")
    
    print("\n" + "="*80)

if __name__ == "__main__":
    diagnose_data()
