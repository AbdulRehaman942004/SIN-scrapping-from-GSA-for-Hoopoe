import pandas as pd
import re
import os
from datetime import datetime
import shutil
from urllib.parse import urlencode, quote

def create_backup(file_path):
    """Create a timestamped backup of the file"""
    try:
        if os.path.exists(file_path):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = f"{file_path}.backup_{timestamp}"
            shutil.copy2(file_path, backup_path)
            print(f"[OK] Backup created: {backup_path}")
            return backup_path
        return None
    except Exception as e:
        print(f"[WARNING] Could not create backup: {str(e)}")
        return None

def extract_item_number_from_link(link):
    """
    Extract item number from GSA search link.
    Example: 'q=7:133041' -> '33041'
    Example: 'q=7:1AVE00166PK' -> 'AVE00166PK'
    Pattern: q=7:1{ITEM_NUMBER}
    """
    if pd.isna(link) or not link:
        return None
    
    # Pattern to match q=7:1{ALPHANUMERIC} - captures everything until & or end
    # Changed from \d+ (digits only) to [^&]+ (anything except &)
    pattern = r'q=7:1([^&]+)'
    match = re.search(pattern, str(link))
    
    if match:
        return match.group(1).strip()  # Return alphanumeric item number
    
    return None

def generate_direct_product_link(item_number, manufacturer_name, contract_number):
    """
    Generate direct GSA product detail link.
    Format: https://www.gsaadvantage.gov/advantage/ws/catalog/product_detail?itemNumber=XXX&mfrName=YYY&contractNumber=ZZZ
    """
    # If any required field is missing, return empty string
    if pd.isna(item_number) or not item_number:
        return ""
    if pd.isna(manufacturer_name) or not manufacturer_name:
        return ""
    if pd.isna(contract_number) or not contract_number:
        return ""
    
    # Base URL
    base_url = "https://www.gsaadvantage.gov/advantage/ws/catalog/product_detail"
    
    # Build parameters
    params = {
        'itemNumber': str(item_number),
        'mfrName': str(manufacturer_name),
        'contractNumber': str(contract_number)
    }
    
    # URL encode and build full URL
    query_string = urlencode(params)
    direct_link = f"{base_url}?{query_string}"
    
    return direct_link

def add_direct_links_column():
    """Add GSA Direct Product Link column to ScrappedProducts.xlsx"""
    
    print("="*80)
    print("GENERATE GSA DIRECT PRODUCT LINKS")
    print("="*80)
    
    excel_file = "ScrappedProducts.xlsx"
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"[ERROR] File not found: {excel_file}")
        return False
    
    print(f"\nReading file: {excel_file}")
    
    # Read Excel file
    try:
        df = pd.read_excel(excel_file)
        print(f"[OK] File loaded: {len(df)} rows, {len(df.columns)} columns")
        print(f"   Columns: {list(df.columns)}")
    except Exception as e:
        print(f"[ERROR] Failed to read file: {str(e)}")
        return False
    
    # Verify required columns exist
    print("\nVerifying required columns...")
    
    required_columns = {
        'Links': None,
        'Manufacturer Long Name': None,
        'contract#:': None
    }
    
    missing_columns = []
    for col in required_columns.keys():
        if col in df.columns:
            print(f"[OK] Found column: '{col}'")
        else:
            print(f"[ERROR] Missing column: '{col}'")
            missing_columns.append(col)
    
    if missing_columns:
        print(f"\n[ERROR] Cannot proceed. Missing required columns: {missing_columns}")
        return False
    
    # Check if column already exists
    new_column_name = "GSA Direct Product Link"
    if new_column_name in df.columns:
        print(f"\n[WARNING] Column '{new_column_name}' already exists and will be OVERWRITTEN")
        df = df.drop(columns=[new_column_name])
        print(f"[OK] Existing column removed, regenerating...")
    
    # Process each row and generate direct links
    print(f"\nGenerating direct product links...")
    
    direct_links = []
    successful = 0
    failed = 0
    
    for idx, row in df.iterrows():
        # Extract item number from Links column
        link = row['Links']
        item_number = extract_item_number_from_link(link)
        
        # Get manufacturer and contract number
        manufacturer_name = row['Manufacturer Long Name']
        contract_number = row['contract#:']
        
        # Generate direct link
        direct_link = generate_direct_product_link(item_number, manufacturer_name, contract_number)
        
        direct_links.append(direct_link)
        
        if direct_link:
            successful += 1
            if idx < 5:  # Show first 5 examples
                print(f"   Row {idx+1}: {direct_link[:80]}...")
        else:
            failed += 1
    
    if successful >= 5:
        print(f"   ... (showing first 5 examples)")
    
    # Add the new column at the end
    df[new_column_name] = direct_links
    
    print(f"\n[OK] Direct links generated!")
    print(f"   Successful: {successful}")
    print(f"   Failed (missing data): {failed}")
    
    # Create backup
    print(f"\nCreating backup...")
    create_backup(excel_file)
    
    # Save file
    print(f"\nSaving file: {excel_file}")
    try:
        df.to_excel(excel_file, index=False)
        print(f"[OK] File saved successfully!")
        print(f"   Total rows: {len(df)}")
        print(f"   Total columns: {len(df.columns)}")
        print(f"   New column added: '{new_column_name}'")
    except Exception as e:
        print(f"[ERROR] Failed to save file: {str(e)}")
        return False
    
    print("\n" + "="*80)
    print("[SUCCESS] GSA DIRECT PRODUCT LINKS GENERATED!")
    print("="*80)
    print(f"Updated file: {excel_file}")
    print(f"New column: '{new_column_name}'")
    print(f"Links created: {successful} out of {len(df)}")
    print("="*80)
    
    # Show example of extraction
    print("\nEXAMPLE - Item Number Extraction:")
    print("="*80)
    sample_link = "https://www.gsaadvantage.gov/advantage/ws/search/advantage_search?searchType=1&q=7:133041&s=7&c=100"
    sample_item = extract_item_number_from_link(sample_link)
    print(f"Original Link: {sample_link}")
    print(f"Extracted Item Number: {sample_item}")
    
    if len(df) > 0 and df['Links'].notna().any():
        first_valid_idx = df['Links'].first_valid_index()
        if first_valid_idx is not None:
            actual_link = df.at[first_valid_idx, 'Links']
            actual_item = extract_item_number_from_link(actual_link)
            print(f"\nActual from your data:")
            print(f"Original Link: {actual_link}")
            print(f"Extracted Item Number: {actual_item}")
    print("="*80)
    
    return True

if __name__ == "__main__":
    try:
        success = add_direct_links_column()
        if not success:
            print("\n[ERROR] Operation failed!")
            exit(1)
    except KeyboardInterrupt:
        print("\n\n[CANCELLED] Operation cancelled by user")
        exit(1)
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        exit(1)
