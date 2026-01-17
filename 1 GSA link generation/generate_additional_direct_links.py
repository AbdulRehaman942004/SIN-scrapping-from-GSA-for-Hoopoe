import pandas as pd
import re
import os
from datetime import datetime
import shutil
from urllib.parse import urlencode

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
    """
    if pd.isna(link) or not link:
        return None
    
    # Pattern to match q=7:1{ALPHANUMERIC}
    pattern = r'q=7:1([^&]+)'
    match = re.search(pattern, str(link))
    
    if match:
        return match.group(1).strip()
    
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
    
    # Check if contract_number is empty string or 'nan'
    contract_str = str(contract_number).strip()
    if contract_str == '' or contract_str.lower() == 'nan':
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

def add_additional_direct_links():
    """Add GSA Direct Product Link 1 and GSA Direct Product Link 2 columns"""
    
    print("="*80)
    print("GENERATE ADDITIONAL GSA DIRECT PRODUCT LINKS")
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
    except Exception as e:
        print(f"[ERROR] Failed to read file: {str(e)}")
        return False
    
    # Verify required columns exist
    print("\nVerifying required columns...")
    
    required_columns = {
        'Links': None,
        'Manufacturer Long Name': None,
        'contract#:.1': None,
        'contract#:.2': None
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
    
    # Check if columns already exist
    new_columns = {
        "GSA Direct Product Link 1": "contract#:.1",
        "GSA Direct Product Link 2": "contract#:.2"
    }
    
    existing_new_columns = [col for col in new_columns.keys() if col in df.columns]
    if existing_new_columns:
        print(f"\n[WARNING] These columns already exist and will be OVERWRITTEN:")
        for col in existing_new_columns:
            print(f"   - {col}")
        df = df.drop(columns=existing_new_columns)
        print(f"[OK] Existing columns removed, regenerating...")
    
    # Process each row and generate direct links
    print(f"\nGenerating additional direct product links...")
    
    # For Link 1 (using contract#:.1)
    print(f"\n1. Generating 'GSA Direct Product Link 1' (using contract#:.1)...")
    direct_links_1 = []
    successful_1 = 0
    
    for idx, row in df.iterrows():
        # Extract item number from Links column
        link = row['Links']
        item_number = extract_item_number_from_link(link)
        
        # Get manufacturer and contract number
        manufacturer_name = row['Manufacturer Long Name']
        contract_number = row['contract#:.1']
        
        # Generate direct link
        direct_link = generate_direct_product_link(item_number, manufacturer_name, contract_number)
        
        direct_links_1.append(direct_link)
        
        if direct_link:
            successful_1 += 1
            if idx < 3:  # Show first 3 examples
                print(f"   Row {idx+1}: {direct_link[:80]}...")
    
    print(f"   [OK] Links generated: {successful_1} out of {len(df)}")
    
    # For Link 2 (using contract#:.2)
    print(f"\n2. Generating 'GSA Direct Product Link 2' (using contract#:.2)...")
    direct_links_2 = []
    successful_2 = 0
    
    for idx, row in df.iterrows():
        # Extract item number from Links column
        link = row['Links']
        item_number = extract_item_number_from_link(link)
        
        # Get manufacturer and contract number
        manufacturer_name = row['Manufacturer Long Name']
        contract_number = row['contract#:.2']
        
        # Generate direct link
        direct_link = generate_direct_product_link(item_number, manufacturer_name, contract_number)
        
        direct_links_2.append(direct_link)
        
        if direct_link:
            successful_2 += 1
            if idx < 3:  # Show first 3 examples
                print(f"   Row {idx+1}: {direct_link[:80]}...")
    
    print(f"   [OK] Links generated: {successful_2} out of {len(df)}")
    
    # Add the new columns at the end
    df["GSA Direct Product Link 1"] = direct_links_1
    df["GSA Direct Product Link 2"] = direct_links_2
    
    print(f"\n[OK] All direct links generated!")
    print(f"   GSA Direct Product Link 1: {successful_1} links")
    print(f"   GSA Direct Product Link 2: {successful_2} links")
    
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
        print(f"   New columns added: 'GSA Direct Product Link 1', 'GSA Direct Product Link 2'")
    except Exception as e:
        print(f"[ERROR] Failed to save file: {str(e)}")
        print(f"\n[INFO] Please close the Excel file and run the script again.")
        return False
    
    print("\n" + "="*80)
    print("[SUCCESS] ADDITIONAL GSA DIRECT PRODUCT LINKS GENERATED!")
    print("="*80)
    print(f"Updated file: {excel_file}")
    print(f"Summary:")
    print(f"   - GSA Direct Product Link 1: {successful_1} links created")
    print(f"   - GSA Direct Product Link 2: {successful_2} links created")
    print("="*80)
    
    return True

if __name__ == "__main__":
    try:
        success = add_additional_direct_links()
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
