import pandas as pd
import re
import os
from datetime import datetime

# Import the identify_missing_rows function from the main script
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# We'll need to create a minimal version of the class to use identify_missing_rows
# Or we can just copy the logic

def identify_missing_rows(df):
    """Identify rows where GSA data is missing or incomplete - same logic as in gsa_scraping_automation.py"""
    missing_rows = []
    
    # Define all 9 GSA columns to check
    gsa_columns = [
        'GSA PRICE', 'Contractor', 'contract#:',
        'GSA PRICE.1', 'Contractor.1', 'contract#:.1',
        'GSA PRICE.2', 'Contractor.2', 'contract#:.2'
    ]
    
    # Check for rows where all 9 GSA columns are empty
    for i, row in df.iterrows():
        # Check if all 9 columns are empty
        all_empty = True
        for col in gsa_columns:
            value = row.get(col, '')
            # Check if value is NaN, empty string, or 'nan' string
            if not (pd.isna(value) or str(value).strip() == '' or str(value).strip().lower() == 'nan'):
                all_empty = False
                break
        
        # Consider a row missing if all 9 columns are empty
        if all_empty:
            missing_rows.append(i)
    
    return missing_rows

def update_links_for_missing_rows(excel_file):
    """Update links for rows with missing GSA data"""
    
    print(f"Reading Excel file: {excel_file}")
    df = pd.read_excel(excel_file)
    
    print(f"Total rows in file: {len(df)}")
    
    # Find the required columns
    item_number_col = None
    item_stock_number_col = None
    links_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if col_lower == 'item number':
            item_number_col = col
        elif 'item stock number' in col_lower and 'butted' in col_lower:
            item_stock_number_col = col
        elif 'links' in col_lower:
            links_col = col
    
    if not item_number_col:
        print("ERROR: 'Item Number' column not found!")
        return False
    
    if not links_col:
        print("ERROR: 'Links' column not found!")
        return False
    
    print(f"\nFound columns:")
    print(f"  - Item Number: {item_number_col}")
    print(f"  - Item Stock Number-Butted: {item_stock_number_col}")
    print(f"  - Links: {links_col}")
    
    # Identify rows with missing GSA data
    print(f"\nIdentifying rows with missing GSA data...")
    missing_rows = identify_missing_rows(df)
    
    if not missing_rows:
        print("\nNo missing rows found! All products have GSA data.")
        return True
    
    print(f"\nFound {len(missing_rows)} rows with missing GSA data")
    
    # Pattern to extract and replace the item number in the link
    # Link format: https://www.gsaadvantage.gov/advantage/ws/search/advantage_search?searchType=1&q=7:ITEM_NUMBER&s=7&c=100
    link_pattern = r'(q=7:)([^&]+)'
    
    updated_count = 0
    skipped_count = 0
    
    print(f"\nUpdating links for {len(missing_rows)} rows...")
    
    for idx in missing_rows:
        try:
            # Get current link
            current_link = df.at[idx, links_col]
            
            # Skip if link is empty or NaN
            if pd.isna(current_link) or not str(current_link).strip():
                print(f"  Row {idx+1}: Skipping - No link found")
                skipped_count += 1
                continue
            
            current_link = str(current_link).strip()
            
            # Get Item Number from column A
            item_number = df.at[idx, item_number_col]
            
            # Skip if Item Number is empty
            if pd.isna(item_number) or not str(item_number).strip():
                print(f"  Row {idx+1}: Skipping - No Item Number found")
                skipped_count += 1
                continue
            
            item_number = str(item_number).strip()
            
            # Replace the item number in the link
            # Pattern: q=7:OLD_VALUE -> q=7:1NEW_VALUE (add ":1" before Item Number)
            def replace_item_number(match):
                return f"{match.group(1)}1{item_number}"
            
            new_link = re.sub(link_pattern, replace_item_number, current_link)
            
            # Update the link in dataframe
            df.at[idx, links_col] = new_link
            updated_count += 1
            
            if updated_count <= 10:  # Show first 10 updates
                print(f"  Row {idx+1}: Updated link")
                print(f"    Old: {current_link}")
                print(f"    New: {new_link}")
            
        except Exception as e:
            print(f"  Row {idx+1}: Error - {str(e)}")
            skipped_count += 1
            continue
    
    if updated_count > 10:
        print(f"  ... and {updated_count - 10} more links updated")
    
    print(f"\n{'='*60}")
    print(f"UPDATE SUMMARY:")
    print(f"{'='*60}")
    print(f"Total missing rows: {len(missing_rows)}")
    print(f"Links updated: {updated_count}")
    print(f"Skipped: {skipped_count}")
    
    if updated_count > 0:
        # Create backup before saving
        backup_file = excel_file.replace('.xlsx', f'.backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
        print(f"\nCreating backup: {backup_file}")
        df_backup = pd.read_excel(excel_file)  # Read original for backup
        df_backup.to_excel(backup_file, index=False)
        print(f"Backup created successfully")
        
        # Save updated file
        print(f"\nSaving updated Excel file...")
        df.to_excel(excel_file, index=False)
        print(f"File saved successfully: {excel_file}")
        print(f"\n{'='*60}")
        print(f"SUCCESS! Updated {updated_count} links in the Excel file.")
        print(f"{'='*60}")
    else:
        print(f"\nNo links were updated.")
    
    return True

if __name__ == "__main__":
    # Get the script directory and construct path to Excel file in "3 Scrapping" folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    excel_file = os.path.join(parent_dir, "3 Scrapping", "essendant-product-list_with_gsa_scraped_data.xlsx")
    
    if not os.path.exists(excel_file):
        print(f"ERROR: Excel file not found: {excel_file}")
        print(f"Please ensure the Excel file exists in the '3 Scrapping' folder.")
        exit(1)
    
    print("="*60)
    print("UPDATE LINKS FOR MISSING ROWS")
    print("="*60)
    print("This script will:")
    print("1. Identify rows with missing GSA data (all 9 columns empty)")
    print("2. Update links by replacing 'Item Stock Number-Butted' with 'Item Number'")
    print("="*60)
    
    success = update_links_for_missing_rows(excel_file)
    
    if success:
        print("\nProcess completed successfully!")
    else:
        print("\nProcess failed!")

