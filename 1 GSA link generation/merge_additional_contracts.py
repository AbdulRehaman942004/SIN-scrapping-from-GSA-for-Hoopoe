import pandas as pd
import os
from datetime import datetime
import shutil
import sys

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

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

def merge_additional_contracts():
    """Merge contract#:.1 and contract#:.2 columns from scraped data to ScrappedProducts"""
    
    print("="*80)
    print("MERGE ADDITIONAL CONTRACT COLUMNS")
    print("="*80)
    
    # File paths
    source_file = "3 Scrapping/essendant-product-list_with_gsa_scraped_data.xlsx"
    destination_file = "ScrappedProducts.xlsx"
    
    # Check if files exist
    if not os.path.exists(source_file):
        print(f"[ERROR] Source file not found: {source_file}")
        return False
    
    if not os.path.exists(destination_file):
        print(f"[ERROR] Destination file not found: {destination_file}")
        return False
    
    print(f"\nReading source file: {source_file}")
    # Read source file (scraped data)
    try:
        df_source = pd.read_excel(source_file)
        print(f"[OK] Source file loaded: {len(df_source)} rows")
        print(f"   Columns: {list(df_source.columns)[:5]}...")
    except Exception as e:
        print(f"[ERROR] reading source file: {str(e)}")
        return False
    
    print(f"\nReading destination file: {destination_file}")
    # Read destination file (ScrappedProducts)
    try:
        df_destination = pd.read_excel(destination_file)
        print(f"[OK] Destination file loaded: {len(df_destination)} rows")
        print(f"   Columns: {list(df_destination.columns)[:5]}...")
    except Exception as e:
        print(f"[ERROR] reading destination file: {str(e)}")
        return False
    
    # Verify required columns exist
    print("\nVerifying columns...")
    
    # Check Item Number in both files
    if "Item Number" not in df_source.columns:
        print(f"[ERROR] 'Item Number' column not found in source file")
        print(f"   Available columns: {list(df_source.columns)}")
        return False
    
    if "Item Number" not in df_destination.columns:
        print(f"[ERROR] 'Item Number' column not found in destination file")
        print(f"   Available columns: {list(df_destination.columns)}")
        return False
    
    print("[OK] 'Item Number' column found in both files")
    
    # Check for columns to merge
    columns_to_merge = []
    
    if "contract#:.1" in df_source.columns:
        columns_to_merge.append("contract#:.1")
        print("[OK] 'contract#:.1' column found in source file")
    else:
        print("[WARNING] 'contract#:.1' column not found in source file")
        print(f"   Available columns: {list(df_source.columns)}")
    
    if "contract#:.2" in df_source.columns:
        columns_to_merge.append("contract#:.2")
        print("[OK] 'contract#:.2' column found in source file")
    else:
        print("[WARNING] 'contract#:.2' column not found in source file")
        print(f"   Available columns: {list(df_source.columns)}")
    
    if not columns_to_merge:
        print("[ERROR] No columns to merge found!")
        return False
    
    print(f"\nColumns to merge: {columns_to_merge}")
    
    # Prepare source data for merging (only Item Number + columns to merge)
    merge_columns_list = ["Item Number"] + columns_to_merge
    df_merge = df_source[merge_columns_list].copy()
    
    # Remove duplicates in source data (keep first occurrence)
    original_count = len(df_merge)
    df_merge = df_merge.drop_duplicates(subset=["Item Number"], keep="first")
    if len(df_merge) < original_count:
        print(f"[WARNING] Removed {original_count - len(df_merge)} duplicate Item Numbers from source")
    
    print(f"\nMerging data...")
    print(f"   Destination rows before merge: {len(df_destination)}")
    
    # Check if columns already exist in destination
    existing_columns = [col for col in columns_to_merge if col in df_destination.columns]
    if existing_columns:
        print(f"[WARNING] These columns already exist in destination and will be OVERWRITTEN:")
        for col in existing_columns:
            print(f"      - {col}")
        
        # Drop existing columns
        df_destination = df_destination.drop(columns=existing_columns)
        print(f"[OK] Existing columns removed")
    
    # Perform left merge (keep all rows from destination)
    df_result = df_destination.merge(
        df_merge,
        on="Item Number",
        how="left"
    )
    
    print(f"[OK] Merge completed!")
    print(f"   Rows after merge: {len(df_result)}")
    
    # Show statistics
    print(f"\nMerge Statistics:")
    for col in columns_to_merge:
        non_null_count = df_result[col].notna().sum()
        null_count = df_result[col].isna().sum()
        # Count non-empty strings (exclude 'nan' strings)
        non_empty_count = ((df_result[col].notna()) & 
                          (df_result[col].astype(str).str.strip() != '') & 
                          (df_result[col].astype(str).str.strip().str.lower() != 'nan')).sum()
        print(f"   {col}:")
        print(f"      - Non-empty values: {non_empty_count}")
        print(f"      - Empty/Missing: {len(df_result) - non_empty_count}")
    
    # Create backup before saving
    print(f"\nCreating backup of destination file...")
    create_backup(destination_file)
    
    # Save result
    print(f"\nSaving merged data to: {destination_file}")
    try:
        df_result.to_excel(destination_file, index=False)
        print(f"[OK] File saved successfully!")
        print(f"   Total rows: {len(df_result)}")
        print(f"   Total columns: {len(df_result.columns)}")
    except Exception as e:
        print(f"[ERROR] saving file: {str(e)}")
        return False
    
    print("\n" + "="*80)
    print("[SUCCESS] ADDITIONAL CONTRACT COLUMNS MERGED!")
    print("="*80)
    print(f"Updated file: {destination_file}")
    print(f"New columns added: {', '.join(columns_to_merge)}")
    print("="*80)
    
    return True

if __name__ == "__main__":
    try:
        success = merge_additional_contracts()
        if not success:
            print("\n[ERROR] Merge operation failed!")
            exit(1)
    except KeyboardInterrupt:
        print("\n\n[CANCELLED] Operation cancelled by user")
        exit(1)
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        exit(1)
