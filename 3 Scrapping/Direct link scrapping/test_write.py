import pandas as pd
import os

excel_file = "../../ScrappedProducts.xlsx"

print("="*80)
print("WRITE TEST")
print("="*80)

# Check if file is writable
if os.path.exists(excel_file):
    print(f"\nFile exists: {excel_file}")
    print(f"File size: {os.path.getsize(excel_file)} bytes")
    
    # Try to open for writing
    try:
        with open(excel_file, 'r+b') as f:
            print("[OK] File is writable (not locked)")
    except PermissionError:
        print("[ERROR] File is LOCKED - Please close Excel/Cursor!")
        print("       The file is open in another program")
        exit(1)
    except Exception as e:
        print(f"[ERROR] Cannot access file: {str(e)}")
        exit(1)
    
    # Try to read and write
    try:
        df = pd.read_excel(excel_file)
        print(f"[OK] Can read file: {len(df)} rows")
        
        # Try to write (test with temp file)
        temp_file = excel_file + ".test_write"
        df.head(10).to_excel(temp_file, index=False)
        print(f"[OK] Can write files")
        
        # Clean up
        if os.path.exists(temp_file):
            os.remove(temp_file)
            print(f"[OK] Test cleanup successful")
        
        print("\n" + "="*80)
        print("[SUCCESS] No file locking issues detected")
        print("="*80)
        
    except Exception as e:
        print(f"[ERROR] Write test failed: {str(e)}")
        exit(1)
else:
    print(f"[ERROR] File not found: {excel_file}")
    exit(1)
