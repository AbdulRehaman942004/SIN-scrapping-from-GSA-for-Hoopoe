# Update Links for Missing Rows

## Overview

This folder contains a script to update GSA Advantage links for products that were not found during the initial scraping process. When products cannot be found using the primary search method (Item Stock Number-Butted), this script updates the links to use the secondary search method (Item Number) so they can be re-scraped.

## Purpose

After running the main scraping script in the `3 Scrapping` folder, some products may not be found because:
- The Item Stock Number-Butted in the link doesn't match the product on GSA Advantage
- The product might be searchable using the Item Number instead
- The link format needs to be updated to use a different identifier

This script identifies rows where all GSA data is missing (all 9 columns empty) and updates their links to use the Item Number from Column A instead of the Item Stock Number-Butted.

## What This Script Does

1. **Identifies Missing Rows**: Finds all rows where all 9 GSA-related columns are completely empty:
   - `GSA PRICE`, `Contractor`, `contract#:`
   - `GSA PRICE.1`, `Contractor.1`, `contract#:.1`
   - `GSA PRICE.2`, `Contractor.2`, `contract#:.2`

2. **Updates Links**: For each missing row, the script:
   - Extracts the current link from the "Links" column
   - Gets the "Item Number" from Column A
   - Replaces the Item Stock Number-Butted in the link with the Item Number
   - Adds "1" before the Item Number in the link (format: `q=7:1ITEM_NUMBER`)

3. **Creates Backup**: Automatically creates a timestamped backup of the Excel file before making any changes

4. **Saves Updates**: Saves the updated Excel file with the new links

## Link Format

### Before (Primary Search - Item Stock Number-Butted):
```
https://www.gsaadvantage.gov/advantage/ws/search/advantage_search?searchType=1&q=7:7029005&s=7&c=100
```

### After (Secondary Search - Item Number):
```
https://www.gsaadvantage.gov/advantage/ws/search/advantage_search?searchType=1&q=7:1AAG7029005&s=7&c=100
```

**Note**: The "1" prefix is added before the Item Number to indicate a different search method.

## How to Use

### Prerequisites

1. Ensure you have Python installed with the required packages:
   ```bash
   pip install pandas openpyxl
   ```

2. Make sure the Excel file `essendant-product-list_with_gsa_scraped_data.xlsx` exists in the `3 Scrapping` folder

### Running the Script

1. Navigate to this folder:
   ```bash
   cd "4 Update links for missing rows"
   ```

2. Run the script:
   ```bash
   python update_links_for_missing_rows.py
   ```

3. The script will:
   - Read the Excel file from `../3 Scrapping/essendant-product-list_with_gsa_scraped_data.xlsx`
   - Identify rows with missing GSA data
   - Update the links for those rows
   - Create a backup file
   - Save the updated Excel file

### Output

The script will display:
- Total number of rows with missing GSA data
- First 10 link updates (old vs new format)
- Summary of how many links were updated
- Location of the backup file created

## Next Steps

After running this script:

1. **Verify the Updates**: Check the Excel file to ensure links were updated correctly

2. **Re-run Scraping**: Go back to the `3 Scrapping` folder and run the main scraping script:
   ```bash
   cd "../3 Scrapping"
   python gsa_scraping_automation.py
   ```

3. **Use Option 5**: When prompted, select **Option 5: Scrape Missing Rows Only** to scrape only the rows that were updated with new links

4. **Monitor Progress**: The scraping script will now use the updated links (with Item Number) to search for products that weren't found previously

## Important Notes

- **Backup Created**: A backup file is automatically created before any changes are made. The backup filename format is: `essendant-product-list_with_gsa_scraped_data.backup_YYYYMMDD_HHMMSS.xlsx`

- **Only Missing Rows**: The script only updates links for rows where ALL 9 GSA columns are empty. Rows with partial data are not modified.

- **Link Format**: The script adds "1" before the Item Number in the link (e.g., `q=7:1AAG7029005`). This is intentional and required for the GSA Advantage search to work correctly with Item Numbers.

- **No Data Loss**: The script only modifies the "Links" column. All other data remains unchanged.

## Troubleshooting

### Error: Excel file not found
- Ensure the Excel file `essendant-product-list_with_gsa_scraped_data.xlsx` exists in the `3 Scrapping` folder
- Check that the folder structure is correct: `3 Scrapping/` should be a sibling folder to `4 Update links for missing rows/`

### Error: Permission denied
- Close the Excel file if it's open in Microsoft Excel
- Ensure you have write permissions to the Excel file

### No links updated
- This means no rows were found with all 9 GSA columns empty
- All products may already have been scraped successfully
- Or the missing rows don't have valid Item Numbers in Column A

## Workflow Summary

```
1. Run main scraping script (3 Scrapping/gsa_scraping_automation.py)
   ↓
2. Some products not found (missing GSA data)
   ↓
3. Run this script (4 Update links for missing rows/update_links_for_missing_rows.py)
   ↓
4. Links updated to use Item Number instead of Item Stock Number-Butted
   ↓
5. Re-run main scraping script with Option 5 (Missing Rows Only)
   ↓
6. Products found using updated links
```

## Technical Details

- **Script Language**: Python 3
- **Dependencies**: pandas, openpyxl
- **Excel File Location**: `../3 Scrapping/essendant-product-list_with_gsa_scraped_data.xlsx`
- **Backup Location**: Same folder as the Excel file
- **Link Pattern**: Uses regex `(q=7:)([^&]+)` to find and replace the item identifier

## Support

If you encounter any issues:
1. Check that all prerequisites are installed
2. Verify the Excel file structure matches the expected format
3. Ensure the folder structure is correct
4. Review the error messages for specific guidance

