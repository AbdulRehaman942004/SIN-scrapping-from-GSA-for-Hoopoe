# GSA SIN Scraping from Direct Product Links

## Overview
This script scrapes SIN (Special Item Numbers) from GSA Advantage product detail pages using the direct product links that were previously generated.

## What It Does
- Reads `ScrappedProducts.xlsx` from the root folder
- **Checks if SIN already exists** - skips scraping if value present (Resume Mode)
- Visits each GSA Direct Product Link (up to 2 per product, max)
- Extracts SIN from "Schedule/SIN" field on the product page
- Extracts only the part after "/" (e.g., "MAS/332510C" → "332510C")
- Fills columns SIN1, SIN2, SIN3 with the scraped values
- **100% Safe to Run Multiple Times** - never overwrites existing data

## Columns Used

### Input Columns:
- `GSA Direct Product Link` → Scrapes to `SIN1`
- `GSA Direct Product Link 1` → Scrapes to `SIN2`
- `GSA Direct Product Link 2` → Scrapes to `SIN3`

### Output Columns:
- `SIN1` - SIN from first contract/link
- `SIN2` - SIN from second contract/link
- `SIN3` - SIN from third contract/link

## How to Run

### Prerequisites
1. Python 3.9+
2. Chrome Browser installed
3. Dependencies installed: `pip install -r requirements.txt`
4. `ScrappedProducts.xlsx` must exist in root folder with direct product link columns

### Running the Script

```bash
cd "3 Scrapping/Direct link scrapping"
python scrape_sin_from_direct_links.py
```

### Menu Options

1. **Test Mode (First 10 rows)**
   - Quick test with first 10 products
   - Recommended for first run to verify everything works

2. **Custom Range (Specify start and end)**
   - Scrape specific rows
   - Useful for resuming interrupted runs

3. **Full Automation (All rows)**
   - Scrapes all products
   - Estimated time: 10-15 hours for ~18,000 products
   - Progress saved every 10 rows

4. **Exit**

## Safe Shutdown (Ctrl+C Protection)

### The Problem

Pressing Ctrl+C while the script is writing to Excel can corrupt the entire file, losing all your data!

### The Solution

The script now has **graceful shutdown protection**:

1. **Signal Handler**: Catches Ctrl+C (SIGINT) signal
2. **Saves Data**: Completes current row and saves all data
3. **Clean Exit**: Closes browser and exits safely
4. **No Corruption**: File is never corrupted

### How It Works

```
User presses Ctrl+C
    ↓
[DETECTED] Shutdown signal received
    ↓
[WAITING] Completes current row
    ↓
[SAVING] Writes all data to disk
    ↓
[CLEANUP] Closes browser
    ↓
[EXIT] Safe shutdown complete
```

### What You See

```
[User presses Ctrl+C]

================================================================================
SHUTDOWN REQUESTED (Ctrl+C detected)
================================================================================
⚠️  Please wait... Saving data safely to prevent corruption!
⚠️  DO NOT force close or press Ctrl+C again!
================================================================================

================================================================================
GRACEFUL SHUTDOWN IN PROGRESS
================================================================================
Last completed row: 1250
Saving all data before exit...
================================================================================

[EMERGENCY SAVE] Saving data...
[EMERGENCY SAVE] Data saved successfully!
[CLEANUP] Closing browser...
[CLEANUP] Browser closed

================================================================================
SHUTDOWN COMPLETE - Data saved safely!
================================================================================
✓ Processed: 1250 rows
✓ Data saved to: ../../ScrappedProducts.xlsx
✓ Resume from row: 1251
================================================================================
```

### Additional Safety: Atomic File Writes

Even during normal saves, the script uses **atomic writes**:

1. Writes to temporary file: `ScrappedProducts.xlsx.tmp`
2. Only after complete, renames to: `ScrappedProducts.xlsx`
3. If interrupted during write, original file is untouched
4. Zero risk of partial write corruption

### Rules for Safe Shutdown

✅ **DO:**
- Press Ctrl+C once and wait
- Let it complete current row
- Let it save data
- Give it 10-20 seconds

❌ **DON'T:**
- Press Ctrl+C twice (force kill)
- Close terminal window immediately
- Kill the process forcefully
- Shut down computer without waiting

### Resume After Shutdown

After graceful shutdown:
```bash
# Script tells you where it stopped
✓ Resume from row: 1251

# Use Custom Range to continue
python scrape_sin_from_direct_links.py
> Choose: 2 (Custom Range)
> Start row: 1251
> End row: 18409
```

## Skip Entire Row (Efficiency Mode)

### The Rule

**If a row has at least 1 SIN (any of SIN1, SIN2, or SIN3), the entire row is skipped.**

### Why?

- If you already have some SIN data for a product, you don't need more
- Saves ~15 seconds per row (no browser requests needed)
- On re-runs with partial data, this can save hours!

### Example

```
Row 150 has:
  SIN1: 332510C (exists)
  SIN2: (empty)
  SIN3: (empty)

Result: ENTIRE ROW SKIPPED
Reason: Already has SIN1, no need for more data
Time Saved: ~15 seconds
```

### Terminal Output

```
================================================================================
ROW 150/18409 | Progress: 0.8%
================================================================================
[ROW SKIP] Row already has SIN data (SIN1: 332510C)
[ROW SKIP] Skipping entire row to save time
```

### Benefits

| Scenario | Old Behavior | New Behavior | Time Saved |
|----------|--------------|--------------|------------|
| **Has 1 SIN** | Tries to fill 2 more | Skips entire row | ~15 seconds |
| **1000 rows with partial data** | Wastes ~4 hours | Skips instantly | ~4 hours! |
| **Re-run after 90% complete** | Wastes ~9 hours | ~1 hour total | ~8 hours! |

## Resume Capability (Skip-if-Exists)

### How It Works

The script automatically checks if a SIN already exists before scraping:

1. **Check Row**: Before processing, checks if row has ANY SIN (SIN1, SIN2, or SIN3)
2. **Skip Entire Row**: If ANY SIN exists, skips entire row (no browser requests)
3. **Save Time**: Skipped rows save ~15 seconds each
4. **Only Scrape Empty**: Only scrapes rows with ALL SINs empty

### Example Terminal Output

**Row with existing data (skipped entirely):**
```
================================================================================
ROW 150/18409 | Progress: 0.8%
================================================================================
[ROW SKIP] Row already has SIN data (SIN2: 665210)
[ROW SKIP] Skipping entire row to save time
```

**Row with no data (scrapes normally):**
```
================================================================================
ROW 151/18409 | Progress: 0.8%
================================================================================
[1/3] Checking SIN1...
      URL: https://www.gsaadvantage.gov/...
      [SUCCESS] SIN1: 332510C
[2/3] Checking SIN2...
      URL: https://www.gsaadvantage.gov/...
      [SUCCESS] SIN2: 778899A
[EARLY STOP] Already have 2 SINs, skipping SIN3

ROW 151 SUMMARY:
  SINs Total: 2/2
  Scraped: 2
  Time: 14.5s
```

### Benefits

- ✅ **100% Safe to Run Multiple Times**: Never overwrites existing data
- ✅ **Resume Interrupted Runs**: Continue from where you left off
- ✅ **Skip Partial Data**: If row has ANY SIN, skips entire row
- ✅ **Huge Time Savings**: Skips thousands of unnecessary requests on re-runs
- ✅ **Faster Re-runs**: Rows with partial data skip instantly (~15s each)

## Scraping Strategy

The script uses multiple strategies to find SIN numbers:

1. **Text Pattern Search**: Searches entire page for "Schedule/SIN" patterns
2. **Table Search**: Looks for SIN in table rows/cells
3. **Element Search**: Uses XPath to find specific elements with Schedule/SIN

## Features

### Scraping Intelligence
- ✅ **Smart Limit**: Maximum 2 SINs per product (stops early if 2 found)
- ✅ **Skip-if-Exists**: Automatically skips SINs that already exist (Resume Mode)
- ✅ **Skip Entire Row**: If row has ANY SIN data, skips entire row (huge time saver)
- ✅ **Multi-Strategy Extraction**: 3 different methods to find SIN numbers
- ✅ **Retry Logic**: Automatic retry (up to 3 attempts) for network failures
- ✅ **Idempotent**: Safe to run multiple times without overwriting data

### Reliability & Stability
- ✅ **Driver Health Checks**: Monitors browser health
- ✅ **Automatic Recovery**: Restarts browser if it crashes
- ✅ **Periodic Restart**: Refreshes browser every 100 rows for long-running stability
- ✅ **Graceful Error Handling**: Continues processing even if individual pages fail
- ✅ **Safe Ctrl+C Shutdown**: Prevents file corruption when interrupting script
- ✅ **Atomic File Writes**: Uses temporary file + rename to prevent corruption

### Performance & Safety
- ✅ **Rate Limiting**: 2 seconds between requests (server-friendly)
- ✅ **Auto-save**: Progress saved every 50 rows (faster processing)
- ✅ **Smart Backups**: Keeps last 5 backups in dedicated `/backups` folder (auto-cleanup)
- ✅ **Early Stopping**: Skips unnecessary requests when 2 SINs found
- ✅ **Clean Structure**: Backups stored separately, no clutter in main folder

### Analytics & Monitoring
- ✅ **Real-time Progress**: Detailed row-by-row status
- ✅ **Session Statistics**: Success rate, averages, breakdowns
- ✅ **ETA Calculation**: Accurate time remaining estimates
- ✅ **Detailed Logging**: Full debug logs for troubleshooting

## Output

The script updates `ScrappedProducts.xlsx` with three new columns:
- `SIN1`: SIN from primary contract
- `SIN2`: SIN from alternative contract 1
- `SIN3`: SIN from alternative contract 2

### Backup Structure

All backups are stored in a dedicated `backups/` folder:

```
Extraction of SIN from GSA for Hoopoe Labs/
├── ScrappedProducts.xlsx (working file)
└── backups/
    ├── ScrappedProducts.xlsx.backup_20260117_040530
    ├── ScrappedProducts.xlsx.backup_20260117_040825
    ├── ScrappedProducts.xlsx.backup_20260117_041120
    ├── ScrappedProducts.xlsx.backup_20260117_041415
    └── ScrappedProducts.xlsx.backup_20260117_041710
    (keeps last 5 backups, auto-deletes older ones)
```

**Benefits:**
- ✅ Main folder stays clean
- ✅ All backups in one place
- ✅ Easy to find and restore
- ✅ Automatic cleanup (keeps last 5)

## Example

### Scenario 1: Resume mode - Some SINs already exist
**Excel Before:**
```
SIN1: 332510C (already exists)
SIN2: (empty)
SIN3: (empty)
```

**Script Behavior:**
- SIN1: Already exists (332510C) - SKIPPED ⏭️
- SIN2: Scrapes from Page 2 → Finds 665210 ✓
- SIN3: Early stop (already have 2 SINs)

**Excel After:**
```
SIN1: 332510C (kept existing)
SIN2: 665210 (newly scraped)
SIN3: (empty - skipped)
```

### Scenario 2: All SINs found (early stop)
**Page 1**: Schedule/SIN: MAS/332510C  
**Page 2**: Schedule/SIN: MAS/665210  
**Page 3**: ⏭️ Skipped (already have 2 SINs)

**Output in Excel:**
```
SIN1: 332510C
SIN2: 665210
SIN3: (empty - skipped)
```

### Scenario 3: One SIN missing
**Page 1**: No SIN found  
**Page 2**: Schedule/SIN: MAS/332510C  
**Page 3**: Schedule/SIN: MAS/665210

**Output in Excel:**
```
SIN1: (empty)
SIN2: 332510C
SIN3: 665210
```

### Scenario 4: Limited links
**Page 1**: Schedule/SIN: MAS/332510C  
**Page 2**: No link available  
**Page 3**: Schedule/SIN: MAS/665210

**Output in Excel:**
```
SIN1: 332510C
SIN2: (empty - no link)
SIN3: 665210
```

## Performance

- **Average**: ~7-10 seconds per product (now faster with max 2 SINs)
- **Full run**: ~8-12 hours for 18,000 products
- **Early stops**: ~30-40% faster when 2 SINs found early
- **Progress saved**: Every 50 rows (optimized for speed)

## Reliability for Long-Running Operations

### ✅ YES - This script is designed for multi-hour runs

The script includes several features specifically for long-running stability:

1. **Automatic Recovery**
   - Detects browser crashes and automatically restarts
   - Continues from where it left off
   - No manual intervention needed

2. **Memory Management**
   - Browser restarts every 100 rows
   - Prevents memory leaks and performance degradation
   - Keeps scraping speed consistent over hours

3. **Network Resilience**
   - Retry logic (up to 3 attempts per page)
   - Handles temporary network issues
   - Doesn't fail entire session for one bad request

4. **Data Safety**
   - Auto-saves every 50 rows
   - Can interrupt and resume anytime
   - Multiple backups prevent data loss

5. **Resumable**
   - Use "Custom Range" to continue from any row
   - Check Excel to see which rows need SINs
   - Start from that row and continue

### Recommended Approach for 18,000+ Rows

**Option 1: Run in batches**
```
Day 1: Rows 1-5000 (4-5 hours)
Day 2: Rows 5001-10000 (4-5 hours)
Day 3: Rows 10001-15000 (4-5 hours)
Day 4: Rows 15001-18409 (3-4 hours)
```

**Option 2: Run overnight**
```
Start before bed, let it run 10-12 hours
Auto-saves protect your progress
Check results in the morning
```

**Option 3: Resume on failure**
```
Run full automation
If interrupted (power, network, etc.)
Check last saved row in Excel
Resume from Custom Range
```

### What if something goes wrong?

- **Browser crash**: Auto-restarts, continues
- **Network issue**: Retries 3 times, then skips row
- **Power outage**: Last auto-save is in Excel (max 50 rows lost)
- **Computer sleep**: Wake up, resume from Custom Range
- **Want to stop**: Press Ctrl+C **once**, wait 10-20s for graceful shutdown ✅
- **Accidental Ctrl+C**: File saved before exit, **no corruption** ✅
- **Force shutdown**: Press Ctrl+C **twice** (not recommended, may corrupt)

**Bottom line**: Safe to run overnight or over multiple days! 🚀

**New**: Ctrl+C is now 100% safe - just press once and wait!

## Troubleshooting

### Common Issues

1. **"Excel file not found"**
   - Make sure `ScrappedProducts.xlsx` exists in root folder
   - Script looks for: `../../ScrappedProducts.xlsx`

2. **"ChromeDriver not found"**
   - Install ChromeDriver or use webdriver-manager
   - Ensure Chrome browser is up to date

3. **"No SIN found"**
   - Some products may not have SIN on their page
   - Script logs detailed information about what it found

4. **"Permission denied when saving"**
   - Close Excel file before running script
   - File must not be open in another application

## Notes

- Script uses Selenium because GSA pages have dynamic content
- Rate limiting prevents server overload
- Backup created before each save
- Can resume from any row using Custom Range mode

## Author
Created: January 2026
Purpose: GSA SIN extraction automation
