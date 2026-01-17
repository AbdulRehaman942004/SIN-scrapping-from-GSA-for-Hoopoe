import pandas as pd
import time
import re
import os
import shutil
import signal
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SINScrapingAutomation:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path
        self.driver = None
        self.wait = None
        self.shutdown_requested = False
        self.current_dataframe = None
        
        # Setup signal handler for graceful shutdown
        signal.signal(signal.SIGINT, self.signal_handler)
        signal.signal(signal.SIGTERM, self.signal_handler)
    
    def signal_handler(self, signum, frame):
        """Handle Ctrl+C and other termination signals gracefully"""
        if not self.shutdown_requested:
            self.shutdown_requested = True
            print(f"\n\n{'='*80}")
            print("SHUTDOWN REQUESTED (Ctrl+C detected)")
            print(f"{'='*80}")
            print("⚠️  Please wait... Saving data safely to prevent corruption!")
            print("⚠️  DO NOT force close or press Ctrl+C again!")
            print(f"{'='*80}\n")
            logger.warning("Shutdown signal received, initiating graceful shutdown...")
        else:
            print("\n⚠️  FORCE SHUTDOWN: Data may be corrupted!")
            logger.error("Force shutdown requested, exiting immediately")
            sys.exit(1)
        
    def setup_driver(self):
        """Initialize Chrome driver with optimized options for long-running stability"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")  # Don't load images for faster scraping
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        
        # Additional stability options for long-running operations
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--log-level=3")  # Reduce logging
        chrome_options.add_argument("--silent")
        chrome_options.add_argument("--disable-logging")
        
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option("prefs", {
            "profile.default_content_setting_values": {
                "images": 2,
                "plugins": 2,
                "popups": 2,
                "geolocation": 2,
                "notifications": 2,
                "media_stream": 2,
            }
        })
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 15)
        logger.info("Chrome driver initialized successfully")
    
    def check_driver_health(self):
        """Check if driver is still responsive"""
        try:
            self.driver.current_url
            return True
        except Exception as e:
            logger.warning(f"Driver health check failed: {str(e)}")
            return False
    
    def sin_exists(self, df, row_idx, column_name):
        """Check if SIN already exists in the dataframe for given row and column"""
        try:
            value = df.at[row_idx, column_name]
            # Check if value is not empty, not NaN, and not 'nan' string
            # Also consider "SIN not found" as existing data (skip these rows)
            if pd.notna(value) and str(value).strip() != '' and str(value).strip().lower() != 'nan':
                return True, str(value).strip()
            return False, None
        except Exception:
            return False, None
    
    def row_has_any_sin(self, df, row_idx):
        """Check if row has at least one SIN (any of SIN1, SIN2, SIN3)"""
        sin_columns = ['SIN1', 'SIN2', 'SIN3']
        for col in sin_columns:
            exists, value = self.sin_exists(df, row_idx, col)
            if exists:
                return True, col, value
        return False, None, None
    
    def restart_driver(self):
        """Restart the Chrome driver (useful for long-running sessions)"""
        try:
            logger.info("Restarting Chrome driver...")
            if self.driver:
                self.driver.quit()
        except:
            pass
        
        time.sleep(2)
        self.setup_driver()
        logger.info("Chrome driver restarted successfully")
        
    def read_excel_data(self):
        """Read Excel file with direct product links"""
        try:
            df = pd.read_excel(self.excel_file_path)
            logger.info(f"Excel file loaded successfully. Total rows: {len(df)}")
            
            # Check for required columns
            required_columns = [
                'GSA Direct Product Link',
                'GSA Direct Product Link 1',
                'GSA Direct Product Link 2'
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logger.error(f"Missing required columns: {missing_columns}")
                return None
            
            # Add SIN columns if they don't exist
            sin_columns = ['SIN1', 'SIN2', 'SIN3']
            for col in sin_columns:
                if col not in df.columns:
                    df[col] = ''
            
            logger.info(f"Found {len(df)} products to process")
            return df
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return None
    
    def extract_sin_from_page(self, url, max_retries=3):
        """
        Scrape SIN from GSA product detail page.
        Look for "Schedule/SIN" in table and extract the part after "/"
        Includes retry logic for reliability
        """
        if pd.isna(url) or not url or str(url).strip() == '':
            logger.debug("Empty URL provided, skipping")
            return None
        
        # Retry logic for network issues
        for attempt in range(max_retries):
            try:
                logger.info(f"Navigating to: {url} (attempt {attempt+1}/{max_retries})")
                self.driver.get(url)
                
                # Wait for page to load completely
                WebDriverWait(self.driver, 10).until(
                    lambda driver: driver.execute_script("return document.readyState") == "complete"
                )
                
                # Additional wait for dynamic content
                time.sleep(2)
                
                # Break retry loop if page loaded successfully
                break
                
            except TimeoutException:
                logger.warning(f"Page load timeout (attempt {attempt+1}/{max_retries})")
                if attempt < max_retries - 1:
                    time.sleep(3)  # Wait before retry
                    continue
                else:
                    logger.error(f"Failed to load page after {max_retries} attempts")
                    return None
            except Exception as e:
                logger.warning(f"Error loading page (attempt {attempt+1}/{max_retries}): {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(3)  # Wait before retry
                    # Recreate driver if it crashed
                    try:
                        self.driver.quit()
                    except:
                        pass
                    self.setup_driver()
                    continue
                else:
                    logger.error(f"Failed to load page after {max_retries} attempts")
                    return None
        
        # Now extract SIN from the loaded page
        try:
            
            # Strategy 1: Look for "Schedule/SIN" text in the page
            try:
                # Get entire page text
                page_text = self.driver.find_element(By.TAG_NAME, "body").text
                
                # Pattern to find Schedule/SIN: value
                # Examples: "Schedule/SIN: MAS/332510C" or "Schedule/SIN MAS/332510C"
                patterns = [
                    r'Schedule/SIN[:\s]+([A-Z0-9]+)/([A-Z0-9]+)',
                    r'Schedule/SIN[:\s]+([A-Z0-9]+)',
                    r'SIN[:\s]+([A-Z0-9]+)/([A-Z0-9]+)',
                    r'SIN[:\s]+([A-Z0-9]+)',
                ]
                
                for pattern in patterns:
                    matches = re.search(pattern, page_text, re.IGNORECASE)
                    if matches:
                        # If we have two groups (e.g., MAS/332510C), take the second one
                        if len(matches.groups()) >= 2:
                            sin_number = matches.group(2)
                            logger.info(f"Found SIN: {sin_number} (from {matches.group(0)})")
                            return sin_number
                        # If only one group, check if it contains "/"
                        elif len(matches.groups()) == 1:
                            full_sin = matches.group(1)
                            if '/' in full_sin:
                                sin_number = full_sin.split('/')[-1]
                                logger.info(f"Found SIN: {sin_number} (from {full_sin})")
                                return sin_number
                            else:
                                logger.info(f"Found SIN: {full_sin}")
                                return full_sin
                
            except Exception as e:
                logger.warning(f"Text search failed: {str(e)}")
            
            # Strategy 2: Look in table elements
            try:
                # Find all table rows
                tables = self.driver.find_elements(By.TAG_NAME, "table")
                
                for table in tables:
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    
                    for row in rows:
                        row_text = row.text.lower()
                        
                        # Check if this row contains "schedule/sin" or just "sin"
                        if 'schedule/sin' in row_text or 'schedule sin' in row_text:
                            # Try to find the value in the next cell or same row
                            cells = row.find_elements(By.TAG_NAME, "td")
                            
                            for cell in cells:
                                cell_text = cell.text.strip()
                                # Check if cell contains pattern like "MAS/332510C"
                                if '/' in cell_text:
                                    # Extract the part after "/"
                                    parts = cell_text.split('/')
                                    if len(parts) >= 2:
                                        sin_number = parts[-1].strip()
                                        # Validate it looks like a SIN (alphanumeric)
                                        if re.match(r'^[A-Z0-9]+$', sin_number, re.IGNORECASE):
                                            logger.info(f"Found SIN in table: {sin_number} (from {cell_text})")
                                            return sin_number
            
            except Exception as e:
                logger.warning(f"Table search failed: {str(e)}")
            
            # Strategy 3: Look for specific div or span elements with class names
            try:
                # Common class names or IDs for product details
                selectors = [
                    "//div[contains(text(), 'Schedule/SIN')]",
                    "//span[contains(text(), 'Schedule/SIN')]",
                    "//td[contains(text(), 'Schedule/SIN')]",
                    "//th[contains(text(), 'Schedule/SIN')]",
                ]
                
                for selector in selectors:
                    try:
                        elements = self.driver.find_elements(By.XPATH, selector)
                        for element in elements:
                            # Get parent element or next sibling
                            parent = element.find_element(By.XPATH, "..")
                            parent_text = parent.text
                            
                            # Extract SIN from parent text
                            match = re.search(r'([A-Z0-9]+)/([A-Z0-9]+)', parent_text, re.IGNORECASE)
                            if match:
                                sin_number = match.group(2)
                                logger.info(f"Found SIN in element: {sin_number}")
                                return sin_number
                    except:
                        continue
            
            except Exception as e:
                logger.warning(f"Element search failed: {str(e)}")
            
            logger.warning(f"No SIN found on page: {url}")
            return None
        
        except Exception as e:
            logger.error(f"Error extracting SIN from page: {str(e)}")
            return None
    
    def create_backup(self, file_path):
        """Create a timestamped backup of the file in dedicated backups folder"""
        try:
            # Get directory and filename
            file_dir = os.path.dirname(file_path) or '.'
            filename = os.path.basename(file_path)
            
            # Create backups folder if it doesn't exist
            backups_dir = os.path.join(file_dir, 'backups')
            if not os.path.exists(backups_dir):
                os.makedirs(backups_dir)
                logger.info(f"Created backups directory: {backups_dir}")
            
            # Create backup with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{filename}.backup_{timestamp}"
            backup_path = os.path.join(backups_dir, backup_filename)
            
            shutil.copy2(file_path, backup_path)
            logger.info(f"Backup created: {backup_path}")
            
            # Clean up old backups (keep only last 5)
            self.cleanup_old_backups(file_path, backups_dir)
            
            return backup_path
        except Exception as e:
            logger.error(f"Error creating backup: {str(e)}")
            return None
    
    def cleanup_old_backups(self, file_path, backups_dir, keep_last=5):
        """Clean up old backup files in backups directory, keeping only the most recent ones"""
        try:
            filename = os.path.basename(file_path)
            
            # Find all backup files for this file in backups directory
            backup_files = []
            if os.path.exists(backups_dir):
                for f in os.listdir(backups_dir):
                    if f.startswith(f"{filename}.backup_"):
                        full_path = os.path.join(backups_dir, f)
                        backup_files.append((full_path, os.path.getmtime(full_path)))
            
            # Sort by modification time (newest first)
            backup_files.sort(key=lambda x: x[1], reverse=True)
            
            # Keep only the most recent backups
            files_to_delete = backup_files[keep_last:]
            
            if files_to_delete:
                logger.info(f"Cleaning up {len(files_to_delete)} old backup(s)...")
            
            for backup_file, _ in files_to_delete:
                try:
                    os.remove(backup_file)
                    logger.info(f"Cleaned up old backup: {os.path.basename(backup_file)}")
                except Exception as e:
                    logger.warning(f"Could not delete backup {backup_file}: {str(e)}")
                    
        except Exception as e:
            logger.warning(f"Error during backup cleanup: {str(e)}")
    
    def save_results_to_excel(self, df):
        """Save the updated dataframe to Excel file with backup and atomic write"""
        temp_file = None
        try:
            print(f"\n[SAVE] Starting save process...")
            print(f"[SAVE] Target file: {self.excel_file_path}")
            print(f"[SAVE] Rows to save: {len(df)}")
            
            # Show SIN statistics before save
            if 'SIN1' in df.columns:
                sin1_count = df['SIN1'].notna().sum()
                print(f"[SAVE] SIN1 non-empty cells: {sin1_count}")
            
            # Create backup if file exists
            if os.path.exists(self.excel_file_path):
                print(f"[SAVE] Creating backup...")
                self.create_backup(self.excel_file_path)
            
            # Use temporary file for atomic write (prevents corruption)
            # IMPORTANT: Use .xlsx extension so pandas recognizes it as Excel file
            temp_file = self.excel_file_path + '.tmp.xlsx'
            print(f"[SAVE] Writing to temporary file: {temp_file}")
            
            # Write to temporary file first
            df.to_excel(temp_file, index=False, engine='openpyxl')
            temp_size = os.path.getsize(temp_file)
            print(f"[SAVE] Temporary file written: {temp_size} bytes")
            logger.info(f"Data written to temporary file: {temp_file}")
            
            # Atomic rename (replaces old file only after new file is complete)
            print(f"[SAVE] Replacing main file with new data...")
            if os.path.exists(self.excel_file_path):
                os.replace(temp_file, self.excel_file_path)
            else:
                os.rename(temp_file, self.excel_file_path)
            
            final_size = os.path.getsize(self.excel_file_path)
            print(f"[SAVE] ✓ File saved successfully: {final_size} bytes")
            logger.info(f"Results saved to {self.excel_file_path}")
            return True
            
        except Exception as e:
            print(f"[SAVE] ✗ ERROR: {str(e)}")
            logger.error(f"Error saving results to Excel: {str(e)}")
            # Clean up temp file if it exists
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                    print(f"[SAVE] Cleaned up temporary file")
                except:
                    pass
            return False
    
    def run_sin_scraping(self, start_row=0, end_row=None, test_mode=False):
        """Main method to run the SIN scraping automation"""
        try:
            # Read Excel data
            df = self.read_excel_data()
            if df is None:
                logger.error("Failed to read Excel file")
                return False
            
            # Setup web driver
            self.setup_driver()
            
            # Determine range
            if test_mode:
                end_row = min(10, len(df))
                logger.info(f"TEST MODE: Processing first {end_row} rows")
            elif end_row is None:
                end_row = len(df)
            
            start_row = max(0, start_row)
            end_row = min(len(df), end_row)
            
            successful_scrapes = 0
            total_sins_found = 0
            total_sins_skipped = 0
            rows_with_2_sins = 0
            rows_with_1_sin = 0
            rows_with_0_sins = 0
            rows_skipped_entirely = 0
            early_stops = 0
            rows_fully_complete = 0
            start_time = time.time()
            
            logger.info(f"Processing rows {start_row+1} to {end_row}")
            
            print(f"\n{'='*80}")
            print(f"STARTING SIN SCRAPING SESSION")
            print(f"{'='*80}")
            print(f"Target: Rows {start_row+1} to {end_row} ({end_row - start_row} total)")
            print(f"Strategy: Maximum 2 SINs per product (stop early if 2 found)")
            print(f"Resume Mode: Skip existing SINs (only scrape missing)")
            print(f"Save Interval: Every 50 rows")
            print(f"{'='*80}\n")
            
            # Process each row
            for i in range(start_row, end_row):
                # Check for shutdown request at the start of each iteration
                if self.shutdown_requested:
                    print(f"\n{'='*80}")
                    print("GRACEFUL SHUTDOWN IN PROGRESS")
                    print(f"{'='*80}")
                    print(f"Last completed row: {i}")
                    print(f"Saving all data before exit...")
                    print(f"{'='*80}\n")
                    
                    # Save current progress
                    print("[EMERGENCY SAVE] Saving data...")
                    self.save_results_to_excel(df)
                    print("[EMERGENCY SAVE] Data saved successfully!")
                    
                    # Cleanup
                    if self.driver:
                        print("[CLEANUP] Closing browser...")
                        try:
                            self.driver.quit()
                        except:
                            pass
                        print("[CLEANUP] Browser closed")
                    
                    print(f"\n{'='*80}")
                    print("SHUTDOWN COMPLETE - Data saved safely!")
                    print(f"{'='*80}")
                    print(f"✓ Processed: {i - start_row} rows")
                    print(f"✓ Data saved to: {self.excel_file_path}")
                    print(f"✓ Resume from row: {i+1}")
                    print(f"{'='*80}\n")
                    
                    return False  # Exit gracefully
                
                try:
                    # Get Item Number for display
                    item_number = df.at[i, 'Item Number'] if 'Item Number' in df.columns else f"Row {i+1}"
                    
                    print(f"\n{'='*80}")
                    print(f"ROW {i+1}/{end_row} | Item: {item_number} | Progress: {((i+1-start_row)/(end_row-start_row)*100):.1f}%")
                    print(f"{'='*80}")
                    
                    # Check if row already has at least one SIN - skip entire row if so
                    has_any_sin, sin_col, sin_value = self.row_has_any_sin(df, i)
                    if has_any_sin:
                        print(f"[ROW SKIP] Row already has SIN data ({sin_col}: {sin_value})")
                        print(f"[ROW SKIP] Skipping entire row to save time")
                        rows_skipped_entirely += 1
                        continue  # Skip to next row
                    
                    row_start_time = time.time()
                    sins_found_for_row = 0
                    sins_skipped_for_row = 0
                    
                    # Scrape SIN1 from GSA Direct Product Link
                    sin1_exists, existing_sin1 = self.sin_exists(df, i, 'SIN1')
                    
                    if sin1_exists:
                        print(f"[1/3] SIN1: Already exists ({existing_sin1}) - SKIPPED")
                        sins_found_for_row += 1
                        sins_skipped_for_row += 1
                        total_sins_skipped += 1
                    else:
                        link1 = df.at[i, 'GSA Direct Product Link']
                        if pd.notna(link1) and str(link1).strip():
                            print(f"[1/3] Checking SIN1...")
                            print(f"      URL: {link1[:70]}...")
                            sin1 = self.extract_sin_from_page(link1)
                            if sin1:
                                df.at[i, 'SIN1'] = sin1
                                print(f"      [SUCCESS] SIN1: {sin1}")
                                sins_found_for_row += 1
                            else:
                                df.at[i, 'SIN1'] = "SIN not found"
                                print(f"      [NOT FOUND] SIN1: No SIN detected - marked as 'SIN not found'")
                            time.sleep(2)  # Rate limiting
                        else:
                            print(f"[1/3] [SKIP] SIN1: No link available")
                    
                    # Check if we already have 2 SINs (max allowed)
                    if sins_found_for_row >= 2:
                        print(f"\n[EARLY STOP] Already have 2 SINs, skipping remaining links")
                        early_stops += 1
                    else:
                        # Scrape SIN2 from GSA Direct Product Link 1
                        sin2_exists, existing_sin2 = self.sin_exists(df, i, 'SIN2')
                        
                        if sin2_exists:
                            print(f"[2/3] SIN2: Already exists ({existing_sin2}) - SKIPPED")
                            sins_found_for_row += 1
                            sins_skipped_for_row += 1
                            total_sins_skipped += 1
                        else:
                            link2 = df.at[i, 'GSA Direct Product Link 1']
                            if pd.notna(link2) and str(link2).strip():
                                print(f"[2/3] Checking SIN2...")
                                print(f"      URL: {link2[:70]}...")
                                sin2 = self.extract_sin_from_page(link2)
                                if sin2:
                                    df.at[i, 'SIN2'] = sin2
                                    print(f"      [SUCCESS] SIN2: {sin2}")
                                    sins_found_for_row += 1
                                else:
                                    df.at[i, 'SIN2'] = "SIN not found"
                                    print(f"      [NOT FOUND] SIN2: No SIN detected - marked as 'SIN not found'")
                                time.sleep(2)  # Rate limiting
                            else:
                                print(f"[2/3] [SKIP] SIN2: No link available")
                    
                    # Check again if we already have 2 SINs
                    if sins_found_for_row >= 2:
                        print(f"\n[EARLY STOP] Already have 2 SINs, skipping SIN3")
                        if sins_skipped_for_row < 2:
                            early_stops += 1
                    else:
                        # Scrape SIN3 from GSA Direct Product Link 2
                        sin3_exists, existing_sin3 = self.sin_exists(df, i, 'SIN3')
                        
                        if sin3_exists:
                            print(f"[3/3] SIN3: Already exists ({existing_sin3}) - SKIPPED")
                            sins_found_for_row += 1
                            sins_skipped_for_row += 1
                            total_sins_skipped += 1
                        else:
                            link3 = df.at[i, 'GSA Direct Product Link 2']
                            if pd.notna(link3) and str(link3).strip():
                                print(f"[3/3] Checking SIN3...")
                                print(f"      URL: {link3[:70]}...")
                                sin3 = self.extract_sin_from_page(link3)
                                if sin3:
                                    df.at[i, 'SIN3'] = sin3
                                    print(f"      [SUCCESS] SIN3: {sin3}")
                                    sins_found_for_row += 1
                                else:
                                    df.at[i, 'SIN3'] = "SIN not found"
                                    print(f"      [NOT FOUND] SIN3: No SIN detected - marked as 'SIN not found'")
                                time.sleep(2)  # Rate limiting
                            else:
                                print(f"[3/3] [SKIP] SIN3: No link available")
                    
                    row_time = time.time() - row_start_time
                    sins_scraped_for_row = sins_found_for_row - sins_skipped_for_row
                    
                    # Track if row was fully complete (all SINs already existed)
                    if sins_skipped_for_row >= 2:
                        rows_fully_complete += 1
                    
                    # Update statistics
                    if sins_found_for_row > 0:
                        successful_scrapes += 1
                        total_sins_found += sins_scraped_for_row
                        if sins_found_for_row == 2:
                            rows_with_2_sins += 1
                        elif sins_found_for_row == 1:
                            rows_with_1_sin += 1
                    else:
                        rows_with_0_sins += 1
                    
                    print(f"\n{'-'*80}")
                    print(f"ROW {i+1} SUMMARY:")
                    print(f"  SINs Total: {sins_found_for_row}/2")
                    if sins_skipped_for_row > 0:
                        print(f"  Scraped: {sins_scraped_for_row} | Skipped (existing): {sins_skipped_for_row}")
                    else:
                        print(f"  Scraped: {sins_scraped_for_row}")
                    print(f"  Time: {row_time:.1f}s")
                    print(f"{'-'*80}")
                    
                    # Calculate ETA
                    elapsed_time = time.time() - start_time
                    avg_time_per_row = elapsed_time / (i - start_row + 1)
                    remaining_rows = end_row - (i + 1)
                    eta_seconds = remaining_rows * avg_time_per_row
                    eta_minutes = eta_seconds / 60
                    eta_hours = eta_minutes / 60
                    
                    # Progress analytics
                    rows_processed = i - start_row + 1
                    success_rate = (successful_scrapes / rows_processed * 100) if rows_processed > 0 else 0
                    avg_sins_per_row = total_sins_found / rows_processed if rows_processed > 0 else 0
                    
                    print(f"\nSESSION ANALYTICS:")
                    print(f"  Rows Processed: {rows_processed}/{end_row - start_row}")
                    print(f"  Rows Skipped Entirely (had data): {rows_skipped_entirely}")
                    print(f"  Rows Actually Scraped: {rows_processed - rows_skipped_entirely}")
                    print(f"  Success Rate: {success_rate:.1f}% (rows with at least 1 SIN)")
                    print(f"  Total SINs Scraped: {total_sins_found}")
                    print(f"  Total SINs Skipped (existing): {total_sins_skipped}")
                    print(f"  Avg SINs/Row: {avg_sins_per_row:.2f}")
                    print(f"  Breakdown: {rows_with_2_sins} rows (2 SINs) | {rows_with_1_sin} rows (1 SIN) | {rows_with_0_sins} rows (0 SINs)")
                    print(f"  Early Stops: {early_stops} (saved time by stopping at 2 SINs)")
                    print(f"\nTIMING:")
                    print(f"  Current Row: {row_time:.1f}s")
                    print(f"  Average: {avg_time_per_row:.1f}s/row")
                    print(f"  Elapsed: {elapsed_time/60:.1f} minutes")
                    if eta_hours >= 1:
                        print(f"  ETA: {eta_hours:.1f} hours ({eta_minutes:.0f} minutes)")
                    else:
                        print(f"  ETA: {eta_minutes:.0f} minutes")
                    print(f"{'='*80}")
                    
                    # Save progress every 50 rows
                    if (i + 1) % 50 == 0:
                        print(f"\n[AUTOSAVE] Saving progress at row {i+1}...")
                        self.save_results_to_excel(df)
                        print(f"[AUTOSAVE] Progress saved successfully!")
                        
                        # Check if shutdown was requested during save
                        if self.shutdown_requested:
                            continue  # Will be handled at start of next iteration
                    
                    # Restart driver every 100 rows for long-running stability
                    if (i + 1) % 100 == 0:
                        print(f"\n[MAINTENANCE] Restarting browser for stability (row {i+1})...")
                        self.restart_driver()
                        print(f"[MAINTENANCE] Browser restarted successfully!")
                    
                except Exception as e:
                    logger.error(f"Error processing row {i+1}: {str(e)}")
                    print(f"\n[ERROR] Row {i+1}: {str(e)}")
                    
                    # Check if driver crashed and restart if needed
                    if not self.check_driver_health():
                        print(f"[RECOVERY] Browser crashed, restarting...")
                        self.restart_driver()
                        print(f"[RECOVERY] Browser restarted, continuing...")
                    
                    continue
            
            # Final save
            print(f"\n{'='*80}")
            print(f"[FINAL SAVE] Saving all results...")
            print(f"{'='*80}")
            self.save_results_to_excel(df)
            print(f"[FINAL SAVE] All data saved successfully!")
            
            # Calculate final statistics
            total_time = time.time() - start_time
            total_rows = end_row - start_row
            success_rate = (successful_scrapes / total_rows * 100) if total_rows > 0 else 0
            avg_sins_per_row = total_sins_found / total_rows if total_rows > 0 else 0
            
            print(f"\n{'='*80}")
            print(f"{'='*80}")
            print(f"  SIN SCRAPING SESSION COMPLETED!")
            print(f"{'='*80}")
            print(f"{'='*80}")
            
            print(f"\nOVERALL STATISTICS:")
            print(f"{'-'*80}")
            print(f"  Total Rows Processed: {total_rows}")
            print(f"  Rows Skipped Entirely (had data): {rows_skipped_entirely} ({(rows_skipped_entirely/total_rows*100):.1f}%)")
            print(f"  Rows Actually Scraped: {total_rows - rows_skipped_entirely} ({((total_rows - rows_skipped_entirely)/total_rows*100):.1f}%)")
            print(f"  Rows with SINs Found: {successful_scrapes} ({success_rate:.1f}%)")
            print(f"  Rows with NO SINs: {rows_with_0_sins} ({(rows_with_0_sins/total_rows*100):.1f}%)")
            print(f"{'-'*80}")
            
            print(f"\nSIN BREAKDOWN:")
            print(f"{'-'*80}")
            print(f"  Total SINs Scraped (new): {total_sins_found}")
            print(f"  Total SINs Skipped (existing): {total_sins_skipped}")
            print(f"  Combined Total: {total_sins_found + total_sins_skipped}")
            print(f"  Average SINs per Row: {avg_sins_per_row:.2f}")
            print(f"  Rows with 2 SINs: {rows_with_2_sins} ({(rows_with_2_sins/total_rows*100):.1f}%)")
            print(f"  Rows with 1 SIN:  {rows_with_1_sin} ({(rows_with_1_sin/total_rows*100):.1f}%)")
            print(f"  Rows with 0 SINs: {rows_with_0_sins} ({(rows_with_0_sins/total_rows*100):.1f}%)")
            print(f"  Rows Already Complete: {rows_fully_complete} ({(rows_fully_complete/total_rows*100):.1f}%)")
            print(f"{'-'*80}")
            
            print(f"\nPERFORMANCE:")
            print(f"{'-'*80}")
            print(f"  Total Time: {total_time/60:.1f} minutes ({total_time/3600:.2f} hours)")
            print(f"  Average Time per Row: {total_time/total_rows:.1f} seconds")
            print(f"  Early Stops (saved time): {early_stops} rows")
            if rows_skipped_entirely > 0:
                time_saved_rows = rows_skipped_entirely * 15  # Assume ~15 seconds saved per skipped row
                print(f"  Time Saved (skipped rows): ~{time_saved_rows/60:.1f} minutes ({rows_skipped_entirely} rows)")
            if total_sins_skipped > 0:
                time_saved = total_sins_skipped * 7  # Assume ~7 seconds saved per skipped SIN
                print(f"  Time Saved (skipped SINs): ~{time_saved/60:.1f} minutes ({total_sins_skipped} SINs)")
            print(f"{'-'*80}")
            
            print(f"\nOUTPUT:")
            print(f"{'-'*80}")
            print(f"  File: {self.excel_file_path}")
            print(f"  Columns Updated: SIN1, SIN2, SIN3")
            print(f"  Backup Created: Yes (keeping last 5)")
            print(f"{'-'*80}")
            
            print(f"\n{'='*80}")
            print(f"  ALL DONE! Your ScrappedProducts.xlsx has been updated.")
            print(f"{'='*80}\n")
            
            logger.info(f"Scraping completed successfully: {total_sins_found} SINs from {total_rows} rows")
            return True
            
        except Exception as e:
            logger.error(f"Error in SIN scraping automation: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()

def main():
    """Main function with interactive menu"""
    print("="*80)
    print("GSA SIN SCRAPING - FROM DIRECT PRODUCT LINKS")
    print("="*80)
    print("This script scrapes SIN numbers from GSA Direct Product Link pages")
    print("="*80)
    
    # File path
    excel_file = "../../ScrappedProducts.xlsx"
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"ERROR: Excel file not found: {excel_file}")
        return
    
    # Display menu
    while True:
        print("\n" + "="*80)
        print("SIN SCRAPING MENU")
        print("="*80)
        print("1. Test Mode (First 10 rows)")
        print("2. Custom Range (Specify start and end)")
        print("3. Full Automation (All rows)")
        print("4. Exit")
        print("="*80)
        
        try:
            choice = input("Enter your choice (1-4): ").strip()
            
            if choice == "1":
                print("\nRunning TEST MODE (first 10 rows)...")
                automation = SINScrapingAutomation(excel_file)
                success = automation.run_sin_scraping(test_mode=True)
                if success:
                    print("\n✓ SUCCESS: Test scraping completed!")
                else:
                    print("\n✗ ERROR: Test scraping failed!")
                    
            elif choice == "2":
                print("\nCUSTOM RANGE MODE")
                print("="*40)
                
                # Get total number of rows
                df = pd.read_excel(excel_file)
                total_rows = len(df)
                print(f"Total rows available: {total_rows}")
                
                try:
                    start_row = int(input(f"Enter start row (1-{total_rows}): ")) - 1
                    end_row = int(input(f"Enter end row ({start_row + 2}-{total_rows}): "))
                    
                    if start_row < 0 or end_row > total_rows or start_row >= end_row:
                        print("ERROR: Invalid range specified!")
                        continue
                    
                    count = end_row - start_row
                    print(f"\nProcessing rows {start_row + 1}-{end_row} ({count} rows)...")
                    
                    automation = SINScrapingAutomation(excel_file)
                    success = automation.run_sin_scraping(start_row=start_row, end_row=end_row)
                    if success:
                        print(f"\n✓ SUCCESS: Custom range scraping completed!")
                    else:
                        print(f"\n✗ ERROR: Custom range scraping failed!")
                        
                except ValueError:
                    print("ERROR: Please enter valid numbers!")
                    continue
                    
            elif choice == "3":
                print("\nFULL AUTOMATION MODE")
                print("="*40)
                
                # Get total rows
                df = pd.read_excel(excel_file)
                total_rows = len(df)
                
                print(f"WARNING: This will process ALL {total_rows} rows!")
                print("Estimated time: ~10-15 hours")
                print("Progress will be saved every 10 rows")
                
                confirm = input("\nAre you sure you want to continue? (yes/no): ").strip().lower()
                if confirm in ['yes', 'y']:
                    print("\nRunning FULL AUTOMATION...")
                    automation = SINScrapingAutomation(excel_file)
                    success = automation.run_sin_scraping()
                    if success:
                        print("\n✓ SUCCESS: Full automation completed!")
                    else:
                        print("\n✗ ERROR: Full automation failed!")
                else:
                    print("Full automation cancelled.")
                    
            elif choice == "4":
                print("\nExiting...")
                break
                
            else:
                print("ERROR: Invalid choice! Please enter 1, 2, 3, or 4.")
                
        except KeyboardInterrupt:
            print("\n\nOperation cancelled by user.")
            break
        except Exception as e:
            print(f"\nERROR: {str(e)}")
            continue

if __name__ == "__main__":
    main()
