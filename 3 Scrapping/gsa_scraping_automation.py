import pandas as pd
import time
import re
import os
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from difflib import SequenceMatcher
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class GSAScrapingAutomation:
    def __init__(self, excel_file_path, manufacturer_mapping_file):
        self.excel_file_path = excel_file_path
        self.manufacturer_mapping_file = manufacturer_mapping_file
        self.driver = None
        self.wait = None
        self.manufacturer_mapping = {}
        self.unit_mapping = self._create_unit_mapping()
        # Add caching for performance
        self._manufacturer_normalization_cache = {}
        self._unit_normalization_cache = {}
        
        # Pre-compile regex patterns for better performance
        self._compile_regex_patterns()
        
        # Compile SIN extraction patterns - comprehensive formats
        # Ordered from most specific to more general
        self._sin_patterns = [
            # Format: "Schedule/SIN: MAS/332510C"
            re.compile(r'schedule[/\s]*sin[:\s]*([A-Z]+)[/\s]+([A-Z0-9]{4,15})', re.IGNORECASE),
            # Format: "Schedule/SIN MAS/332510C" (no colon)
            re.compile(r'schedule[/\s]*sin\s+([A-Z]+)[/\s]+([A-Z0-9]{4,15})', re.IGNORECASE),
            # Format: "SIN: 332510C" or "SIN 332510C"
            re.compile(r'sin[:\s]+([A-Z0-9]{4,15})(?:\s|$|\n)', re.IGNORECASE),
            # Format: "MAS/332510C" or "FSS/332510"
            re.compile(r'\b(MAS|FSS|SIN)[/\s]+([0-9]{4,10}[A-Z]{0,3})\b', re.IGNORECASE),
            # Format: Just the SIN number after "Schedule:"
            re.compile(r'schedule[:\s]+([A-Z0-9]{4,15})(?:\s|$|\n)', re.IGNORECASE),
            # Format: 6 digits + letter (e.g., "332510C")
            re.compile(r'\b([0-9]{6}[A-Z]{1,3})\b'),
            # Format: 5-7 digits alone
            re.compile(r'\b([0-9]{5,7})\b'),
        ]
        
    def _create_unit_mapping(self):
        """Create unit of measure standardization mapping"""
        return {
            'each': ['ea', 'piece', 'pc', 'unit', 'u', 'pcs'],
            'box': ['bx', 'case', 'cs', 'carton'],
            'pack': ['pk', 'package', 'pkg'],
            'dozen': ['dz', '12', 'doz'],
            'gross': ['144', 'gr'],
            'ream': ['rm', '500'],
            'roll': ['rl'],
            'set': ['st'],
            'pair': ['pr'],
            'gallon': ['gal', 'g'],
            'pound': ['lb', 'lbs', '#'],
            'ounce': ['oz'],
            'inch': ['in', '"'],
            'foot': ['ft', "'"],
            'yard': ['yd'],
            'meter': ['m'],
            'centimeter': ['cm'],
            'millimeter': ['mm']
        }
    
    def _compile_regex_patterns(self):
        """Pre-compile regex patterns for better performance"""
        # Price patterns
        self._price_patterns = [
            re.compile(r'\$\s*([\d,]+\.?\d*)', re.IGNORECASE),
            re.compile(r'([\d,]+\.\d{2})\s*EA', re.IGNORECASE),
            re.compile(r'([\d,]+\.\d{2})\s*USD', re.IGNORECASE),
            re.compile(r'price[:\s]*\$?\s*([\d,]+\.?\d*)', re.IGNORECASE),
            re.compile(r'unit[:\s]*\$?\s*([\d,]+\.?\d*)', re.IGNORECASE),
            re.compile(r'each[:\s]*\$?\s*([\d,]+\.?\d*)', re.IGNORECASE),
        ]
        
        # Contractor patterns
        # Updated to handle special characters in contractor names (apostrophes, parentheses, slashes, etc.)
        # Using [^\n]+? to match any character except newline, stopping before delimiters
        self._contractor_patterns = [
            re.compile(r'contractor[:\s]*\n([^\n]+?)(?:\n|contract#|Contract#|includes)', re.IGNORECASE | re.MULTILINE),
            re.compile(r'contractor[:\s]*([^\n]+?)(?:\n|contract#|Contract#|includes)', re.IGNORECASE | re.MULTILINE),
            re.compile(r'vendor[:\s]*\n([^\n]+?)(?:\n|contract#|Contract#|includes)', re.IGNORECASE | re.MULTILINE),
            re.compile(r'supplier[:\s]*\n([^\n]+?)(?:\n|contract#|Contract#|includes)', re.IGNORECASE | re.MULTILINE),
            re.compile(r'company[:\s]*\n([^\n]+?)(?:\n|contract#|Contract#|includes)', re.IGNORECASE | re.MULTILINE),
            re.compile(r'distributor[:\s]*\n([^\n]+?)(?:\n|contract#|Contract#|includes)', re.IGNORECASE | re.MULTILINE),
        ]
        
        # Contract patterns
        self._contract_patterns = [
            re.compile(r'contract#:\s*([a-z0-9-]+)', re.IGNORECASE),
            re.compile(r'contract\s*number[:\s#]*([a-z0-9-]+)', re.IGNORECASE),
            re.compile(r'gsa[:\s#]*([a-z0-9-]+)', re.IGNORECASE),
            re.compile(r'gsa\s*contract[:\s#]*([a-z0-9-]+)', re.IGNORECASE),
            re.compile(r'contract[:\s#]*([a-z0-9-]+)', re.IGNORECASE),
        ]
        
        # Manufacturer patterns
        self._manufacturer_patterns = [
            re.compile(r'\bmfr[:\s]*([a-z0-9\s&.,®\-]+)', re.IGNORECASE),
            re.compile(r'\bmanufacturer[:\s]*([a-z0-9\s&.,®\-]+)', re.IGNORECASE),
            re.compile(r'\bmfg[:\s]*([a-z0-9\s&.,®\-]+)', re.IGNORECASE),
            re.compile(r'\bbrand[:\s]*([a-z0-9\s&.,®\-]+)', re.IGNORECASE)
        ]
        
        # Unit patterns
        self._unit_patterns = [
            re.compile(r'\$\s*[\d,]+\.?\d*\s*([a-z]+)', re.IGNORECASE),
            re.compile(r'([a-z]+)\s*from', re.IGNORECASE),
            re.compile(r'unit[:\s]*([a-z0-9\s]+)', re.IGNORECASE),
            re.compile(r'uom[:\s]*([a-z0-9\s]+)', re.IGNORECASE),
            re.compile(r'per[:\s]*([a-z0-9\s]+)', re.IGNORECASE),
            re.compile(r'each[:\s]*([a-z0-9\s]+)', re.IGNORECASE),
        ]
    
    def setup_driver(self, headless=False):
        """Initialize Chrome driver with optimized options for speed"""
        chrome_options = Options()
        
        # Headless mode for overnight runs (faster, less resource-intensive)
        if headless:
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--window-size=1920,1080")
            logger.info("Running in HEADLESS mode")
        
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")  # Don't load images for faster scraping
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_argument("--disable-background-timer-throttling")
        chrome_options.add_argument("--disable-backgrounding-occluded-windows")
        chrome_options.add_argument("--disable-renderer-backgrounding")
        chrome_options.add_argument("--disable-background-networking")
        chrome_options.add_argument("--disable-sync")
        chrome_options.add_argument("--disable-translate")
        chrome_options.add_argument("--disable-ipc-flooding-protection")
        chrome_options.add_argument("--memory-pressure-off")
        chrome_options.add_argument("--max_old_space_size=4096")
        chrome_options.add_argument("--disable-gpu")  # Faster rendering
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--no-default-browser-check")
        
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option("prefs", {
            "profile.default_content_setting_values": {
                "images": 2,  # Block images
                "plugins": 2,  # Block plugins
                "popups": 2,  # Block popups
                "geolocation": 2,  # Block geolocation
                "notifications": 2,  # Block notifications
                "media_stream": 2,  # Block media stream
            }
        })
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # Store headless mode setting for browser restarts
        self._headless_mode = headless
        
        # Optimized timeouts for faster execution
        self.wait = WebDriverWait(self.driver, 10)  # Reduced from 15s
        self.driver.set_page_load_timeout(25)  # Prevent hanging pages
        
    def load_manufacturer_mapping(self):
        """Load manufacturer root form mapping from CSV"""
        try:
            df_mapping = pd.read_csv(self.manufacturer_mapping_file)
            # Primary mapping: original -> root (as provided)
            self.manufacturer_mapping = dict(zip(df_mapping['original'], df_mapping['root']))

            # Build a normalized-key mapping to handle punctuation/symbol variants like "AT-A-GLANCE®"
            # This keeps things generic – no hard-coded brand names
            self._normalized_manufacturer_lookup = {}
            for original, root in zip(df_mapping['original'], df_mapping['root']):
                norm_key = self.normalize_manufacturer(str(original))
                if norm_key:
                    self._normalized_manufacturer_lookup[norm_key] = root

            logger.info(f"Loaded {len(self.manufacturer_mapping)} manufacturer mappings")
            return True
        except Exception as e:
            logger.error(f"Error loading manufacturer mapping: {str(e)}")
            return False
    
    def normalize_manufacturer(self, manufacturer_name):
        """Normalize manufacturer name using the same logic as Step 2 with caching"""
        if not manufacturer_name:
            return ""
        
        # Check cache first
        if manufacturer_name in self._manufacturer_normalization_cache:
            return self._manufacturer_normalization_cache[manufacturer_name]
        
        # Use the same normalization logic as in Step 2
        # This ensures consistency with the root forms in our CSV
        result = self._normalize_to_root_like(manufacturer_name)
        
        # Cache the result
        self._manufacturer_normalization_cache[manufacturer_name] = result
        return result
    
    def _normalize_to_root_like(self, name):
        """Convert manufacturer name to root-like form (same logic as Step 2)"""
        if not name:
            return ""
        
        # Terms to strip (same as Step 2)
        REMOVABLE_TERMS = {
            "inc", "incorporated", "corp", "corporation", "co", "company", "llc", "l.l.c",
            "ltd", "limited", "gmbh", "s.a.", "s.a", "s.p.a.", "spa", "ag", "kg", "nv",
            "plc", "pty", "pte", "sro", "s.r.o", "srl", "lp", "llp", "pc",
            "products", "product", "brands", "brand", "group", "international", "industries",
            "industry", "mfg", "manufacturing", "manufacturers", "division", "div",
            "usa", "u.s.a", "u.s.", "us", "america", "american", "north", "south",
            "europe", "european", "asia", "pacific",
        }
        
        lower = name.lower()
        
        # Tokenize by converting non-alnum to spaces, then splitting
        tokens = re.sub(r"[^0-9a-z]+", " ", lower).split()
        
        filtered = []
        for token in tokens:
            if token not in REMOVABLE_TERMS:
                filtered.append(token)
        
        chosen = ""
        if filtered:
            chosen = filtered[0]
        else:
            # Fallback: take first alphanumeric run from original
            alnum_runs = re.findall(r"[0-9a-z]+", lower)
            chosen = alnum_runs[0] if alnum_runs else ""
        
        # Remove any lingering non-alphanumeric chars
        root = re.sub(r"[^0-9a-z]", "", chosen)
        return root
    
    def normalize_unit(self, unit_name):
        """Normalize unit of measure for matching with caching"""
        if not unit_name:
            return ""
        
        # Check cache first
        if unit_name in self._unit_normalization_cache:
            return self._unit_normalization_cache[unit_name]
        
        normalized = str(unit_name).lower().strip()
        
        # Remove special characters except spaces and alphanumeric
        normalized = re.sub(r'[^a-z0-9\s]', '', normalized)
        
        # Remove extra spaces
        normalized = re.sub(r'\s+', '', normalized)
        
        # Cache the result
        self._unit_normalization_cache[unit_name] = normalized
        return normalized
    
    def fuzzy_match_manufacturer(self, original_manufacturer, website_manufacturer, threshold=0.85):
        """Generic fuzzy match for manufacturer.

        Strategy (generic, no hard-coding of brands):
        1) Use CSV mapping directly (most reliable)
        2) Normalize website manufacturer and check if root appears in it
        3) Fallback to direct normalization comparison
        """
        if not original_manufacturer or not website_manufacturer:
            return False

        # Strategy 1: Use CSV mapping directly (most reliable)
        root_form = self.manufacturer_mapping.get(original_manufacturer)
        if root_form:
            # Deterministic: concatenate website manufacturer to alphanumeric lowercase and check substring
            website_alnum = re.sub(r"[^a-z0-9]", "", str(website_manufacturer).lower())
            if website_alnum and root_form in website_alnum:
                logger.debug(f"CSV mapping match: '{root_form}' found in alnum website '{website_alnum}'")
                return True
            
            # Fallback to previous normalization + fuzzy similarity if needed
            norm_website = self.normalize_manufacturer(website_manufacturer)
            if norm_website:
                # Check if root appears in normalized website name
                if root_form in norm_website:
                    logger.debug(f"CSV mapping match: '{root_form}' found in '{norm_website}'")
                    return True
                
                # Fuzzy similarity with root
                sim_root = SequenceMatcher(None, root_form, norm_website).ratio()
                if sim_root >= threshold:
                    logger.debug(f"CSV mapping fuzzy match: '{root_form}' vs '{norm_website}' = {sim_root:.3f}")
                    return True
            
            # Additional check: see if original manufacturer name appears in website name
            # This handles cases where normalization loses important parts
            original_lower = original_manufacturer.lower()
            website_lower = website_manufacturer.lower()
            
            # Remove common suffixes and check containment
            original_clean = re.sub(r'\s+(inc|incorporated|corp|corporation|co|company|llc|ltd|limited|products|product|brands|brand)$', '', original_lower)
            
            # Also normalize spaces/hyphens for better matching
            original_normalized = re.sub(r'[-\s]+', ' ', original_clean)
            website_normalized = re.sub(r'[-\s]+', ' ', website_lower)
            
            if original_normalized in website_normalized:
                logger.debug(f"Original name containment: '{original_normalized}' found in '{website_normalized}'")
                return True

        # Strategy 2: Try normalized-key mapping if exact missing
        if not root_form:
            norm_key = self.normalize_manufacturer(original_manufacturer)
            if hasattr(self, '_normalized_manufacturer_lookup'):
                root_form = self._normalized_manufacturer_lookup.get(norm_key)
                if root_form:
                    # Deterministic alnum-concat containment on website manufacturer
                    website_alnum = re.sub(r"[^a-z0-9]", "", str(website_manufacturer).lower())
                    if website_alnum and root_form in website_alnum:
                        logger.debug(f"Normalized-key mapping match: '{root_form}' found in alnum website '{website_alnum}'")
                        return True
                    # Fallback to previous normalization containment
                    norm_website = self.normalize_manufacturer(website_manufacturer)
                    if norm_website and root_form in norm_website:
                        logger.debug(f"Normalized-key mapping match: '{root_form}' found in '{norm_website}'")
                        return True

        # Strategy 3: Direct normalization comparison (fallback)
        norm_original = self.normalize_manufacturer(original_manufacturer)
        norm_website = self.normalize_manufacturer(website_manufacturer)
        
        if norm_original and norm_website:
            # Only use containment if both are substantial (not single letters)
            if len(norm_original) >= 3 and len(norm_website) >= 3:
                if norm_original in norm_website or norm_website in norm_original:
                    logger.debug(f"Direct containment match: '{norm_original}' in '{norm_website}'")
                    return True
            
            # Fuzzy similarity fallback - but require higher threshold for short strings
            sim_direct = SequenceMatcher(None, norm_original, norm_website).ratio()
            required_threshold = threshold if len(norm_original) >= 4 and len(norm_website) >= 4 else 0.95
            
            # Don't match if both strings are identical and very short (likely over-normalized)
            if norm_original == norm_website and len(norm_original) <= 2:
                logger.debug(f"Rejecting identical short strings: '{norm_original}' == '{norm_website}'")
            elif sim_direct >= required_threshold:
                logger.debug(f"Direct fuzzy match: '{norm_original}' vs '{norm_website}' = {sim_direct:.3f} (threshold: {required_threshold})")
                return True

        logger.debug(f"No match found for '{original_manufacturer}' vs '{website_manufacturer}'")
        return False
    
    def fuzzy_match_unit(self, original_unit, website_unit, threshold=0.8):
        """Fuzzy match unit of measure"""
        if not original_unit or not website_unit:
            return False
        
        norm_original = self.normalize_unit(original_unit)
        norm_website = self.normalize_unit(website_unit)
        
        # Check direct match first
        if norm_original == norm_website:
            return True
        
        # Check against unit mapping
        for standard_unit, variations in self.unit_mapping.items():
            if norm_original in variations and norm_website in variations:
                return True
        
        # Calculate similarity for fuzzy matching
        similarity = SequenceMatcher(None, norm_original, norm_website).ratio()
        
        logger.debug(f"Unit match: '{original_unit}' vs '{website_unit}' = {similarity:.3f}")
        
        return similarity >= threshold
    
    def read_excel_data(self):
        """Read Excel file with GSA links and product data"""
        try:
            df = pd.read_excel(self.excel_file_path)
            logger.info(f"Excel file loaded successfully. Columns: {list(df.columns)}")
            
            # Find required columns
            required_columns = {
                'stock_number': None,
                'manufacturer': None,
                'unit_of_measure': None,
                'links': None
            }
            
# if you want to search by Column B that is "Item Stock Number-Butted" use this code
            # for col in df.columns:
            #     col_lower = col.lower()
            #     if 'item stock number' in col_lower and 'butted' in col_lower:
            #         required_columns['stock_number'] = col
                   
# if you want to search by Column B that is "Item Number" use this code
            for col in df.columns:
                col_lower = col.lower()
                if 'item number' == col_lower:
                    required_columns['stock_number'] = col
                    
                elif 'manufacturer' in col_lower:
                    required_columns['manufacturer'] = col
                elif 'unit of measure' in col_lower:
                    required_columns['unit_of_measure'] = col
                elif 'links' in col_lower:
                    required_columns['links'] = col
            
            # Check if all required columns are found
            missing_columns = [k for k, v in required_columns.items() if v is None]
            if missing_columns:
                logger.error(f"Missing required columns: {missing_columns}")
                return None
            
            # Add result columns if they don't exist
            result_columns = [
                'GSA_Price_1', 'GSA_Contractor_1', 'GSA_Contract_1',
                'GSA_Price_2', 'GSA_Contractor_2', 'GSA_Contract_2',
                'GSA_Price_3', 'GSA_Contractor_3', 'GSA_Contract_3'
            ]
            
            for col in result_columns:
                if col not in df.columns:
                    df[col] = ''
            
            logger.info(f"Found {len(df)} products to process")
            return df, required_columns
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return None, None
    
    def scrape_gsa_page_for_sins(self, gsa_url, target_manufacturer, target_unit, max_sins=2):
        """Scrape GSA page for SIN numbers by clicking on matching products"""
        try:
            # Verify driver is ready
            if not self.driver:
                logger.error("Driver is not initialized in scrape_gsa_page_for_sins!")
                return []
            
            logger.info(f"Scraping GSA page for SINs: {gsa_url}")
            
            # Navigate to the GSA page - this MUST complete and wait
            try:
                self.driver.get(gsa_url)
                logger.info(f"Page navigation initiated, waiting for page to load...")
                
                # Wait for page to be fully loaded
                try:
                    WebDriverWait(self.driver, 15).until(
                        lambda driver: driver.execute_script("return document.readyState") == "complete"
                    )
                    logger.info("Page readyState is 'complete'")
                except TimeoutException:
                    logger.warning("Page readyState did not become 'complete' within timeout, continuing anyway")
                
                # Additional fixed wait to ensure dynamic content starts loading
                # Increased to prevent "Unexpected Error" from GSA (rate limiting)
                time.sleep(4.0)  # Stay longer on search page to stabilize and avoid GSA errors
                
                # Check for product elements
                def any_product_element_present(driver):
                    selectors = [
                        (By.CSS_SELECTOR, ".productViewControl"),
                        (By.CSS_SELECTOR, "app-ux-product-display-inline"),
                        (By.CSS_SELECTOR, ".product-item"),
                        (By.CSS_SELECTOR, ".result-item"),
                        (By.CSS_SELECTOR, ".product"),
                        (By.XPATH, "//div[contains(@class, 'product')]"),
                        (By.XPATH, "//div[contains(@class, 'result')]")
                    ]
                    for selector_type, selector_value in selectors:
                        try:
                            elements = driver.find_elements(selector_type, selector_value)
                            if elements:
                                return True
                        except:
                            continue
                    return False
                
                # Wait for products to appear
                try:
                    WebDriverWait(self.driver, 10).until(any_product_element_present)
                    logger.info("Product elements detected - page is loaded")
                except TimeoutException:
                    logger.warning("No product elements found within 10 seconds")
                    return []
                
            except Exception as nav_error:
                logger.error(f"Error navigating to page {gsa_url}: {str(nav_error)}")
                return []
            
            # Find products
            products = self._find_product_elements()
            
            if not products:
                logger.warning(f"No products found on page: {gsa_url}")
                return []
            
            logger.info(f"Found {len(products)} products on page")
            print(f"   📦 Found {len(products)} products on page - Looking for {max_sins} matching SIN(s)...")
            
            # Stay on search page longer to avoid rate limiting
            time.sleep(3.0)  # Additional wait on search page before processing products
            logger.info("Waiting on search page to prevent rate limiting...")
            
            # Extract SINs from matching products
            sins_collected = []
            products_checked = 0
            
            # Check if first product is header text
            start_index = 0
            if len(products) > 0:
                first_product_text = products[0].text.lower()
                if any(header_indicator in first_product_text for header_indicator in [
                    'name contract number price', 'contractor name', 'price low to high', 
                    'view as grid', 'sort by', 'filter by'
                ]):
                    start_index = 1
                    logger.info("Skipping first product as it appears to be header text")
            
            # Iterate through products and extract SINs from matching ones
            # Strategy: Check ALL products on page until we collect the required number of SINs
            i = start_index
            
            print(f"   📋 Checking up to {len(products)} products on page...")
            
            while i < len(products):
                # Stop if we have 2 SINs
                if len(sins_collected) >= max_sins:
                    logger.info(f"Collected {max_sins} SINs, stopping...")
                    break
                
                try:
                    # Always re-check products list length (it might have changed after navigation)
                    if i >= len(products):
                        logger.debug(f"Product index {i} out of range (list has {len(products)} products), breaking...")
                        break
                    
                    product_element = products[i]
                    product_text = product_element.text.lower()
                    product_num = i + 1
                    
                    # Extract manufacturer and unit for matching
                    website_manufacturer = self._extract_manufacturer(product_text)
                    website_unit = self._extract_unit(product_text)
                    
                    # Check if manufacturer and unit match
                    manufacturer_match = self.fuzzy_match_manufacturer(target_manufacturer, website_manufacturer)
                    unit_match = self.fuzzy_match_unit(target_unit, website_unit)
                    
                    products_checked += 1
                    
                    # Show progress every 20 products
                    if products_checked % 20 == 0:
                        print(f"   🔍 Checked {products_checked} products, found {len(sins_collected)}/{max_sins} SINs so far...")
                    
                    if manufacturer_match and unit_match:
                        logger.info(f"Product {product_num} MATCHED: Clicking to extract SIN...")
                        print(f"   ✅ Product {product_num} matched - Extracting SIN...")
                        
                        # Click product and extract SIN
                        sin_value = self.click_product_and_extract_sin(product_element, product_num)
                        
                        if sin_value:
                            sins_collected.append(sin_value)
                            logger.info(f"Successfully extracted SIN {len(sins_collected)}/{max_sins}: {sin_value}")
                            print(f"   🎯 SIN extracted: {sin_value}")
                        else:
                            logger.warning(f"Product {product_num} matched but no SIN found")
                            print(f"   ⚠️  Product matched but SIN not found on detail page")
                        
                        # CRITICAL: Re-find products after navigation (DOM has changed)
                        # Don't increment i yet - products list has been reset
                        products = self._find_product_elements()
                        
                        # Skip products we've already checked by starting from where we left off
                        # But since the list was refreshed, we need to stay at the same index
                        # Actually, we need to start fresh from the beginning since we can't track which we checked
                        # Let's just continue from next product
                        i += 1
                        continue
                    else:
                        logger.debug(f"Product {product_num} SKIPPED: Manufacturer match={manufacturer_match}, Unit match={unit_match}")
                    
                    i += 1
                        
                except Exception as e:
                    logger.warning(f"Error processing product {i+1}: {str(e)}")
                    i += 1
                    continue
            
            # If we still need more SINs and haven't checked all products, scroll to load more
            if len(sins_collected) < max_sins and i >= len(products):
                logger.info(f"Found {len(sins_collected)} SIN(s), need {max_sins - len(sins_collected)} more. Scrolling to load more products...")
                print(f"   📜 Scrolling to load more products ({i}/{len(products)} checked so far)...")
                # Scroll to load more products
                try:
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(4.0)  # Stay longer on search page after scrolling to prevent rate limiting
                    products = self._find_product_elements()
                    logger.info(f"After scrolling, found {len(products)} products total")
                    
                    # Continue checking from where we left off
                    while i < len(products):
                        if len(sins_collected) >= max_sins:
                            break
                        
                        if i >= len(products):
                            break
                        
                        try:
                            product_element = products[i]
                            product_text = product_element.text.lower()
                            product_num = i + 1
                            
                            website_manufacturer = self._extract_manufacturer(product_text)
                            website_unit = self._extract_unit(product_text)
                            
                            manufacturer_match = self.fuzzy_match_manufacturer(target_manufacturer, website_manufacturer)
                            unit_match = self.fuzzy_match_unit(target_unit, website_unit)
                            
                            products_checked += 1
                            
                            if manufacturer_match and unit_match:
                                logger.info(f"Product {product_num} MATCHED after scroll: Clicking to extract SIN...")
                                print(f"   ✅ Product {product_num} matched - Extracting SIN...")
                                sin_value = self.click_product_and_extract_sin(product_element, product_num)
                                if sin_value:
                                    sins_collected.append(sin_value)
                                    logger.info(f"Successfully extracted SIN {len(sins_collected)}/{max_sins}: {sin_value}")
                                    print(f"   🎯 SIN extracted: {sin_value}")
                                else:
                                    print(f"   ⚠️  Product matched but SIN not found")
                                products = self._find_product_elements()
                            
                            i += 1
                        except Exception as e:
                            logger.warning(f"Error processing product {i+1} after scroll: {str(e)}")
                            i += 1
                            continue
                except Exception as e:
                    logger.warning(f"Error during single scroll attempt: {str(e)}")
            
            # Log final result
            if len(sins_collected) >= max_sins:
                logger.info(f"✅ SIN extraction SUCCESS: Found {len(sins_collected)}/{max_sins} SINs from {products_checked} products checked")
                print(f"   ✅ Successfully collected {len(sins_collected)}/{max_sins} SINs from {products_checked} products!")
            elif len(sins_collected) > 0:
                logger.info(f"⚠️  SIN extraction PARTIAL: Found {len(sins_collected)}/{max_sins} SINs from {products_checked} products checked")
                print(f"   ⚠️  Found {len(sins_collected)}/{max_sins} SINs from {products_checked} products (could not find more)")
            else:
                logger.info(f"❌ SIN extraction FAILED: Found 0/{max_sins} SINs from {products_checked} products checked")
                print(f"   ❌ No SINs found after checking {products_checked} products")
            
            return sins_collected
            
        except Exception as e:
            logger.error(f"Error scraping SINs from GSA page {gsa_url}: {str(e)}")
            return []
    
    def scrape_gsa_page(self, gsa_url, target_manufacturer, target_unit):
        """Scrape GSA page for product information with optimized strategy"""
        try:
            # Verify driver is ready
            if not self.driver:
                logger.error("Driver is not initialized in scrape_gsa_page!")
                return []
            
            logger.info(f"Scraping GSA page: {gsa_url}")
            
            # Navigate to the GSA page - this MUST complete and wait
            try:
                self.driver.get(gsa_url)
                logger.info(f"Page navigation initiated, waiting for page to load...")
                
                # Wait for page to be fully loaded
                # Strategy 1: Wait for document.readyState to be 'complete'
                try:
                    WebDriverWait(self.driver, 15).until(
                        lambda driver: driver.execute_script("return document.readyState") == "complete"
                    )
                    logger.info("Page readyState is 'complete'")
                except TimeoutException:
                    logger.warning("Page readyState did not become 'complete' within 15 seconds, continuing anyway")
                
                # Additional fixed wait to ensure dynamic content starts loading
                time.sleep(2)  # Give time for initial content to load
                
                # Strategy 2: Check for product elements - if found, page is loaded
                # Custom function that checks multiple selectors
                def any_product_element_present(driver):
                    selectors = [
                        (By.CSS_SELECTOR, ".productViewControl"),
                        (By.CSS_SELECTOR, "app-ux-product-display-inline"),
                        (By.CSS_SELECTOR, ".product-item"),
                        (By.CSS_SELECTOR, ".result-item"),
                        (By.CSS_SELECTOR, ".product"),
                        (By.XPATH, "//div[contains(@class, 'product')]"),
                        (By.XPATH, "//div[contains(@class, 'result')]")
                    ]
                    for selector_type, selector_value in selectors:
                        try:
                            elements = driver.find_elements(selector_type, selector_value)
                            if elements:
                                return True
                        except:
                            continue
                    return False
                
                # Wait for products to appear (up to 10 seconds)
                try:
                    WebDriverWait(self.driver, 10).until(any_product_element_present)
                    logger.info("Product elements detected - page is loaded")
                except TimeoutException:
                    # If no products found after waiting, check once more and return early
                    logger.warning("No product elements found within 10 seconds")
                    products = self._find_product_elements()
                    if not products:
                        logger.warning(f"No products found on page after waiting: {gsa_url}")
                        return []
                    logger.info(f"Found {len(products)} products on delayed check")
                
            except Exception as nav_error:
                logger.error(f"Error navigating to page {gsa_url}: {str(nav_error)}")
                return []
            
            # First, try to find products without scrolling
            products = self._find_product_elements()
            
            if not products:
                logger.warning(f"No products found on page: {gsa_url}")
                return []
            
            logger.info(f"Found {len(products)} products on initial page load")
            
            # Extract and filter products to see if we have enough matches
            initial_matches = self._extract_and_filter_products(products, target_manufacturer, target_unit)
            
            # If we have 3+ matches, return immediately (major time saver)
            if len(initial_matches) >= 3:
                logger.info(f"Found {len(initial_matches)} matching products without scrolling - proceeding")
                return initial_matches[:3]
            
            # If we have 1-2 matches, try smart scrolling (load more but not all)
            if len(initial_matches) > 0:
                logger.info(f"Found {len(initial_matches)} matching products, doing smart scrolling...")
                self._smart_scroll_to_load_more_products()
                
                # Find products again after smart scrolling
                products = self._find_product_elements()
                logger.info(f"Found {len(products)} products after smart scrolling")
                
                # Extract and filter products to get final results
                final_matches = self._extract_and_filter_products(products, target_manufacturer, target_unit)
                
                # If we now have 3+ matches, return them
                if len(final_matches) >= 3:
                    return final_matches[:3]
                else:
                    return final_matches
            else:
                # No matches found initially, do full scroll as last resort
                logger.info(f"No matching products found initially, doing full scroll...")
                self._scroll_to_load_all_products()
                
                # Find products again after full scrolling
                products = self._find_product_elements()
                logger.info(f"Found {len(products)} products after full scrolling")
                
                # Extract and filter products to get final results
                final_matches = self._extract_and_filter_products(products, target_manufacturer, target_unit)
                
                # Take top 3 matching products
                extracted_products = final_matches[:3]
                
                # If we don't have 3 matching products, log the issue
                if len(extracted_products) < 3:
                    logger.warning(f"Only found {len(extracted_products)} products matching manufacturer '{target_manufacturer}' and unit '{target_unit}'")
                
                return extracted_products
            
        except Exception as e:
            logger.error(f"Error scraping GSA page {gsa_url}: {str(e)}")
            return []
    
    def _smart_scroll_to_load_more_products(self):
        """Smart scrolling - load more products but stop early if we have enough matches"""
        try:
            # Get initial page height
            last_height = self.driver.execute_script("return document.body.scrollHeight")
            scroll_attempts = 0
            max_scroll_attempts = 5  # Reduced from 10 to 5 for smart scrolling
            
            logger.info("Smart scrolling to load more products...")
            
            while scroll_attempts < max_scroll_attempts:
                # Scroll down to bottom
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                # Wait for new content to load (reduced time)
                time.sleep(2)  # Wait for new content to load
                
                # Calculate new scroll height
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                
                # If height hasn't changed, we've reached the bottom
                if new_height == last_height:
                    break
                    
                last_height = new_height
                scroll_attempts += 1
                logger.debug(f"Smart scroll attempt {scroll_attempts}, new height: {new_height}")
            
            # Scroll back to top to ensure we can see all products
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)  # Wait for scroll to complete
            
            logger.info(f"Finished smart scrolling after {scroll_attempts} attempts")
            
        except Exception as e:
            logger.warning(f"Error during smart scrolling: {str(e)}")

    def _scroll_to_load_all_products(self):
        """Scroll down to load all products (GSA uses lazy loading) - used as last resort"""
        try:
            # Get initial page height
            last_height = self.driver.execute_script("return document.body.scrollHeight")
            scroll_attempts = 0
            max_scroll_attempts = 8  # Reduced from 10 to 8
            
            logger.info("Full scrolling to load all products...")
            
            while scroll_attempts < max_scroll_attempts:
                # Scroll down to bottom
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                # Wait for new content to load
                time.sleep(2)  # Wait for new content to load
                
                # Calculate new scroll height
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                
                # If height hasn't changed, we've reached the bottom
                if new_height == last_height:
                    break
                    
                last_height = new_height
                scroll_attempts += 1
                logger.debug(f"Full scroll attempt {scroll_attempts}, new height: {new_height}")
            
            # Scroll back to top to ensure we can see all products
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)  # Wait for scroll to complete
            
            logger.info(f"Finished full scrolling after {scroll_attempts} attempts")
            
        except Exception as e:
            logger.warning(f"Error during full scrolling: {str(e)}")

    def _extract_and_filter_products(self, products, target_manufacturer, target_unit):
        """Extract product info and filter by manufacturer + unit match"""
        # Check if first product is header text
        start_index = 0
        if len(products) > 0:
            first_product_text = products[0].text.lower()
            if any(header_indicator in first_product_text for header_indicator in [
                'name contract number price', 'contractor name', 'price low to high', 
                'view as grid', 'sort by', 'filter by'
            ]):
                start_index = 1
                logger.info("Skipping first product as it appears to be header text")
        
        # Extract ALL products and filter by manufacturer + unit match
        all_products_info = []
        for i in range(start_index, len(products)):
            try:
                product_info = self._extract_product_info(products[i], i+1, target_manufacturer, target_unit)
                if product_info and (product_info.get('price') is not None or product_info.get('contractor') is not None):
                    # Additional check to skip header-like products
                    if product_info.get('contractor') and any(header_word in product_info.get('contractor', '').lower() for header_word in [
                        'name contract', 'price low', 'view as', 'sort by'
                    ]):
                        logger.info(f"Skipping product {i+1} as it appears to be header text")
                        continue
                    
                    all_products_info.append(product_info)
            except Exception as e:
                logger.warning(f"Error extracting info from product {i+1}: {str(e)}")
        
        # Filter products by manufacturer and unit match
        matching_products = []
        rejected_products = []
        for product in all_products_info:
            manufacturer_match = product.get('manufacturer_match', False)
            unit_match = product.get('unit_match', False)
            
            if manufacturer_match and unit_match:
                matching_products.append(product)
                logger.info(f"MATCHED Product {product['product_num']}: Price={product['price']}, Contractor={product['contractor']}, Manufacturer={product.get('website_manufacturer')}, Unit={product.get('website_unit')}")
            else:
                # Log why product was rejected
                rejection_reason = []
                if not manufacturer_match:
                    rejection_reason.append(f"Manufacturer mismatch (target: '{target_manufacturer}', found: '{product.get('website_manufacturer')}')")
                if not unit_match:
                    rejection_reason.append(f"Unit mismatch (target: '{target_unit}', found: '{product.get('website_unit')}')")
                
                rejected_products.append({
                    'product_num': product.get('product_num'),
                    'price': product.get('price'),
                    'contractor': product.get('contractor'),
                    'reason': ' | '.join(rejection_reason)
                })
                logger.debug(f"REJECTED Product {product.get('product_num')}: {' | '.join(rejection_reason)}")
        
        # Log summary of rejected products if any
        if rejected_products and len(matching_products) == 0:
            logger.warning(f"Found {len(rejected_products)} products but all were rejected due to matching criteria:")
            for rejected in rejected_products[:5]:  # Show first 5
                logger.warning(f"  - Product {rejected['product_num']}: {rejected['reason']}")
            if len(rejected_products) > 5:
                logger.warning(f"  ... and {len(rejected_products) - 5} more rejected products")
        
        return matching_products

    def _find_product_elements(self):
        """Find product elements on the GSA page"""
        product_selectors = [
            (By.CSS_SELECTOR, ".productViewControl"),  # Primary selector for GSA products
            (By.CSS_SELECTOR, "app-ux-product-display-inline"),  # Alternative selector
            (By.CSS_SELECTOR, ".product-item"),
            (By.CSS_SELECTOR, ".result-item"),
            (By.CSS_SELECTOR, ".product"),
            (By.CSS_SELECTOR, "[class*='product']"),
            (By.CSS_SELECTOR, "[class*='result']"),
            (By.XPATH, "//div[contains(@class, 'product')]"),
            (By.XPATH, "//div[contains(@class, 'result')]"),
            (By.XPATH, "//div[contains(@class, 'item')]"),
            (By.XPATH, "//tr[contains(@class, 'product')]"),
            (By.XPATH, "//div[contains(@class, 'row')]")
        ]
        
        for selector_type, selector_value in product_selectors:
            try:
                products = self.driver.find_elements(selector_type, selector_value)
                if products:
                    logger.info(f"Found {len(products)} products using selector: {selector_type} = {selector_value}")
                    return products
            except:
                continue
        
        logger.warning("No product elements found with any selector")
        return []
    
    def _extract_product_info(self, product_element, product_num, target_manufacturer, target_unit):
        """Extract price, contractor, and contract information from a product element"""
        try:
            product_text = product_element.text.lower()
            
            # Extract price
            price = self._extract_price(product_text)
            
            # Extract contractor
            contractor = self._extract_contractor(product_text)
            
            # Extract contract number
            contract = self._extract_contract(product_text)
            
            # Extract manufacturer and unit for matching
            website_manufacturer = self._extract_manufacturer(product_text)
            website_unit = self._extract_unit(product_text)
            
            # Check if manufacturer and unit match
            manufacturer_match = self.fuzzy_match_manufacturer(target_manufacturer, website_manufacturer)
            unit_match = self.fuzzy_match_unit(target_unit, website_unit)
            
            logger.debug(f"Product {product_num}: Manufacturer match={manufacturer_match}, Unit match={unit_match}")
            
            return {
                'product_num': product_num,
                'price': price,
                'contractor': contractor,
                'contract': contract,
                'manufacturer_match': manufacturer_match,
                'unit_match': unit_match,
                'website_manufacturer': website_manufacturer,
                'website_unit': website_unit,
                'raw_text': product_element.text[:200] + '...' if len(product_element.text) > 200 else product_element.text
            }
            
        except Exception as e:
            logger.error(f"Error extracting product info: {str(e)}")
            return None
    
    def _extract_price(self, text):
        """Extract price from product text using pre-compiled patterns"""
        for pattern in self._price_patterns:
            matches = pattern.findall(text)
            if matches:
                price_str = matches[0].replace(',', '').strip()
                try:
                    return float(price_str)
                except:
                    continue
        
        return None
    
    def _extract_contractor(self, text):
        """Extract contractor name from product text using pre-compiled patterns"""
        for pattern in self._contractor_patterns:
            matches = pattern.findall(text)
            if matches:
                contractor = matches[0].strip()
                # Clean up the contractor name
                contractor = re.sub(r'\s+', ' ', contractor)
                # Remove unwanted suffixes
                contractor = re.sub(r'\s+contract\s*$', '', contractor, flags=re.IGNORECASE)
                contractor = re.sub(r'\s+includes\s*$', '', contractor, flags=re.IGNORECASE)
                contractor = re.sub(r'\s+inc\.?\s*$', ' Inc.', contractor, flags=re.IGNORECASE)
                contractor = re.sub(r'\s+llc\s*$', ' LLC', contractor, flags=re.IGNORECASE)
                contractor = re.sub(r'\s+corp\.?\s*$', ' Corp.', contractor, flags=re.IGNORECASE)
                return contractor.title()
        
        return None
    
    def _extract_contract(self, text):
        """Extract contract number from product text using pre-compiled patterns"""
        for pattern in self._contract_patterns:
            matches = pattern.findall(text)
            if matches:
                contract = matches[0].strip().upper()
                # Filter out common false positives
                if contract not in ['OR', 'AND', 'THE', 'TO', 'OF', 'IN', 'ON', 'AT', 'BY', 'FOR']:
                    return contract
        return None
    
    def _extract_manufacturer(self, text):
        """Extract manufacturer name from product text using pre-compiled patterns.

        Preference order:
        1) mfr:
        2) manufacturer:
        3) mfg:
        4) brand:
        """
        for pattern in self._manufacturer_patterns:
            m = pattern.search(text)
            if m:
                value = m.group(1).strip()
                value = re.sub(r'\s+', ' ', value)
                return value
        return None
    
    def _extract_unit(self, text):
        """Extract unit of measure from product text using pre-compiled patterns"""
        for pattern in self._unit_patterns:
            matches = pattern.findall(text)
            if matches:
                return matches[0].strip().upper()
        
        return None
    
    def row_has_two_sins(self, row):
        """Check if row already has at least 2 SINs filled OR contains 'SIN not found' (already attempted)"""
        # First check if any column has "SIN not found" - if so, skip this row
        for col in ['SIN1', 'SIN2', 'SIN3']:
            if col in row.index:
                value = row[col]
                if pd.notna(value) and str(value).strip().lower() == 'sin not found':
                    return True  # Skip rows with "SIN not found"
        
        # Otherwise, check if row has at least 2 valid SINs
        sin_count = 0
        for col in ['SIN1', 'SIN2', 'SIN3']:
            if col in row.index:
                value = row[col]
                if pd.notna(value) and str(value).strip() != '' and str(value).strip().lower() not in ['nan', 'sin not found']:
                    sin_count += 1
        return sin_count >= 2
    
    def extract_sin_from_product_page(self, product_url, max_attempts=2):
        """Navigate to product detail page and extract SIN number - FAST & SIMPLE METHOD"""
        for attempt in range(max_attempts):
            try:
                logger.info(f"Navigating to product page (attempt {attempt + 1}/{max_attempts}): {product_url}")
                
                # Navigate and wait for page to load
                self.driver.get(product_url)
                
                # Wait for page to be fully loaded
                try:
                    WebDriverWait(self.driver, 10).until(
                        lambda driver: driver.execute_script("return document.readyState") == "complete"
                    )
                    logger.info("Product page readyState is 'complete'")
                except:
                    logger.warning("Product page readyState timeout, continuing anyway")
                
                # Initial wait for page structure
                time.sleep(3.0)
                
                # CRITICAL: Scroll to trigger lazy-loaded content
                logger.info("Scrolling to trigger lazy-loaded product details...")
                try:
                    # Scroll to middle of page
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/3);")
                    time.sleep(2.0)
                    
                    # Scroll to bottom
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2.0)
                    
                    # Scroll back to top where product details are
                    self.driver.execute_script("window.scrollTo(0, 0);")
                    time.sleep(2.0)
                except Exception as scroll_err:
                    logger.warning(f"Scrolling error: {str(scroll_err)}")
                
                # Wait for product details to be present (not just header/footer)
                try:
                    # Wait for tables or product detail content to appear
                    WebDriverWait(self.driver, 15).until(
                        lambda d: len(d.find_element(By.TAG_NAME, "body").text) > 2000
                    )
                    logger.info("Product details content loaded (page has substantial text)")
                except:
                    logger.warning("Product details may not be fully loaded (page has minimal content)")
                
                # Final wait to ensure everything is rendered
                time.sleep(3.0)
                
                # Get page text
                page_text = self.driver.find_element(By.TAG_NAME, "body").text
                logger.info(f"Product page loaded with {len(page_text)} characters of text")
                
                # Check if page has actual product details (not just header/footer)
                if len(page_text) < 1500:
                    logger.warning(f"Product page has minimal content ({len(page_text)} chars) - likely only header/footer loaded")
                    if attempt < max_attempts - 1:
                        logger.info("Retrying page load due to insufficient content...")
                        time.sleep(3)
                        continue
                    else:
                        logger.error("Product details failed to load after all attempts")
                        return None
                
                # Check if page contains product detail indicators
                if 'schedule' not in page_text.lower() and 'sin' not in page_text.lower():
                    logger.warning("Product page does not contain 'Schedule' or 'SIN' keywords - content may not be loaded")
                    if attempt < max_attempts - 1:
                        logger.info("Retrying page load due to missing product details...")
                        time.sleep(3)
                        continue
                
                # STRATEGY 1: Simple text search with regex (FAST!)
                patterns = [
                    r'Schedule/SIN[:\s]+([A-Z0-9]+)/([A-Z0-9]+)',  # "Schedule/SIN: MAS/511210"
                    r'Schedule/SIN[:\s]+([A-Z0-9]+)',               # "Schedule/SIN: 511210"
                    r'SIN[:\s]+([A-Z0-9]+)/([A-Z0-9]+)',            # "SIN: MAS/511210"
                    r'SIN[:\s]+([A-Z0-9]+)',                        # "SIN: 511210"
                ]
                
                for pattern in patterns:
                    matches = re.search(pattern, page_text, re.IGNORECASE)
                    if matches:
                        # If two groups (MAS/511210), take second one
                        if len(matches.groups()) >= 2 and matches.group(2):
                            sin_number = matches.group(2).upper()
                            logger.info(f"✅ Found SIN: {sin_number}")
                            print(f"      ✅ Found SIN: {sin_number}")
                            return sin_number
                        # If one group
                        elif len(matches.groups()) == 1:
                            full_sin = matches.group(1)
                            if '/' in full_sin:
                                sin_number = full_sin.split('/')[-1].strip().upper()
                            else:
                                sin_number = full_sin.strip().upper()
                            logger.info(f"✅ Found SIN: {sin_number}")
                            print(f"      ✅ Found SIN: {sin_number}")
                            return sin_number
                
                # STRATEGY 2: Look in tables (backup method)
                tables = self.driver.find_elements(By.TAG_NAME, "table")
                for table in tables:
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    for row in rows:
                        row_text = row.text.lower()
                        if 'schedule/sin' in row_text or 'schedule sin' in row_text:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            for cell in cells:
                                cell_text = cell.text.strip()
                                if '/' in cell_text:
                                    parts = cell_text.split('/')
                                    if len(parts) >= 2:
                                        sin_number = parts[-1].strip().upper()
                                        if re.match(r'^[A-Z0-9]+$', sin_number, re.IGNORECASE):
                                            logger.info(f"✅ Found SIN in table: {sin_number}")
                                            print(f"      ✅ Found SIN in table: {sin_number}")
                                            return sin_number
                
                # Retry if needed
                if attempt < max_attempts - 1:
                    logger.info(f"Retrying (attempt {attempt + 2}/{max_attempts})...")
                    time.sleep(1)
                    continue
                
                # No SIN found
                logger.warning(f"No SIN found: {product_url}")
                print(f"      ❌ Could not extract SIN")
                return None
                    
            except Exception as e:
                logger.error(f"Error extracting SIN (attempt {attempt + 1}): {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(1)
                    continue
                return None
        
        return None
    
    def click_product_and_extract_sin(self, product_element, product_num):
        """Click on product name to navigate to detail page and extract SIN"""
        try:
            # Find clickable product link within the product element
            # Try multiple selectors for product name/link
            link_selectors = [
                (By.CSS_SELECTOR, "a.product-link"),
                (By.CSS_SELECTOR, "a[href*='product_detail']"),
                (By.CSS_SELECTOR, "h3 a"),
                (By.CSS_SELECTOR, "h4 a"),
                (By.CSS_SELECTOR, ".product-title a"),
                (By.CSS_SELECTOR, "a"),
                (By.XPATH, ".//a[contains(@href, 'product_detail')]"),
                (By.XPATH, ".//a[contains(@class, 'product')]"),
            ]
            
            product_link = None
            product_url = None
            
            for selector_type, selector_value in link_selectors:
                try:
                    links = product_element.find_elements(selector_type, selector_value)
                    for link in links:
                        href = link.get_attribute('href')
                        if href and 'product_detail' in href:
                            product_link = link
                            product_url = href
                            break
                    if product_link:
                        break
                except:
                    continue
            
            if not product_link or not product_url:
                logger.warning(f"Product {product_num}: No product detail link found")
                return None
            
            logger.info(f"Product {product_num}: Found detail link: {product_url}")
            
            # Extract SIN from product detail page
            sin_value = self.extract_sin_from_product_page(product_url)
            
            # Navigate back to search results page
            self.driver.back()
            # Increased wait to prevent "Unexpected Error" from GSA when returning to search page
            time.sleep(4.0)  # Stay longer on search page after returning to prevent rate limiting
            
            return sin_value
            
        except Exception as e:
            logger.error(f"Error clicking product and extracting SIN: {str(e)}")
            # Try to navigate back in case we're stuck
            try:
                self.driver.back()
                time.sleep(2)
            except:
                pass
            return None
    
    def update_dataframe_with_results(self, df, row_idx, products_data):
        """Update dataframe with scraped product information"""
        try:
            # Update columns for up to 3 products
            for i, product in enumerate(products_data[:3]):
                if i == 0:
                    # First product goes to columns without suffix
                    df.at[row_idx, 'GSA PRICE'] = product.get('price', '')
                    df.at[row_idx, 'Contractor'] = product.get('contractor', '')
                    df.at[row_idx, 'contract#:'] = product.get('contract', '')
                elif i == 1:
                    # Second product goes to .1 columns
                    df.at[row_idx, 'GSA PRICE.1'] = product.get('price', '')
                    df.at[row_idx, 'Contractor.1'] = product.get('contractor', '')
                    df.at[row_idx, 'contract#:.1'] = product.get('contract', '')
                elif i == 2:
                    # Third product goes to .2 columns
                    df.at[row_idx, 'GSA PRICE.2'] = product.get('price', '')
                    df.at[row_idx, 'Contractor.2'] = product.get('contractor', '')
                    df.at[row_idx, 'contract#:.2'] = product.get('contract', '')
            
            logger.info(f"Updated dataframe row {row_idx} with {len(products_data)} products")
            
        except Exception as e:
            logger.error(f"Error updating dataframe row {row_idx}: {str(e)}")
    
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
            return backup_path
        except Exception as e:
            logger.error(f"Error creating backup: {str(e)}")
            return None
    
    def cleanup_old_backups(self, output_file, keep_last=5):
        """Clean up old backup files in backups directory, keeping only the most recent ones"""
        try:
            # Get directory and filename
            file_dir = os.path.dirname(output_file) if os.path.dirname(output_file) else '.'
            file_name = os.path.basename(output_file)
            
            # Backups folder
            backups_dir = os.path.join(file_dir, 'backups')
            if not os.path.exists(backups_dir):
                return  # No backups folder yet
            
            # Find backup files in backups directory
            backup_files = [f for f in os.listdir(backups_dir) if f.startswith(f"{file_name}.backup_")]
            backup_files.sort(reverse=True)  # Sort by name (timestamp) descending
            
            # Keep only the most recent backups
            files_to_delete = backup_files[keep_last:]
            
            for backup_file in files_to_delete:
                try:
                    backup_path = os.path.join(backups_dir, backup_file)
                    os.remove(backup_path)
                    logger.info(f"Cleaned up old backup: {backup_file}")
                except Exception as e:
                    logger.warning(f"Could not delete backup {backup_file}: {str(e)}")
                    
        except Exception as e:
            logger.warning(f"Error during backup cleanup: {str(e)}")
    
    def save_results_to_excel(self, df):
        """Save the updated dataframe to Excel file with backup"""
        # Use the same file that was loaded (self.excel_file_path)
        output_file = self.excel_file_path
        
        try:
            # Create backup if file exists
            if os.path.exists(output_file):
                self.create_backup(output_file)
            
            # Save to Excel
            df.to_excel(output_file, index=False)
            logger.info(f"Results saved to {output_file}")
            
            # Clean up old backup files (keep only the most recent 5)
            self.cleanup_old_backups(output_file)
            return True
            
        except Exception as e:
            logger.error(f"Error saving results to Excel: {str(e)}")
            return False
    
    def run_scraping_automation(self):
        """Main method to run the scraping automation"""
        try:
            # Load manufacturer mapping
            if not self.load_manufacturer_mapping():
                logger.error("Failed to load manufacturer mapping")
                return False
            
            # Read Excel data
            df, column_mapping = self.read_excel_data()
            if df is None:
                logger.error("No data found in Excel file")
                return False
            
            # Setup web driver
            self.setup_driver()
            
            successful_scrapes = 0
            start_time = time.time()
            
            # Process each row
            for i, row in df.iterrows():
                try:
                    gsa_url = row[column_mapping['links']]
                    manufacturer = row[column_mapping['manufacturer']]
                    unit_of_measure = row[column_mapping['unit_of_measure']]
                    stock_number = row[column_mapping['stock_number']]
                    
                    if pd.isna(gsa_url) or not gsa_url.strip():
                        logger.warning(f"Row {i+1}: No GSA URL found for stock number {stock_number}")
                        continue
                    
                    print(f"\nProgress: {i+1}/{len(df)} ({((i+1)/len(df)*100):.1f}%) - Processing: {stock_number}")
                    logger.info(f"Processing row {i+1}/{len(df)}: {stock_number}")
                    
                    # Start timing for this product
                    product_start_time = time.time()
                    
                    # Scrape the GSA page
                    products_data = self.scrape_gsa_page(gsa_url, manufacturer, unit_of_measure)
                    
                    # Calculate timing
                    product_time = time.time() - product_start_time
                    
                    if products_data:
                        successful_scrapes += 1
                        self.update_dataframe_with_results(df, i, products_data)
                        print(f"SUCCESS! Found {len(products_data)} products - Row {i+1} completed ({product_time:.1f}s)")
                        logger.info(f"Successfully scraped {len(products_data)} products for row {i+1} in {product_time:.1f}s")
                    else:
                        print(f"WARNING: No products found for: {stock_number} - Row {i+1} ({product_time:.1f}s)")
                        logger.warning(f"No products found for row {i+1}: {stock_number} in {product_time:.1f}s")
                    
                    # Calculate ETA
                    elapsed_time = time.time() - start_time
                    avg_time_per_product = elapsed_time / (i + 1)
                    remaining_products = len(df) - (i + 1)
                    eta_seconds = remaining_products * avg_time_per_product
                    eta_hours = eta_seconds / 3600
                    
                    print(f"Timing: {product_time:.1f}s | Avg: {avg_time_per_product:.1f}s/product | ETA: {eta_hours:.1f}h")
                    
                    # Save results every 100 rows
                    if (i + 1) % 100 == 0:
                        self.save_results_to_excel(df)
                        print(f"Progress saved at row {i+1}")
                    
                    # Rate limiting - wait between requests (reduced from 2 to 1 second)
                    time.sleep(2)  # Rate limiting - wait between requests
                    
                except Exception as e:
                    logger.error(f"Error processing row {i+1}: {str(e)}")
                    continue
            
            # Final save
            self.save_results_to_excel(df)
            
            # Calculate final statistics
            total_time = time.time() - start_time
            
            logger.info(f"Scraping automation completed!")
            logger.info(f"Processed: {len(df)} products")
            logger.info(f"Successful scrapes: {successful_scrapes}")
            logger.info(f"Total time: {total_time:.1f} seconds ({total_time/60:.1f} minutes)")
            
            return True
            
        except Exception as e:
            logger.error(f"Error in scraping automation: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()

    def run_scraping_single(self, stock_number_query: str) -> bool:
        """Run scraping for exactly one product identified by Item Stock Number-Butted."""
        try:
            # Load manufacturer mapping
            if not self.load_manufacturer_mapping():
                logger.error("Failed to load manufacturer mapping")
                return False

            # Read Excel data
            df, column_mapping = self.read_excel_data()
            if df is None:
                logger.error("No data found in Excel file")
                return False

            stock_col = column_mapping['stock_number']
            # Find exact matching row (string compare, strip)
            mask = df[stock_col].astype(str).str.strip().str.lower() == str(stock_number_query).strip().lower()
            matches = df[mask]
            if matches.empty:
                print(f"ERROR: No product found with Item Stock Number-Butted: {stock_number_query}")
                logger.error(f"Single-run: No match for stock number '{stock_number_query}'")
                return False

            # Use the first exact match
            row_idx = matches.index[0]
            gsa_url = df.at[row_idx, column_mapping['links']]
            manufacturer = df.at[row_idx, column_mapping['manufacturer']]
            unit_of_measure = df.at[row_idx, column_mapping['unit_of_measure']]
            stock_number = df.at[row_idx, column_mapping['stock_number']]

            if pd.isna(gsa_url) or not str(gsa_url).strip():
                print(f"ERROR: No GSA URL for '{stock_number}'")
                logger.error(f"Single-run: Missing GSA URL for '{stock_number}'")
                return False

            print(f"Processing single product: {stock_number}")
            logger.info(f"Single-run: Processing row {row_idx+1}: {stock_number}")

            # Setup web driver
            self.setup_driver()

            # Start timing for this product
            product_start_time = time.time()

            # Scrape and update
            products_data = self.scrape_gsa_page(gsa_url, manufacturer, unit_of_measure)
            
            # Calculate timing
            product_time = time.time() - product_start_time
            if products_data:
                self.update_dataframe_with_results(df, row_idx, products_data)
                self.save_results_to_excel(df)
                print(f"SUCCESS! Found {len(products_data)} products - Saved for {stock_number} ({product_time:.1f}s)")
                logger.info(f"Single-run: Successfully scraped {len(products_data)} products for '{stock_number}' in {product_time:.1f}s")
                return True
            else:
                print(f"WARNING: No matching products for: {stock_number} ({product_time:.1f}s)")
                logger.warning(f"Single-run: No matching products for '{stock_number}' in {product_time:.1f}s")
                return False

        except Exception as e:
            logger.error(f"Error in single product mode: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()
    
    def run_scraping_full(self):
        """Convenience wrapper used by the menu for full automation."""
        return self.run_scraping_automation()

    def run_scraping_custom_range(self, start_row: int, end_row: int) -> bool:
        """Run scraping for a specific 0-based row range [start_row, end_row]."""
        try:
            # Load manufacturer mapping
            if not self.load_manufacturer_mapping():
                logger.error("Failed to load manufacturer mapping")
                return False

            # Read Excel data
            df, column_mapping = self.read_excel_data()
            if df is None:
                logger.error("No data found in Excel file")
                return False

            # Clamp indices to valid bounds
            start_row = max(0, start_row)
            end_row = min(len(df) - 1, end_row)
            if start_row > end_row:
                logger.error(f"Invalid range: {start_row}-{end_row}")
                return False

            # Setup web driver
            self.setup_driver()

            successful_scrapes = 0
            start_time = time.time()

            total = end_row - start_row + 1
            for offset, i in enumerate(range(start_row, end_row + 1), 1):
                try:
                    gsa_url = df.at[i, column_mapping['links']]
                    manufacturer = df.at[i, column_mapping['manufacturer']]
                    unit_of_measure = df.at[i, column_mapping['unit_of_measure']]
                    stock_number = df.at[i, column_mapping['stock_number']]

                    if pd.isna(gsa_url) or not str(gsa_url).strip():
                        logger.warning(f"Row {i+1}: No GSA URL found for stock number {stock_number}")
                        continue

                    print(
                        f"Progress: {offset}/{total} (Row {i+1}) - Processing: {stock_number}"
                    )
                    logger.info(f"Processing row {i+1}: {stock_number}")

                    # Start timing for this product
                    product_start_time = time.time()

                    # Scrape and filter by manufacturer + unit inside scrape_gsa_page
                    products_data = self.scrape_gsa_page(gsa_url, manufacturer, unit_of_measure)
                    
                    # Calculate timing
                    product_time = time.time() - product_start_time

                    if products_data:
                        successful_scrapes += 1
                        self.update_dataframe_with_results(df, i, products_data)
                        print(
                            f"SUCCESS! Found {len(products_data)} products - Row {i+1} completed ({product_time:.1f}s)"
                        )
                        logger.info(
                            f"Successfully scraped {len(products_data)} products for row {i+1} in {product_time:.1f}s"
                        )
                    else:
                        print(f"WARNING: No matching products for: {stock_number} (Row {i+1}) ({product_time:.1f}s)")
                        logger.warning(
                            f"No matching products for row {i+1}: {stock_number} in {product_time:.1f}s"
                        )

                    # Calculate ETA for custom range
                    elapsed_time = time.time() - start_time
                    avg_time_per_product = elapsed_time / offset
                    remaining_products = total - offset
                    eta_seconds = remaining_products * avg_time_per_product
                    eta_hours = eta_seconds / 3600
                    
                    print(f"Timing: {product_time:.1f}s | Avg: {avg_time_per_product:.1f}s/product | ETA: {eta_hours:.1f}h")

                    # Save periodically inside ranges as well
                    if offset % 100 == 0:
                        self.save_results_to_excel(df)
                        print(f"Progress saved at row {i+1}")

                    time.sleep(2)  # Rate limiting - wait between requests

                except Exception as e:
                    logger.error(f"Error processing row {i+1}: {str(e)}")
                    continue

            # Final save
            self.save_results_to_excel(df)

            total_time = time.time() - start_time
            logger.info("Custom range scraping completed!")
            logger.info(f"Rows: {start_row+1}-{end_row+1} ({total} items)")
            logger.info(f"Successful scrapes: {successful_scrapes}")
            logger.info(
                f"Total time: {total_time:.1f} seconds ({total_time/60:.1f} minutes)"
            )
            return True

        except Exception as e:
            logger.error(f"Error in custom range scraping: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()
    
    def run_scraping_test_mode(self, test_count=1):  # Test with just 1 product
        """Test method to run scraping with limited items"""
        try:
            # Load manufacturer mapping
            if not self.load_manufacturer_mapping():
                logger.error("Failed to load manufacturer mapping")
                return False
            
            # Read Excel data
            df, column_mapping = self.read_excel_data()
            if df is None:
                logger.error("No data found in Excel file")
                return False
            
            # Take only first few rows for testing
            test_df = df.head(test_count)
            logger.info(f"Test mode: Processing {len(test_df)} products")
            
            # Setup web driver
            self.setup_driver()
            
            successful_scrapes = 0
            start_time = time.time()
            
            # Process test rows
            for i, row in test_df.iterrows():
                try:
                    gsa_url = row[column_mapping['links']]
                    manufacturer = row[column_mapping['manufacturer']]
                    unit_of_measure = row[column_mapping['unit_of_measure']]
                    stock_number = row[column_mapping['stock_number']]
                    
                    if pd.isna(gsa_url) or not gsa_url.strip():
                        logger.warning(f"Row {i+1}: No GSA URL found for stock number {stock_number}")
                        continue
                    
                    print(f"\nTest Progress: {i+1}/{len(test_df)} - Processing: {stock_number}")
                    logger.info(f"Test processing row {i+1}/{len(test_df)}: {stock_number}")
                    
                    # Start timing for this product
                    product_start_time = time.time()
                    
                    # Scrape the GSA page
                    products_data = self.scrape_gsa_page(gsa_url, manufacturer, unit_of_measure)
                    
                    # Calculate timing
                    product_time = time.time() - product_start_time
                    
                    if products_data:
                        successful_scrapes += 1
                        self.update_dataframe_with_results(df, i, products_data)
                        print(f"SUCCESS! Found {len(products_data)} products ({product_time:.1f}s)")
                        logger.info(f"Successfully scraped {len(products_data)} products for test row {i+1} in {product_time:.1f}s")
                    else:
                        print(f"WARNING: No products found for: {stock_number} ({product_time:.1f}s)")
                        logger.warning(f"No products found for test row {i+1}: {stock_number} in {product_time:.1f}s")
                    
                    print(f"Timing: {product_time:.1f}s")
                    
                    # Save results every 100 products (for test mode with large test counts)
                    if (i + 1) % 100 == 0:
                        self.save_results_to_excel(df)
                        print(f"Progress saved at row {i+1}")
                    
                    # Rate limiting (reduced from 3 to 1.5 seconds)
                    time.sleep(3)  # Rate limiting
                    
                except Exception as e:
                    logger.error(f"Error processing test row {i+1}: {str(e)}")
                    continue
            
            # Save results
            self.save_results_to_excel(df)
            
            # Calculate final statistics
            total_time = time.time() - start_time
            
            print(f"\nTest completed!")
            print(f"Processed: {len(test_df)} products")
            print(f"Successful scrapes: {successful_scrapes}")
            print(f"Total time: {total_time:.2f} seconds")
            
            logger.info(f"Test scraping completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error in test scraping: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()

    def identify_missing_rows(self, df):
        """Identify rows where GSA data is missing or incomplete"""
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
    
    def run_sin_scraping_menu(self):
        """Interactive menu for SIN scraping"""
        scrapped_products_file = "../ScrappedProducts.xlsx"
        
        # Check if file exists
        if not os.path.exists(scrapped_products_file):
            print(f"\n❌ ERROR: {scrapped_products_file} not found!")
            print(f"   Looking in: {os.path.abspath(scrapped_products_file)}")
            return False
        
        while True:
            print("\n" + "="*80)
            print("🎯 SIN SCRAPING MENU")
            print("="*80)
            print("📁 Target file: ScrappedProducts.xlsx (parent folder)")
            print("✅ Smart Logic: Skips rows with 2+ SINs, completes partial rows")
            print("="*80)
            print("1. Test Mode (First 10 products)")
            print("2. Custom Range (Specify start and end row)")
            print("3. Single Product (By Item Number)")
            print("4. Resume from Row (Skip already processed)")
            print("5. Full Automation (All products)")
            print("6. Back to Main Menu")
            print("="*80)
            
            try:
                choice = input("Enter your choice (1-6): ").strip()
                
                if choice == "1":
                    print("\n" + "="*60)
                    print("🧪 TEST MODE - SIN SCRAPING")
                    print("="*60)
                    print("Processing first 10 products...")
                    success = self.run_sin_scraping_range(0, 9)
                    if success:
                        print("\n✅ Test mode completed!")
                    else:
                        print("\n❌ Test mode failed!")
                
                elif choice == "2":
                    print("\n" + "="*60)
                    print("📏 CUSTOM RANGE MODE - SIN SCRAPING")
                    print("="*60)
                    
                    # Get total rows
                    try:
                        df = pd.read_excel(scrapped_products_file)
                        total_rows = len(df)
                        print(f"Total products available: {total_rows}")
                        
                        start_row = int(input(f"Enter start row (1-{total_rows}): ")) - 1
                        end_row = int(input(f"Enter end row ({start_row + 1}-{total_rows}): ")) - 1
                        
                        if start_row < 0 or end_row >= total_rows or start_row > end_row:
                            print("❌ ERROR: Invalid range!")
                            continue
                        
                        count = end_row - start_row + 1
                        print(f"\n📊 Processing rows {start_row + 1}-{end_row + 1} ({count} products)...")
                        
                        success = self.run_sin_scraping_range(start_row, end_row)
                        if success:
                            print(f"\n✅ Custom range completed!")
                        else:
                            print(f"\n❌ Custom range failed!")
                    
                    except ValueError:
                        print("❌ ERROR: Please enter valid numbers!")
                        continue
                
                elif choice == "3":
                    print("\n" + "="*60)
                    print("🎯 SINGLE PRODUCT MODE - SIN SCRAPING")
                    print("="*60)
                    item_number = input("Enter Item Number: ").strip()
                    
                    if not item_number:
                        print("❌ ERROR: Item number cannot be empty!")
                        continue
                    
                    success = self.run_sin_scraping_single(item_number)
                    if success:
                        print(f"\n✅ Single product SIN scraping completed!")
                    else:
                        print(f"\n❌ Single product SIN scraping failed!")
                
                elif choice == "4":
                    print("\n" + "="*60)
                    print("▶️  RESUME MODE - SIN SCRAPING")
                    print("="*60)
                    
                    try:
                        df = pd.read_excel(scrapped_products_file)
                        total_rows = len(df)
                        print(f"Total products: {total_rows}")
                        
                        start_row = int(input(f"Resume from row (1-{total_rows}): ")) - 1
                        
                        if start_row < 0 or start_row >= total_rows:
                            print("❌ ERROR: Invalid row number!")
                            continue
                        
                        end_row = total_rows - 1
                        count = end_row - start_row + 1
                        print(f"\n📊 Processing rows {start_row + 1}-{end_row + 1} ({count} products)...")
                        
                        success = self.run_sin_scraping_range(start_row, end_row)
                        if success:
                            print(f"\n✅ Resume mode completed!")
                        else:
                            print(f"\n❌ Resume mode failed!")
                    
                    except ValueError:
                        print("❌ ERROR: Please enter a valid number!")
                        continue
                
                elif choice == "5":
                    print("\n" + "="*60)
                    print("🚀 FULL AUTOMATION MODE - SIN SCRAPING")
                    print("="*60)
                    
                    try:
                        df = pd.read_excel(scrapped_products_file)
                        total_rows = len(df)
                        
                        print(f"⚠️  WARNING: This will process ALL {total_rows} products!")
                        print(f"⏱️  Estimated time: ~{total_rows * 15 / 3600:.1f} hours")
                        print("📊 Products with 2+ SINs will be skipped")
                        
                        confirm = input("\nContinue? (yes/no): ").strip().lower()
                        if confirm not in ['yes', 'y']:
                            print("❌ Full automation cancelled.")
                            continue
                        
                        success = self.run_sin_scraping_range(0, total_rows - 1)
                        if success:
                            print(f"\n✅ Full automation completed!")
                        else:
                            print(f"\n❌ Full automation failed!")
                    
                    except Exception as e:
                        print(f"❌ ERROR: {str(e)}")
                        continue
                
                elif choice == "6":
                    print("\nReturning to main menu...")
                    return True
                
                else:
                    print("❌ ERROR: Invalid choice! Please enter 1-6.")
            
            except KeyboardInterrupt:
                print("\n\n⚠️  Operation cancelled by user.")
                return True
            except Exception as e:
                print(f"\n❌ ERROR: {str(e)}")
                continue
    
    def run_sin_scraping_single(self, item_number):
        """Scrape SIN for a single product by Item Number"""
        try:
            scrapped_products_file = "../ScrappedProducts.xlsx"
            
            print(f"\n🔍 Searching for Item Number: {item_number}")
            
            # Load file
            df = pd.read_excel(scrapped_products_file)
            
            # Find product
            mask = df['Item Number'].astype(str).str.strip().str.lower() == str(item_number).strip().lower()
            matches = df[mask]
            
            if matches.empty:
                print(f"❌ ERROR: No product found with Item Number: {item_number}")
                return False
            
            # Get row index
            row_idx = matches.index[0]
            print(f"✅ Found at Row {row_idx + 1}")
            
            # Process single row
            return self.run_sin_scraping_range(row_idx, row_idx)
        
        except Exception as e:
            logger.error(f"Error in single product SIN scraping: {str(e)}")
            print(f"❌ ERROR: {str(e)}")
            return False
    
    def run_sin_scraping_range(self, start_row, end_row):
        """Scrape SINs for a specific range of rows"""
        try:
            scrapped_products_file = "../ScrappedProducts.xlsx"
            
            print("\n" + "="*80)
            print("🎯 SIN NUMBER SCRAPING")
            print("="*80)
            print("📋 Scraping SIN numbers from GSA product detail pages")
            print("📁 Target file: ScrappedProducts.xlsx (parent folder)")
            print("✅ Skip logic: Products with 2+ SINs will be skipped")
            print("="*80)
            
            if not os.path.exists(scrapped_products_file):
                print(f"❌ ERROR: {scrapped_products_file} not found!")
                logger.error(f"{scrapped_products_file} not found")
                return False
            
            print(f"✅ Found: {os.path.abspath(scrapped_products_file)}")
            
            # Load manufacturer mapping
            print("\n📂 Loading manufacturer mapping...")
            if not self.load_manufacturer_mapping():
                print("❌ ERROR: Failed to load manufacturer mapping")
                return False
            print("✅ Manufacturer mapping loaded successfully!")
            
            # Read Excel data from ScrappedProducts.xlsx
            print("\n📊 Reading ScrappedProducts.xlsx...")
            try:
                df = pd.read_excel(scrapped_products_file)
                logger.info(f"Loaded {len(df)} rows from {scrapped_products_file}")
                print(f"✅ Loaded {len(df)} total rows")
            except Exception as e:
                print(f"❌ ERROR: Could not read {scrapped_products_file}: {str(e)}")
                return False
            
            # Verify required columns exist
            required_cols = ['Item Number', 'Manufacturer Long Name', 'Unit of Measure', 'Links', 'SIN1', 'SIN2', 'SIN3']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                print(f"❌ ERROR: Missing required columns: {missing_cols}")
                logger.error(f"Missing columns: {missing_cols}")
                return False
            
            print(f"✅ All required columns found")
            
            # Validate range
            if start_row < 0 or end_row >= len(df) or start_row > end_row:
                print(f"❌ ERROR: Invalid range {start_row}-{end_row} for {len(df)} rows")
                return False
            
            # Get rows in range that need scraping
            rows_to_scrape = []
            for i in range(start_row, end_row + 1):
                row = df.iloc[i]
                if not self.row_has_two_sins(row):
                    # Check if Links column is not empty
                    if pd.notna(row['Links']) and str(row['Links']).strip() != '':
                        rows_to_scrape.append(i)
            
            if not rows_to_scrape:
                print(f"\n✅ All products in range (rows {start_row+1}-{end_row+1}) already have 2+ SINs!")
                print("="*80)
                return True
            
            total_in_range = end_row - start_row + 1
            print(f"\n📋 Range: Rows {start_row+1}-{end_row+1} ({total_in_range} products)")
            print(f"✅ Already complete: {total_in_range - len(rows_to_scrape)} rows")
            print(f"🔄 Need scraping: {len(rows_to_scrape)} rows")
            print(f"⏱️  Estimated time: ~{len(rows_to_scrape) * 15 / 60:.1f} minutes")
            
            # Show sample
            if len(rows_to_scrape) <= 5:
                print(f"\n📝 Rows to process:")
                for idx in rows_to_scrape:
                    item_num = df.at[idx, 'Item Number']
                    sin1 = df.at[idx, 'SIN1'] if pd.notna(df.at[idx, 'SIN1']) else 'Empty'
                    sin2 = df.at[idx, 'SIN2'] if pd.notna(df.at[idx, 'SIN2']) else 'Empty'
                    print(f"   • Row {idx+1}: {item_num} (SIN1={sin1}, SIN2={sin2})")
            else:
                print(f"\n📝 Sample of rows to process (first 5):")
                for idx in rows_to_scrape[:5]:
                    item_num = df.at[idx, 'Item Number']
                    sin1 = df.at[idx, 'SIN1'] if pd.notna(df.at[idx, 'SIN1']) else 'Empty'
                    sin2 = df.at[idx, 'SIN2'] if pd.notna(df.at[idx, 'SIN2']) else 'Empty'
                    print(f"   • Row {idx+1}: {item_num} (SIN1={sin1}, SIN2={sin2})")
                print(f"   ... and {len(rows_to_scrape) - 5} more rows")
            
            # Ask for confirmation for large ranges
            if len(rows_to_scrape) > 10:
                print("\n" + "="*80)
                confirm = input("❓ Continue with SIN scraping? (yes/no): ").strip().lower()
                if confirm not in ['yes', 'y']:
                    print("\n❌ Operation cancelled by user.")
                    return False
            
            # Setup web driver - ask user preference
            print("\n🌐 Setting up web driver...")
            
            # Ask if user wants headless mode
            headless_choice = input("Run in headless mode (faster, no browser window)? (yes/no, default=no): ").strip().lower()
            use_headless = headless_choice in ['yes', 'y']
            
            if use_headless:
                print("🚀 Running in HEADLESS mode (faster, browser hidden)")
            else:
                print("👁️  Running in VISIBLE mode (you can see the browser)")
            
            self.setup_driver(headless=use_headless)
            print("✅ Web driver initialized successfully!")
            
            successful_scrapes = 0
            start_time = time.time()
            total = len(rows_to_scrape)
            
            print("\n" + "="*80)
            print("🚀 STARTING SIN SCRAPING")
            print("="*80)
            print(f"💾 Auto-save: Every 50 products")
            print("="*80)
            
            for offset, row_idx in enumerate(rows_to_scrape, 1):
                try:
                    row = df.iloc[row_idx]
                    item_number = row['Item Number']
                    manufacturer = row['Manufacturer Long Name']
                    unit_of_measure = row['Unit of Measure']
                    gsa_url = row['Links']
                    
                    print(f"\n{'='*80}")
                    print(f"🔄 [{offset}/{total}] Row {row_idx+1} | Item: {item_number}")
                    print(f"{'='*80}")
                    
                    # FIRST: Check if any SIN column contains "SIN not found" - skip immediately
                    has_sin_not_found = False
                    for col in ['SIN1', 'SIN2', 'SIN3']:
                        val = row[col]
                        if pd.notna(val) and str(val).strip().lower() == 'sin not found':
                            has_sin_not_found = True
                            break
                    
                    if has_sin_not_found:
                        print(f"⏭️  Row contains 'SIN not found' - SKIPPING (already attempted)")
                        logger.info(f"Row {row_idx+1} contains 'SIN not found', skipping")
                        continue
                    
                    # Check how many SINs are already filled
                    existing_sins = []
                    existing_sin_details = []
                    for col in ['SIN1', 'SIN2', 'SIN3']:
                        val = row[col]
                        if pd.notna(val) and str(val).strip() != '' and str(val).strip().lower() not in ['nan', 'sin not found']:
                            sin_value = str(val).strip()
                            existing_sins.append(sin_value)
                            existing_sin_details.append(f"{col}={sin_value}")
                    
                    sins_needed = 2 - len(existing_sins)
                    
                    if existing_sins:
                        print(f"📊 Existing SINs: {len(existing_sins)} ({', '.join(existing_sin_details)}) | Need: {sins_needed} more")
                    else:
                        print(f"📊 Existing SINs: 0 | Need: {sins_needed} SINs")
                    
                    if sins_needed <= 0:
                        print(f"✅ Row already has 2+ SINs - SKIPPING")
                        logger.info(f"Row {row_idx+1} already has {len(existing_sins)} SINs, skipping")
                        continue
                    
                    # Start timing
                    product_start_time = time.time()
                    
                    # Scrape SINs (only scrape what we need)
                    print(f"🌐 Navigating to GSA page to scrape {sins_needed} SIN(s)...")
                    sins_scraped = self.scrape_gsa_page_for_sins(gsa_url, manufacturer, unit_of_measure, max_sins=sins_needed)
                    
                    product_time = time.time() - product_start_time
                    
                    if sins_scraped:
                        # Fill SIN columns
                        sins_filled_count = 0
                        sin_columns = ['SIN1', 'SIN2', 'SIN3']
                        for col in sin_columns:
                            if pd.isna(df.at[row_idx, col]) or str(df.at[row_idx, col]).strip() == '' or str(df.at[row_idx, col]).strip().lower() == 'nan':
                                if sins_scraped:
                                    df.at[row_idx, col] = sins_scraped.pop(0)
                                    sins_filled_count += 1
                                    if not sins_scraped:
                                        break
                        
                        successful_scrapes += 1
                        total_sins_for_row = len(existing_sins) + sins_filled_count
                        
                        if total_sins_for_row >= 2:
                            if existing_sins:
                                print(f"✅ SUCCESS! Completed row with {sins_filled_count} new SIN(s) [Total: {total_sins_for_row}/2] - Time: {product_time:.1f}s")
                            else:
                                print(f"✅ SUCCESS! Scraped {sins_filled_count} SIN(s) [Total: {total_sins_for_row}/2] - Time: {product_time:.1f}s")
                        else:
                            print(f"⚠️  PARTIAL! Scraped {sins_filled_count} new SIN(s) [Total: {total_sins_for_row}/2, need {2-total_sins_for_row} more] - Time: {product_time:.1f}s")
                        logger.info(f"Successfully scraped {sins_filled_count} SINs for row {row_idx+1} (total now: {total_sins_for_row})")
                    else:
                        # Mark as "SIN not found" in empty columns
                        for col in ['SIN1', 'SIN2', 'SIN3']:
                            if pd.isna(df.at[row_idx, col]) or str(df.at[row_idx, col]).strip() == '' or str(df.at[row_idx, col]).strip().lower() == 'nan':
                                df.at[row_idx, col] = 'SIN not found'
                                sins_needed -= 1
                                if sins_needed <= 0:
                                    break
                        
                        print(f"⚠️  No SINs found - Marked as 'SIN not found' - Time: {product_time:.1f}s")
                        logger.warning(f"No SINs found for row {row_idx+1}")
                    
                    # Calculate ETA
                    elapsed_time = time.time() - start_time
                    avg_time = elapsed_time / offset
                    remaining = total - offset
                    eta_minutes = (remaining * avg_time) / 60
                    
                    print(f"⏱️  Current: {product_time:.1f}s | Avg: {avg_time:.1f}s | ETA: {eta_minutes:.1f}min")
                    print(f"📊 Success Rate: {(successful_scrapes/offset*100):.1f}% ({successful_scrapes}/{offset})")
                    
                    # Auto-save every 50 rows (with backup)
                    if offset % 50 == 0:
                        print(f"\n💾 Auto-saving progress (with backup)...")
                        try:
                            # Temporarily set excel file path for save function
                            original_path = self.excel_file_path
                            self.excel_file_path = scrapped_products_file
                            self.save_results_to_excel(df)
                            self.excel_file_path = original_path
                            print(f"✅ Progress saved at row {row_idx+1} (backup created)")
                            logger.info(f"Auto-save with backup completed at row {row_idx+1}")
                        except Exception as e:
                            print(f"⚠️  Auto-save failed: {str(e)}")
                            logger.error(f"Auto-save error: {str(e)}")
                    
                    # Rate limiting - prevent "Unexpected Error" from GSA
                    time.sleep(2.0)  # Wait between products to avoid rate limiting errors
                    
                    # Automatic browser restart every 100 products to prevent memory leaks
                    if offset % 100 == 0 and offset > 0:
                        try:
                            print(f"\n🔄 Restarting browser to prevent memory leaks...")
                            # Remember if we're in headless mode
                            was_headless = hasattr(self, '_headless_mode') and self._headless_mode
                            self.driver.quit()
                            time.sleep(2)
                            self.setup_driver(headless=was_headless)
                            print(f"✅ Browser restarted successfully")
                            logger.info(f"Browser restarted at product {offset}")
                        except Exception as restart_err:
                            logger.error(f"Browser restart error: {str(restart_err)}")
                            print(f"⚠️  Browser restart failed, continuing with current session")
                    
                except Exception as e:
                    logger.error(f"Error processing row {row_idx+1}: {str(e)}")
                    print(f"❌ ERROR processing row: {str(e)}")
                    continue
            
            # Final save (with backup)
            print("\n" + "="*80)
            print("💾 SAVING FINAL RESULTS (with backup)")
            print("="*80)
            try:
                # Temporarily set excel file path for save function
                original_path = self.excel_file_path
                self.excel_file_path = scrapped_products_file
                self.save_results_to_excel(df)
                self.excel_file_path = original_path
                print("✅ All data saved successfully with backup!")
                logger.info("Final save with backup completed")
            except Exception as e:
                print(f"❌ ERROR saving final results: {str(e)}")
                logger.error(f"Final save error: {str(e)}")
                return False
            
            # Statistics
            total_time = time.time() - start_time
            print("\n" + "="*80)
            print("🎉 SIN SCRAPING COMPLETED!")
            print("="*80)
            print(f"📊 Final Statistics:")
            print(f"   • Total rows processed: {total}")
            print(f"   • ✅ Successful scrapes: {successful_scrapes} ({(successful_scrapes/total*100):.1f}%)")
            print(f"   • ⚠️  No SINs found: {total - successful_scrapes}")
            print(f"   • ⏱️  Total time: {total_time/60:.1f} minutes")
            print(f"   • ⚡ Average speed: {total_time/total:.1f}s per product")
            print("="*80)
            logger.info(f"SIN scraping completed: {successful_scrapes}/{total} successful")
            
            return True
            
        except Exception as e:
            logger.error(f"Error in SIN scraping automation: {str(e)}")
            print(f"\n❌ FATAL ERROR: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()
    
    def run_scraping_missing_only(self):
        """Scrape only rows with missing GSA data - EXACT same flow as run_scraping_custom_range"""
        try:
            print("\n" + "="*60)
            print("🚀 OPTION 5: SCRAPE MISSING ROWS ONLY")
            print("="*60)
            print("📋 This option will identify and scrape rows with missing GSA data")
            print("="*60)
            
            # Load manufacturer mapping
            print("\n📂 Loading manufacturer mapping...")
            if not self.load_manufacturer_mapping():
                print("❌ ERROR: Failed to load manufacturer mapping")
                logger.error("Failed to load manufacturer mapping")
                return False
            print("✅ Manufacturer mapping loaded successfully!")

            # Read Excel data
            print("\n📊 Reading Excel file...")
            df, column_mapping = self.read_excel_data()
            if df is None:
                print("❌ ERROR: No data found in Excel file")
                logger.error("No data found in Excel file")
                return False
            print(f"✅ Excel file loaded! Total rows: {len(df)}")

            # Identify rows with missing data (unique logic for this function)
            print("\n🔍 Analyzing data to find missing rows...")
            missing_rows = self.identify_missing_rows(df)
            
            START_ROW= int(input("Enter the row number to start from: "))
            original_count = len(missing_rows)
            missing_rows = [row_idx for row_idx in missing_rows if row_idx >= START_ROW]
            print(f"⚠️  TEMPORARY: Filtering to start from row {START_ROW + 1}")
            print(f"   Original missing rows: {original_count}")
            print(f"   Filtered missing rows (from row {START_ROW + 1}): {len(missing_rows)}")
            
            if not missing_rows:
                print("\n" + "="*60)
                print("✅ SUCCESS! All products have been scraped!")
                print("="*60)
                print("🎉 No missing rows found - all GSA data is complete!")
                print("="*60)
                logger.info("No missing rows found in the dataset")
                return True
            
            print("\n" + "="*60)
            print("📋 MISSING ROWS ANALYSIS")
            print("="*60)
            print(f"🔍 Found {len(missing_rows)} rows with missing GSA data (starting from row {START_ROW + 1})")
            logger.info(f"Found {len(missing_rows)} rows with missing data (filtered from row {START_ROW + 1})")
            
            # Show sample of missing rows
            print(f"\n📝 Sample of missing rows (showing first 10):")
            for idx in missing_rows[:10]:
                stock_number = df.at[idx, column_mapping['stock_number']]
                print(f"   • Row {idx+1}: {stock_number}")
            
            if len(missing_rows) > 10:
                print(f"   ... and {len(missing_rows) - 10} more rows")
            
            # Ask for confirmation
            print("\n" + "="*60)
            print(f"🚀 READY TO SCRAPE")
            print("="*60)
            print(f"📊 Total missing rows to scrape: {len(missing_rows)}")
            print(f"⏱️  Estimated time: ~{len(missing_rows) * 5 / 60:.1f} minutes (approx 5s per product)")
            print("="*60)
            confirm = input("\n❓ Continue with scraping? (yes/no): ").strip().lower()
            
            if confirm not in ['yes', 'y']:
                print("\n❌ Operation cancelled by user.")
                return False

            print("\n" + "="*60)
            print("🔧 INITIALIZING SCRAPER")
            print("="*60)
            print("🌐 Setting up web driver...")
            # Setup web driver
            self.setup_driver()
            print("✅ Web driver initialized successfully!")
            print("="*60)

            successful_scrapes = 0
            start_time = time.time()

            total = len(missing_rows)
            
            print("\n" + "="*60)
            print("🎯 STARTING SCRAPING PROCESS")
            print("="*60)
            print(f"📊 Total products to scrape: {total}")
            print(f"💾 Auto-save: Every 10 products")
            print("="*60)
            print("\n🚀 Beginning scraping...\n")
            
            for offset, i in enumerate(missing_rows, 1):
                try:
                    gsa_url = df.at[i, column_mapping['links']]
                    manufacturer = df.at[i, column_mapping['manufacturer']]
                    unit_of_measure = df.at[i, column_mapping['unit_of_measure']]
                    stock_number = df.at[i, column_mapping['stock_number']]

                    if pd.isna(gsa_url) or not str(gsa_url).strip():
                        logger.warning(f"Row {i+1}: No GSA URL found for stock number {stock_number}")
                        continue

                    print("\n" + "-"*60)
                    print(f"🔄 [{offset}/{total}] Processing Row {i+1}")
                    print(f"📦 Product: {stock_number}")
                    print("-"*60)
                    logger.info(f"Processing row {i+1}: {stock_number}")

                    # Verify driver is ready before scraping
                    if not self.driver:
                        logger.error(f"ERROR: Driver is None before scraping row {i+1}!")
                        print(f"❌ ERROR: Driver not initialized for row {i+1} - skipping")
                        continue

                    # Start timing for this product
                    product_start_time = time.time()

                    # Scrape and filter by manufacturer + unit inside scrape_gsa_page
                    print(f"🌐 Loading page...")
                    logger.info(f"About to scrape row {i+1}: {gsa_url}")
                    products_data = self.scrape_gsa_page(gsa_url, manufacturer, unit_of_measure)
                    logger.info(f"Scraping completed for row {i+1}, got {len(products_data) if products_data else 0} products")
                    
                    # Calculate timing
                    product_time = time.time() - product_start_time
                    
                    # Warn if scraping was suspiciously fast (less than 3 seconds - should at least wait for page load)
                    if product_time < 3.0:
                        logger.warning(f"WARNING: Scraping completed very quickly ({product_time:.2f}s) for row {i+1} - this might indicate an issue")
                        print(f"⚠️  WARNING: Scraping was very fast ({product_time:.2f}s) - might not have waited properly")

                    if products_data:
                        successful_scrapes += 1
                        self.update_dataframe_with_results(df, i, products_data)
                        print(f"✅ SUCCESS! Found {len(products_data)} matching product(s)")
                        print(f"⏱️  Time taken: {product_time:.1f}s")
                        print(f"💾 Data saved to Excel")
                        logger.info(
                            f"Successfully scraped {len(products_data)} products for row {i+1} in {product_time:.1f}s"
                        )
                    else:
                        print(f"⚠️  No matching products found")
                        print(f"📝 Product: {stock_number}")
                        print(f"⏱️  Time taken: {product_time:.1f}s")
                        logger.warning(
                            f"No matching products for row {i+1}: {stock_number} in {product_time:.1f}s"
                        )

                    # Calculate ETA for custom range
                    elapsed_time = time.time() - start_time
                    avg_time_per_product = elapsed_time / offset
                    remaining_products = total - offset
                    eta_seconds = remaining_products * avg_time_per_product
                    eta_hours = eta_seconds / 3600
                    eta_minutes = (eta_seconds % 3600) / 60
                    
                    print(f"📊 Progress Stats:")
                    print(f"   • Current: {product_time:.1f}s")
                    print(f"   • Average: {avg_time_per_product:.1f}s/product")
                    if eta_hours >= 1:
                        print(f"   • ETA: {eta_hours:.1f}h {eta_minutes:.0f}m remaining")
                    else:
                        print(f"   • ETA: {eta_minutes:.0f}m remaining")

                    # Save periodically inside ranges as well
                    if offset % 10 == 0:
                        self.save_results_to_excel(df)
                        print(f"💾 Progress saved! (Every 10 products)")
                        print(f"📁 Data saved at row {i+1}")

                    time.sleep(2)  # Rate limiting - wait between requests

                except Exception as e:
                    logger.error(f"Error processing row {i+1}: {str(e)}")
                    continue

            # Final save
            print("\n" + "="*60)
            print("💾 SAVING FINAL RESULTS")
            print("="*60)
            self.save_results_to_excel(df)
            print("✅ All data saved successfully!")

            total_time = time.time() - start_time
            logger.info("Missing rows scraping completed!")
            logger.info(f"Rows: {len(missing_rows)} missing rows ({total} items)")
            logger.info(f"Successful scrapes: {successful_scrapes}")
            logger.info(
                f"Total time: {total_time:.1f} seconds ({total_time/60:.1f} minutes)"
            )
            
            print("\n" + "="*60)
            print("🎉 MISSING ROWS SCRAPING COMPLETED!")
            print("="*60)
            print(f"📊 Final Statistics:")
            print(f"   • Total rows processed: {total}")
            print(f"   • ✅ Successful scrapes: {successful_scrapes}")
            print(f"   • ⚠️  Failed/No matches: {total - successful_scrapes}")
            print(f"   • 📈 Success rate: {(successful_scrapes/total*100):.1f}%")
            print(f"   • ⏱️  Total time: {total_time/60:.1f} minutes ({total_time:.1f} seconds)")
            print(f"   • ⚡ Average speed: {total_time/total:.1f}s per product")
            print("="*60)
            print("✅ All results have been saved to Excel file!")
            print("="*60)
            
            return True

        except Exception as e:
            logger.error(f"Error in missing rows scraping: {str(e)}")
            return False
        finally:
            if self.driver:
                self.driver.quit()

def main():
    """Main function with interactive menu"""
    print("="*60)
    print("GSA SCRAPING AUTOMATION - STEP 3")
    print("="*60)
    print("Scraping Price, Contractor, Contract# from GSA pages")
    print("Fuzzy matching: Manufacturer + Unit of Measure")
    print("Rate limited for stability")
    print("="*60)
    
    # File paths - Updated to use ScrappedProducts.xlsx in parent folder
    # For Options 1-5 (old scraping modes), use essendant file if it exists
    # For Option 6 (SIN scraping), always use ScrappedProducts.xlsx
    essendant_excel_file = "essendant-product-list_with_gsa_scraped_data.xlsx"
    scrapped_products_file = "../ScrappedProducts.xlsx"
    manufacturer_mapping_file = "../2 coverting mfr names into root form/coverting to root form/original_to_root.csv"
    
    # Check which Excel file to use
    # For now, default to ScrappedProducts.xlsx since essendant file is deleted
    if os.path.exists(scrapped_products_file):
        excel_file = scrapped_products_file
        print(f"✅ Using: ScrappedProducts.xlsx (parent folder)")
    elif os.path.exists(essendant_excel_file):
        excel_file = essendant_excel_file
        print(f"✅ Using: essendant-product-list_with_gsa_scraped_data.xlsx (current folder)")
    else:
        print(f"❌ ERROR: No Excel file found!")
        print(f"   Looked for:")
        print(f"   - {scrapped_products_file}")
        print(f"   - {essendant_excel_file}")
        return
    
    if not os.path.exists(manufacturer_mapping_file):
        print(f"❌ ERROR: Manufacturer mapping file not found: {manufacturer_mapping_file}")
        return
    
    print(f"✅ Manufacturer mapping: {manufacturer_mapping_file}")
    print("="*60)
    
    # Display menu
    while True:
        print("\n" + "="*60)
        print("SCRAPING AUTOMATION MENU")
        print("="*60)
        print("1. Test Mode (First 10 products)")
        print("2. Custom Range (Specify start and end)")
        print("3. Full Automation (All 19,590 products)")
        print("4. Single Product (by Item Stock Number-Butted)")
        print("5. Scrape Missing Rows Only (Re-scrape failed/empty rows)")
        print("6. SIN Scraping Mode (Extract SIN numbers from product pages)")
        print("7. Exit")
        print("="*60)
        
        try:
            choice = input("Enter your choice (1-7): ").strip()
            
            if choice == "1":
                print("\nRunning TEST MODE (first 10 products)...")
                automation = GSAScrapingAutomation(excel_file, manufacturer_mapping_file)
                success = automation.run_scraping_test_mode(10)
                if success:
                    print("\nSUCCESS: Test scraping completed successfully!")
                    print(f"All data saved to {os.path.basename(excel_file)}")
                else:
                    print("\nERROR: Test scraping failed!")
                    
            elif choice == "2":
                print("\nCUSTOM RANGE MODE")
                print("="*40)
                
                # Get total number of products
                df = pd.read_excel(excel_file)
                total_products = len(df)
                print(f"Total products available: {total_products}")
                
                try:
                    start_row = int(input(f"Enter start row (1-{total_products}): ")) - 1
                    end_row = int(input(f"Enter end row ({start_row + 1}-{total_products}): ")) - 1
                    
                    if start_row < 0 or end_row >= total_products or start_row > end_row:
                        print("ERROR: Invalid range specified!")
                        continue
                    
                    count = end_row - start_row + 1
                    print(f"\nRunning CUSTOM RANGE MODE (rows {start_row + 1}-{end_row + 1}, {count} products)...")
                    
                    automation = GSAScrapingAutomation(excel_file, manufacturer_mapping_file)
                    success = automation.run_scraping_custom_range(start_row, end_row)
                    if success:
                        print(f"\nSUCCESS: Custom range scraping completed successfully!")
                        print(f"All data saved to {os.path.basename(excel_file)}")
                    else:
                        print(f"\nERROR: Custom range scraping failed!")
                        
                except ValueError:
                    print("ERROR: Please enter valid numbers!")
                    continue
                    
            elif choice == "3":
                print("\nFULL AUTOMATION MODE")
                print("="*40)
                print("WARNING: This will process ALL 19,590 products!")
                print("Estimated time: 10-15 hours")
                print("Progress will be saved every 100 products")
                
                confirm = input("\nAre you sure you want to continue? (yes/no): ").strip().lower()
                if confirm in ['yes', 'y']:
                    print("\nRunning FULL AUTOMATION (all products)...")
                    automation = GSAScrapingAutomation(excel_file, manufacturer_mapping_file)
                    success = automation.run_scraping_full()
                    if success:
                        print("\nSUCCESS: Full automation completed successfully!")
                        print(f"All data saved to {os.path.basename(excel_file)}")
                    else:
                        print("\nERROR: Full automation failed!")
                else:
                    print("Full automation cancelled.")
                    
            elif choice == "4":
                print("\nSINGLE PRODUCT MODE")
                print("="*40)
                stock_query = input("Enter Item Stock Number-Butted: ").strip()
                if not stock_query:
                    print("ERROR: Stock number cannot be empty!")
                    continue
                automation = GSAScrapingAutomation(excel_file, manufacturer_mapping_file)
                success = automation.run_scraping_single(stock_query)
                if success:
                    print("\nSUCCESS: Single product scraping completed!")
                    print(f"All data saved to {os.path.basename(excel_file)}")
                else:
                    print("\nERROR: Single product scraping failed!")

            elif choice == "5":
                print("\nSCRAPE MISSING ROWS ONLY MODE")
                print("="*40)
                print("This will identify and re-scrape rows with missing GSA data")
                print("(Where GSA PRICE, Contractor, and Contract# are all empty)")
                
                automation = GSAScrapingAutomation(excel_file, manufacturer_mapping_file)
                success = automation.run_scraping_missing_only()
                if success:
                    print("\nSUCCESS: Missing rows scraping completed!")
                    print(f"All data saved to {os.path.basename(excel_file)}")
                else:
                    print("\nERROR: Missing rows scraping failed!")

            elif choice == "6":
                print("\n🎯 OPENING SIN SCRAPING MENU...")
                automation = GSAScrapingAutomation(excel_file, manufacturer_mapping_file)
                automation.run_sin_scraping_menu()

            elif choice == "7":
                print("\nExiting...")
                break
                
            else:
                print("ERROR: Invalid choice! Please enter 1, 2, 3, 4, 5, 6, or 7.")
                
        except KeyboardInterrupt:
            print("\n\nOperation cancelled by user.")
            break
        except Exception as e:
            print(f"\nERROR: {str(e)}")
            continue

if __name__ == "__main__":
    main()
