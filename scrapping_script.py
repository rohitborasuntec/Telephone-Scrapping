# scrapping_script_optimized.py
import re
import time
import random
import zipfile
import shutil
import os
import csv
import traceback
import json
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List
import pandas as pd
import fitz
import subprocess
import re
import platform
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
    ElementClickInterceptedException,
)

# Import unified logger
from logger import get_logger

# Initialize logger
logger = get_logger()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DATE_TO_EX = "29 Jun 2026"
MAX_RETRIES_PER_ROW = 3
MAX_PROXY_ATTEMPTS = 5
CHECKPOINT_EVERY = 5

# Application types to filter
APPLICATION_TYPES = [
    "Full Planning Application", "Householder", "Full Planning Permission",
    "Listed Building Consent (S8 P&LBC 1990)", "Fast Track FLH (Householder App)",
    "Full Planning Permission (Householder)", "Householder Planning Application",
    "Householder Application", "Full Application", "Lawful Development Certificate -Proposed",
    "Listed Building Consent", "Householders Extensions Prior Approval",
    "Lawful Development Certificate proposed", "Householder Prior Approval",
    "General Permitted Development - Extns", "Full planning",
    "Lawful Development - Proposed Use", "Full", "Outline",
    "Householder Planning Permission", "Planning Application", "Listed Building",
    "Planning Permission in Principle", "Prior Approval - Larger Household Extension",
    "Householder Planning Consent", "Proposed Lawful Development",
    "Prior Approval Larger Home Extension", "Domestic Application (Householder)",
    "C of Lawfulness for Proposed Use or Dev", "Residential Extensions", "Pre-Application",
    "Lawful Development Certificate", "Outline Application", "Full App via planning portal",
    "Certificate Proposed Development", "Certificate of Lawfulness (Proposed)",
    "Full Application (8 Weeks)", "Outline Application (8 Weeks)"
]

# ---------------------------------------------------------------------------
# Directory Setup
# ---------------------------------------------------------------------------
BASE_DIR = Path("Data") / DATE_TO_EX
LOG_DIR = Path("Logs") / DATE_TO_EX
HTML_DEBUG_DIR = Path("HTML_Debugger")
FAILED_HTML_DIR = HTML_DEBUG_DIR / "Failed_Dir"
LINK_HTML_DIR = HTML_DEBUG_DIR / "Link_Dir"
ZIP_DIR = BASE_DIR / "zip"
PDF_DIR = BASE_DIR / "pdf"
TEMP_DIR = BASE_DIR / "temp_dir"
OUTPUT_DIR = Path("Output")

# Create all directories
for dir_path in [LOG_DIR, HTML_DEBUG_DIR, FAILED_HTML_DIR, LINK_HTML_DIR,
                 ZIP_DIR, PDF_DIR, TEMP_DIR, OUTPUT_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# State Management
# ---------------------------------------------------------------------------
class ScraperState:
    """Manage scraper state with checkpointing"""
    
    def __init__(self, state_file: str = "scraper_state.json"):
        self.state_file = Path(state_file)
        self.state = self._load_state()
        
    def _load_state(self) -> Dict:
        """Load state from file or create new"""
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {
            "completed_urls": [],
            "failed_urls": [],
            "partial_urls": [],  # List of dicts: {url: last_link_index}
            "current_url": None,
            "current_link_index": 0,
            "processed_count": 0
        }
    
    def save(self):
        """Save state to file"""
        with open(self.state_file, 'w') as f:
            json.dump(self.state, f, indent=2)
    
    def mark_url_completed(self, url: str):
        if url not in self.state["completed_urls"]:
            self.state["completed_urls"].append(url)
        self.state["processed_count"] += 1
        self.save()
    
    def mark_url_failed(self, url: str, error: str = ""):
        if url not in self.state["failed_urls"]:
            self.state["failed_urls"].append({"url": url, "error": error})
        self.save()
    
    def mark_url_partial(self, url: str, last_index: int):
        # Remove existing partial entry if any
        self.state["partial_urls"] = [p for p in self.state["partial_urls"] if p["url"] != url]
        self.state["partial_urls"].append({"url": url, "last_index": last_index})
        self.save()
    
    def get_partial_info(self, url: str) -> Optional[int]:
        """Get last index for partial URL, None if not partial"""
        for item in self.state["partial_urls"]:
            if item["url"] == url:
                return item["last_index"]
        return None
    
    def is_url_completed(self, url: str) -> bool:
        return url in self.state["completed_urls"]
    
    def is_url_failed(self, url: str) -> bool:
        return any(item["url"] == url for item in self.state["failed_urls"])

# ---------------------------------------------------------------------------
# Proxy Manager
# ---------------------------------------------------------------------------
class ProxyManager:
    """Manage proxy rotation"""
    
    def __init__(self):
        self.proxies = self._load_proxies()
        self.current_index = -1
        self.attempts = 0
    
    def _load_proxies(self) -> List[str]:
        """Load proxies from file or return empty list"""
        proxy_file = Path("proxies.txt")
        if proxy_file.exists():
            with open(proxy_file, 'r') as f:
                return [line.strip() for line in f if line.strip()]
        return []
    
    def get_next_proxy(self) -> Optional[str]:
        """Get next proxy in rotation"""
        if not self.proxies:
            return None
        self.current_index = (self.current_index + 1) % len(self.proxies)
        self.attempts += 1
        return self.proxies[self.current_index]
    
    def reset_attempts(self):
        self.attempts = 0
    
    def should_skip(self) -> bool:
        """Check if we should skip to next URL"""
        return self.attempts >= MAX_PROXY_ATTEMPTS

# ---------------------------------------------------------------------------
# HTML Debugger
# ---------------------------------------------------------------------------
class HTMLDebugger:
    """Save HTML pages for debugging"""
    
    @staticmethod
    def save_page(driver, base_dir: Path, filename: str, suffix: str = ""):
        """Save current page HTML"""
        try:
            if suffix:
                filename = f"{filename}_{suffix}"
            file_path = base_dir / f"{filename}.html"
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logger.debug(f"HTML saved: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"Failed to save HTML: {e}")
            return None
    
    @staticmethod
    def save_link_page(driver, index: int, link_index: int):
        """Save link page HTML"""
        return HTMLDebugger.save_page(
            driver, 
            LINK_HTML_DIR, 
            f"index_{index}_link_{link_index}"
        )
    
    @staticmethod
    def save_failed_page(driver, index: int, link_index: int = None, error: str = ""):
        """Save failed page HTML"""
        suffix = f"link_{link_index}" if link_index is not None else "row"
        return HTMLDebugger.save_page(
            driver,
            FAILED_HTML_DIR,
            f"failed_row_{index}_{suffix}",
            error[:50] if error else ""
        )

# ---------------------------------------------------------------------------
# Checkpoint Manager
# ---------------------------------------------------------------------------
class CheckpointManager:
    """Manage data checkpoints"""
    
    def __init__(self, base_name: str = "checkpoint"):
        self.base_name = base_name
        self.temp_file = Path(f"{base_name}_temp.csv")
        self.final_file = Path(f"{base_name}_final.csv")
        self.failed_file = Path("failed_items.csv")
        self.partial_file = Path("partial_urls.csv")
    
    def save_temp(self, data: List[Dict]):
        """Save temporary data"""
        if data:
            pd.DataFrame(data).to_csv(self.temp_file, index=False)
            logger.debug(f"Temp data saved: {len(data)} rows")
    
    def save_final(self, data: List[Dict]):
        """Save final output"""
        if data:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            final_path = OUTPUT_DIR / f"{self.base_name}_{timestamp}.csv"
            pd.DataFrame(data).to_csv(final_path, index=False)
            logger.info(f"Final data saved: {final_path}")
            return final_path
        return None
    
    def log_failed(self, row_index: int, url: str, case_ref: str, stage: str, error: str):
        """Log failed item"""
        is_new = not self.failed_file.exists()
        with open(self.failed_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if is_new:
                writer.writerow(["timestamp", "row_index", "url", "case_ref", "stage", "error"])
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                row_index, url, case_ref, stage, str(error)
            ])
    
    def log_partial(self, url: str, last_index: int):
        """Log partial URL with last processed index"""
        # Read existing
        existing = {}
        if self.partial_file.exists():
            df = pd.read_csv(self.partial_file)
            for _, row in df.iterrows():
                existing[row['url']] = row['last_index']
        
        existing[url] = last_index
        
        # Write back
        pd.DataFrame([
            {"url": k, "last_index": v, "updated": datetime.now().isoformat()}
            for k, v in existing.items()
        ]).to_csv(self.partial_file, index=False)

# ---------------------------------------------------------------------------
# Browser Manager
# ---------------------------------------------------------------------------
class BrowserManager:
    """Manage browser lifecycle with proxy support"""
    
    def __init__(self, download_dir: Path):
        self.driver = None
        self.download_dir = download_dir
        self.current_proxy = None
        self.use_proxy = False
    
    def _create_options(self, proxy: Optional[str] = None):
        """Create Chrome options"""
        options = uc.ChromeOptions()
        
        # Download preferences
        prefs = {
            "download.default_directory": str(self.download_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
        }
        options.add_experimental_option("prefs", prefs)
        
        # Common arguments
        options.add_argument("--disable-backgrounding-occluded-windows")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--no-sandbox")  # For Linux
        options.add_argument("--disable-dev-shm-usage")  # For Linux
        
        # User agent
        from commons import user_agent_list
        options.add_argument(f"--user-agent={random.choice(user_agent_list)}")
        
        # Proxy if provided
        if proxy:
            options.add_argument(f'--proxy-server={proxy}')
            logger.info(f"Using proxy: {proxy}")
        
        return options
    
    def _get_chrome_version(self):
        """Get installed Chrome major version"""
        try:
            if platform.system() == "Windows":
                import winreg
                try:
                    # Try current user first
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
                    version = winreg.QueryValueEx(key, "version")[0]
                    return int(version.split('.')[0])
                except:
                    # Try local machine
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\BLBeacon")
                    version = winreg.QueryValueEx(key, "version")[0]
                    return int(version.split('.')[0])
                    
            elif platform.system() == "Linux":
                # Try common Chrome executables
                for cmd in ['google-chrome', 'google-chrome-stable', 'chromium-browser', 'chromium']:
                    try:
                        result = subprocess.run([cmd, '--version'], capture_output=True, text=True)
                        if result.returncode == 0:
                            version = result.stdout.strip().split()[-1]
                            return int(version.split('.')[0])
                    except:
                        continue
                        
            elif platform.system() == "Darwin":  # macOS
                try:
                    result = subprocess.run(
                        ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'],
                        capture_output=True, text=True
                    )
                    if result.returncode == 0:
                        version = result.stdout.strip().split()[-1]
                        return int(version.split('.')[0])
                except:
                    pass
                    
        except Exception as e:
            logger.debug(f"Could not detect Chrome version: {e}")
        
        return None

    def start(self, proxy: Optional[str] = None, headless: bool = False):
        """Start browser with dynamic Chrome version detection"""
        self.close()
        
        options = self._create_options(proxy)
        self.current_proxy = proxy
        
        # Try to auto-detect Chrome version
        chrome_version = self._get_chrome_version()
        
        if chrome_version:
            logger.info(f"Detected Chrome version: {chrome_version}")
            try:
                self.driver = uc.Chrome(
                    options=options,
                    headless=headless,
                    version_main=chrome_version
                )
                logger.info(f"Browser started with Chrome {chrome_version}")
            except Exception as e:
                logger.warning(f"Failed with detected version {chrome_version}: {e}")
                # Fallback to auto-detection
                try:
                    self.driver = uc.Chrome(
                        options=options,
                        headless=headless
                    )
                    logger.info("Browser started with auto-detection")
                except Exception as e2:
                    logger.error(f"Failed to start browser: {e2}")
                    return False
        else:
            # No version detected, use auto-detection
            try:
                self.driver = uc.Chrome(
                    options=options,
                    headless=headless
                )
                logger.info("Browser started with auto-detection")
            except Exception as e:
                logger.error(f"Failed to start browser: {e}")
                return False
        
        # Set download behavior
        self.driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": str(self.download_dir.resolve())
            }
        )
        
        self.driver.maximize_window()
        self.driver.set_page_load_timeout(100)
        logger.info("Browser configured successfully" + (f" with proxy: {proxy}" if proxy else ""))
        return True

    def close(self):
        """Close browser"""
        if self.driver:
            try:
                self.driver.quit()
                logger.debug("Browser closed")
            except:
                pass
            self.driver = None
    
    def restart_with_proxy(self, proxy: Optional[str] = None):
        """Restart browser with new proxy"""
        self.close()
        time.sleep(2)
        return self.start(proxy)
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

# ---------------------------------------------------------------------------
# Utility Functions
# ---------------------------------------------------------------------------
def random_scroll(driver, min_scroll=200, max_scroll=800, min_pause=0.5, max_pause=2, iterations=None):
    """Random scroll to avoid detection"""
    if iterations is None:
        iterations = random.randint(5, 15)
    
    for _ in range(iterations):
        direction = 1 if random.random() < 0.8 else -1
        pixels = random.randint(min_scroll, max_scroll) * direction
        driver.execute_script("window.scrollBy({top: arguments[0], behavior: 'smooth'});", pixels)
        time.sleep(random.uniform(min_pause, max_pause))

def check_for_bot(driver):
    """Check if bot detection triggered"""
    time.sleep(random.randint(5, 8))
    page_source = driver.page_source
    
    bot_indicators = [
        "One moment, we're checking you're not a bot.",
        "Performing security function",
        "Just a moment",
        "Checking your browser",
        "Access denied"
    ]
    
    for indicator in bot_indicators:
        if indicator in page_source:
            logger.warning(f"Bot check detected: {indicator}")
            random_scroll(driver, iterations=random.randint(1, 4))
            time.sleep(random.randint(3, 8))
            return True
    
    return False

def safe_get(driver, url, wait_ready=True, timeout=20):
    """Safe page navigation with error handling"""
    logger.debug(f"Navigating to: {url}")
    
    try:
        driver.get(url)
    except TimeoutException:
        logger.error(f"Page load timed out: {url}")
        return False
    except WebDriverException as e:
        error = str(e).lower()
        browser_errors = [
            "err_name_not_resolved", "err_connection_refused",
            "err_connection_timed_out", "err_connection_reset",
            "err_internet_disconnected", "err_address_unreachable",
        ]
        if any(err in error for err in browser_errors):
            logger.error(f"Network error: {e}")
            return False
        raise
    
    if driver.current_url.startswith("chrome-error://"):
        logger.error(f"Chrome error page: {url}")
        return False
    
    # Check for error pages
    title = (driver.title or "").lower()
    source = (driver.page_source or "").lower()
    error_strings = [
        "this site can't be reached", "err_name_not_resolved",
        "access denied", "service unavailable", "maintenance",
        "502 bad gateway", "503 service unavailable",
        "internal server error"
    ]
    
    if any(err in title or err in source for err in error_strings):
        logger.error(f"Website returned error page: {url}")
        return False
    
    if wait_ready:
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            logger.error(f"Document readyState timeout: {url}")
            return False
    
    return True

def safe_find(driver, by, value, timeout=10):
    """Safe element find with retry"""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )

def safe_click(driver, by, value, timeout=10):
    """Safe element click with retry"""
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )
    el.click()
    return el

def safe_select_by_text(driver, by, value, text, timeout=10):
    """Safe select dropdown by text"""
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    Select(el).select_by_visible_text(text)

def handle_advanced_button(driver):
    """Handle browser advanced button for certificate errors"""
    try:
        driver.find_element(By.XPATH, "//button[@id='details-button'][contains(text(),'Advanced')]").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//a[@id='proceed-link']").click()
        logger.debug("Handled advanced button")
        return True
    except:
        return False

# ---------------------------------------------------------------------------
# Main Scraper Class
# ---------------------------------------------------------------------------
class PlanningScraper:
    """Main scraper class"""
    
    def __init__(self, input_file: str = "Url.xlsx"):
        self.input_file = Path(input_file)
        self.browser = None
        self.state = ScraperState()
        self.checkpoint = CheckpointManager()
        self.proxy_manager = ProxyManager()
        self.html_debugger = HTMLDebugger()
        self.results = []
        self.processed_urls = set()
        
        # App type counter
        self.app_type_count = {app: 0 for app in APPLICATION_TYPES}
        
        # Load input data
        self.df = self._load_input()
        self._load_previous_results()
    
    def _load_input(self) -> pd.DataFrame:
        """Load input Excel file"""
        if not self.input_file.exists():
            logger.error(f"Input file not found: {self.input_file}")
            return pd.DataFrame()
        
        df = pd.read_excel(self.input_file)
        logger.info(f"Loaded {len(df)} rows from {self.input_file}")
        
        # Add columns if not present
        for col in ['Status', 'Start', 'Processed']:
            if col not in df.columns:
                df[col] = ''
        
        return df
    
    def _load_previous_results(self):
        """Load previous results if available"""
        temp_file = OUTPUT_DIR / "scraped_data_temp.csv"
        if temp_file.exists():
            try:
                prev_df = pd.read_csv(temp_file)
                self.results = prev_df.to_dict('records')
                logger.info(f"Loaded {len(self.results)} previous results")
            except:
                pass
    
    def _is_application_type_valid(self, app_type: str) -> bool:
        """Check if application type is in target list"""
        return app_type in APPLICATION_TYPES
    
    def _wait_for_download(self, before_files, timeout=60, stable_checks=2, poll=1.0):
        """Wait for download to complete"""
        start = time.time()
        stable_path = None
        last_size = None
        stable_count = 0
        
        while time.time() - start < timeout:
            try:
                current = set(self.browser.download_dir.iterdir())
            except FileNotFoundError:
                time.sleep(poll)
                continue
            
            new_files = [
                f for f in (current - before_files)
                if f.is_file() and f.suffix.lower() not in (".crdownload", ".tmp", ".part")
            ]
            
            if not new_files:
                time.sleep(poll)
                continue
            
            candidate = max(new_files, key=lambda p: p.stat().st_ctime)
            
            try:
                size_now = candidate.stat().st_size
            except FileNotFoundError:
                stable_path, last_size, stable_count = None, None, 0
                time.sleep(poll)
                continue
            
            if candidate == stable_path and size_now == last_size:
                stable_count += 1
            else:
                stable_count = 1
            
            stable_path, last_size = candidate, size_now
            
            if stable_count >= stable_checks:
                suffix = candidate.suffix.lower()
                if suffix == ".zip":
                    return "zip", candidate
                elif suffix == ".pdf":
                    return "pdf", candidate
                raise ValueError(f"Unexpected file type: {candidate.name}")
            
            time.sleep(poll)
        
        raise TimeoutError(f"No download completed within {timeout}s")
    
    def _extract_via_web(self, driver, case_ref=None):
        """Extract data from web page"""
        time.sleep(random.randint(5, 12))
        
        try:
            details = safe_find(driver, By.XPATH, '//div[@class="addressCrumb"]')
        except:
            details = None
        
        # Extract fields with fallbacks
        fields = {
            "Applicant Name": ("//th[contains(text(),'Applicant Name')]/following-sibling::td", ""),
            "Address": (".//div[@class='address']", ""),
            "Reference": (".//div[@class='caseNumber']", ""),
            "Proposal": (".//div[@class='description']", ""),
            "Agent Name": ("//th[contains(text(),'Agent Name')]/following-sibling::td", ""),
            "Agent Address": ("//th[contains(text(),'Agent Address')]/following-sibling::td", ""),
        }
        
        result = {"Case Ref (internal)": case_ref}
        for key, (xpath, default) in fields.items():
            try:
                if details and xpath.startswith(".//"):
                    result[key] = details.find_element(By.XPATH, xpath).text
                else:
                    result[key] = driver.find_element(By.XPATH, xpath).text
            except:
                result[key] = default
        
        # Check if applicant equals agent
        result["Applicant = Agent"] = result.get("Agent Name") == result.get("Applicant Name")
        
        return result
    
    def _extract_via_pdf(self, pdf_path, case_ref=None):
        """Extract data from PDF"""
        doc = fitz.open(pdf_path)
        pdf_text = ""
        for page in doc:
            pdf_text += page.get_text()
        doc.close()
        
        pdf_text = re.sub(r'\r\n?', '\n', pdf_text)
        pdf_text = re.sub(r'[ \t]+', ' ', pdf_text)
        
        if not pdf_text:
            raise ValueError(f"No text in PDF: {pdf_path}")
        
        # Extract reference
        ref_match = re.search(r'Planning Portal Reference:\s*PP-(\d+)', pdf_text, re.S)
        ref_number = ref_match.group(0) if ref_match else ""
        
        # Extract address
        address_match = re.search(
            r'Address line 1\s*(.*?)\s*Address line 2(.*?)\s*Address line 3(.*?)\s*Town/City\s*(.*?)\s*County\s*(.*?)\s*Country(.*?)Postcode\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})',
            pdf_text, re.S
        )
        if address_match:
            address = ", ".join([
                g.strip() for g in address_match.groups()[:6] if g and g.strip()
            ])
        else:
            address = ""
        
        # Extract applicant name
        applicant_block = re.search(
            r"Name/Company(.*?)Are you an agent acting on behalf of the applicant\?",
            pdf_text, re.S
        )
        applicant_name = ""
        if applicant_block:
            block = applicant_block.group(1)
            title_match = re.search(r"Title\s*\n([^\n]*)", block)
            first_match = re.search(r"First name\s*\n([^\n]*)", block)
            surname_match = re.search(r"Surname\s*\n([^\n]*)", block)
            title = title_match.group(1).strip() if title_match else ""
            first = first_match.group(1).strip() if first_match else ""
            surname = surname_match.group(1).strip() if surname_match else ""
            applicant_name = " ".join(filter(None, [title, first, surname]))
        
        # Extract agent name
        agent_match = re.search(
            r"Agent Details.*?Title\s*\n(?P<title>.*?)\nFirst name\s*\n(?P<first_name>.*?)\nSurname\s*\n(?P<surname>.*?)\nCompany Name\s*\n(?P<company>.*?)\nAddress",
            pdf_text, re.DOTALL
        )
        agent_name = ""
        if agent_match:
            agent_name = " ".join([
                agent_match.group("title").strip(),
                agent_match.group("first_name").strip(),
                agent_match.group("surname").strip()
            ])
        
        return {
            "Address": address.replace(ref_number, ""),
            "Applicant Name": applicant_name,
            "Agent Name": agent_name,
            "Applicant = Agent": agent_name == applicant_name,
            "Application Form PDF Available": "Yes",
            "Case Ref (internal)": case_ref,
            "Source File": str(pdf_path),
        }
    
    def _download_and_extract(self, driver, pdf_name, case_ref=None):
        """Download and extract from application form"""
        check_for_bot(driver)
        
        app_form_xpath = (
            "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
            "'abcdefghijklmnopqrstuvwxyz'), 'application form')]"
            "/../td/input[@class='bulkCheck']"
        )
        
        if not driver.find_elements(By.XPATH, app_form_xpath):
            logger.info(f"No application form found for case: {case_ref}")
            return None
        
        # Get PDF link
        pdf_link = ""
        try:
            pdf_link = driver.find_element(
                By.XPATH,
                "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                "'abcdefghijklmnopqrstuvwxyz'), 'application form')]"
                "/../td/input[@class='bulkCheck']/../..//a"
            ).get_attribute('href')
        except:
            pass
        
        # Click checkbox
        driver.find_element(By.XPATH, app_form_xpath).click()
        time.sleep(random.randint(5, 12))
        
        # Click download
        before = set(self.browser.download_dir.iterdir())
        safe_click(driver, By.ID, "downloadFiles")
        logger.info(f"Download triggered for case: {case_ref}")
        
        file_type, downloaded_path = self._wait_for_download(before, timeout=60)
        logger.info(f"Downloaded: {downloaded_path}")
        
        # Process downloaded file
        if file_type == "zip":
            zip_path = ZIP_DIR / f"{pdf_name}.zip"
            shutil.copy2(downloaded_path, zip_path)
            downloaded_path.unlink(missing_ok=True)
            
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(PDF_DIR)
            
            extracted_pdfs = list(PDF_DIR.glob("*.pdf"))
            if not extracted_pdfs:
                raise FileNotFoundError(f"No PDF in ZIP: {zip_path}")
            
            final_path = PDF_DIR / f"{pdf_name}.pdf"
            max(extracted_pdfs, key=lambda p: p.stat().st_mtime).rename(final_path)
            
        elif file_type == "pdf":
            final_path = PDF_DIR / f"{pdf_name}.pdf"
            shutil.copy2(downloaded_path, final_path)
            downloaded_path.unlink(missing_ok=True)
        else:
            raise ValueError(f"Unsupported file type: {file_type}")
        
        # Extract data from PDF
        item_ext = self._extract_via_pdf(final_path, case_ref=case_ref)
        item_ext["Pdf Link"] = pdf_link
        return item_ext
    
    def _process_link(self, driver, index: int, link_index: int, url: str) -> Dict:
        """Process a single link and return extracted data"""
        result = {
            "S.No.": link_index,
            "Url": "",
            "Reference": "",
            "Address": "",
            "Proposal": "",
            "Application Type": "",
            "Application Form PDF Available": "",
            "Applicant = Agent": "",
            "Agent Name": "",
            "Agent Address": "",
            "Applicant Name": "",
            "Applicant Address": "",
            "Pdf Link": ""
        }
        
        link_row = {
            "Total": 0,
            "Row Index": index,
            "Link Index": link_index,
            "Case Ref": "",
            "Status": "",
            "Error": ""
        }
        
        try:
            # Find and click the link
            element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
            if link_index >= len(element_links):
                link_row["Status"] = "Skipped"
                return result
            
            element = element_links[link_index]
            case_ref = element.text.strip() or f"row{index}_link{link_index}"
            link_row["Case Ref"] = case_ref
            
            # Open in new tab
            element.send_keys(Keys.CONTROL + Keys.RETURN)
            driver.switch_to.window(driver.window_handles[-1])
            
            # Check for application summary
            if not driver.find_elements(By.XPATH, "//h1[contains(.,'Application Summary')]"):
                logger.warning(f"No Application Summary for {case_ref}")
                self.html_debugger.save_failed_page(driver, index, link_index, "no_summary")
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                link_row["Status"] = "Failed"
                return result
            
            check_for_bot(driver)
            
            # Click information tab
            safe_click(
                driver,
                By.XPATH,
                "//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                "'abcdefghijklmnopqrstuvwxyz'), 'information')]"
            )
            time.sleep(random.randint(5, 12))
            
            # Check application type
            app_type = ""
            for _ in range(5):
                if driver.find_elements(By.XPATH, "//th[contains(text(),'Application Type')]"):
                    app_type = driver.find_element(
                        By.XPATH,
                        "//th[contains(text(),'Application Type')]/following-sibling::td"
                    ).text
                    break
                time.sleep(random.randint(5, 15))
            
            if not app_type:
                link_row["Status"] = "Failed"
                link_row["Error"] = "Application Type not found"
                return result
            
            if not self._is_application_type_valid(app_type):
                link_row["Status"] = "Skipped"
                link_row["Error"] = f"Invalid type: {app_type}"
                logger.info(f"Skipped {case_ref}: {app_type}")
                return result
            
            # Valid application type - process it
            result["Url"] = driver.current_url
            result["Application Type"] = app_type
            link_row["Url"] = driver.current_url
            link_row["Status"] = "Found"
            
            # Extract web data
            try:
                web_data = self._extract_via_web(driver, case_ref=case_ref)
                result.update(web_data)
                link_row["Extracted"] = "Web Data"
            except Exception as e:
                logger.error(f"Web extraction failed for {case_ref}: {e}")
                link_row["Extracted"] = "Not Extracted"
            
            self.app_type_count[app_type] += 1
            
            # Try to download PDF
            pdf_name = f"{index}_{link_index}_{int(time.time())}"
            try:
                # Check for document tab
                if driver.find_elements(
                    By.XPATH,
                    "//a[@id='tab_documents']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                    "'abcdefghijklmnopqrstuvwxyz'), 'document')]"
                ):
                    doc_link = safe_find(driver, By.XPATH, '//a[@id="tab_documents"]').get_attribute("href")
                    safe_get(driver, doc_link)
                    check_for_bot(driver)
                    
                    item_ext = self._download_and_extract(driver, pdf_name, case_ref=case_ref)
                    if item_ext:
                        result.update({k: v for k, v in item_ext.items() if v})
                        link_row["PDF Data"] = "Extracted"
                        link_row["Status"] = "Completed"
                
                elif driver.find_elements(
                    By.XPATH,
                    "//a[@id='tab_externalDocuments']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                    "'abcdefghijklmnopqrstuvwxyz'), 'documents')]"
                ):
                    driver.find_element(
                        By.XPATH,
                        "//a[@id='tab_externalDocuments']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                        "'abcdefghijklmnopqrstuvwxyz'), 'documents')]"
                    ).click()
                    time.sleep(random.randint(5, 8))
                    
                    # Process external documents
                    try:
                        pdf_link = ""
                        try:
                            select_el = driver.find_element(By.XPATH, '//select[@name="searchResult_length"]')
                            Select(select_el).select_by_value("100")
                        except:
                            pass
                        
                        try:
                            pdf_link = driver.find_element(
                                By.XPATH,
                                '//a[@data-bind="click: OpenDocument, href: Link"]'
                                '//div[contains(text(),"ApplicationForm")]//ancestor::a'
                            ).get_attribute('href')
                            driver.find_element(
                                By.XPATH,
                                '//a[@data-bind="click: OpenDocument, href: Link"]'
                                '//div[contains(text(),"ApplicationForm")]'
                            ).click()
                        except:
                            try:
                                driver.find_element(
                                    By.XPATH,
                                    "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                                    "'abcdefghijklmnopqrstuvwxyz'), 'application form')]"
                                ).click()
                            except:
                                pass
                        
                        before = set(self.browser.download_dir.iterdir())
                        try:
                            driver.find_element(
                                By.XPATH,
                                '//tr[@class="selected"]//a[@class="viewDocument"] | '
                                '//a[contains(text(),"Download")] | '
                                '//a[@data-bind="click: OpenDocument, href: Link"]'
                                '//div[contains(text(),"Application Form")]/..'
                            ).click()
                        except:
                            pass
                        
                        result["Pdf Link"] = pdf_link
                        file_type, downloaded_path = self._wait_for_download(before, timeout=60)
                        
                        if file_type == "pdf":
                            final_path = PDF_DIR / f"{pdf_name}.pdf"
                            shutil.copy2(downloaded_path, final_path)
                            downloaded_path.unlink(missing_ok=True)
                            
                            item_ext = self._extract_via_pdf(final_path, case_ref=case_ref)
                            result.update({k: v for k, v in item_ext.items() if v})
                            link_row["PDF Data"] = "Extracted"
                            link_row["Status"] = "Completed"
                    except Exception as e:
                        logger.error(f"External document extraction failed: {e}")
                        link_row["PDF Data"] = "Not Extracted"
                else:
                    logger.warning(f"No document tab for {case_ref}")
                    link_row["PDF Data"] = "Not Extracted"
                
            except Exception as e:
                logger.error(f"PDF download failed for {case_ref}: {e}")
                link_row["PDF Data"] = "Not Extracted"
            
            # Close tab and switch back
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
            
            # Save HTML for debugging
            self.html_debugger.save_link_page(driver, index, link_index)
            
        except Exception as e:
            link_row["Status"] = "Failed"
            link_row["Error"] = str(e)
            logger.error(f"Link {link_index} failed: {e}")
            self.html_debugger.save_failed_page(driver, index, link_index, str(e))
            
            if len(driver.window_handles) > 1:
                try:
                    driver.close()
                except:
                    pass
            try:
                driver.switch_to.window(driver.window_handles[0])
            except:
                pass
        
        # Update result and row
        self.results.append(result)
        return result
    
    def _process_row(self, index: int, row: pd.Series) -> bool:
        """Process a single row/website"""
        url = row["Url"]
        if "https://" not in url:
            url = "https://" + url
        
        # Check state
        if self.state.is_url_completed(url):
            logger.info(f"URL already completed: {url}")
            return True
        
        # Check partial progress
        start_index = self.state.get_partial_info(url) or 0
        if start_index > 0:
            logger.info(f"Resuming {url} from link {start_index}")
        
        # Build search URL
        if "weeklyList" not in url:
            base_url = url.split("uk/")[0] + "uk/"
            search_url = base_url + "online-applications/search.do?action=weeklyList"
        else:
            search_url = url
        
        logger.info(f"Processing row {index}: {url}")
        
        # Try with retries
        for attempt in range(MAX_RETRIES_PER_ROW):
            logger.info(f"Attempt {attempt + 1}/{MAX_RETRIES_PER_ROW}")
            
            try:
                # Start browser (no proxy for first attempt)
                use_proxy = attempt > 0
                if use_proxy:
                    proxy = self.proxy_manager.get_next_proxy()
                    if proxy:
                        if not self.browser.restart_with_proxy(proxy):
                            continue
                    else:
                        logger.warning("No proxy available, continuing without")
                        if not self.browser.start():
                            continue
                else:
                    if not self.browser.start():
                        continue
                
                driver = self.browser.driver
                
                # Navigate to page
                if not safe_get(driver, search_url):
                    logger.warning(f"Failed to load {search_url}")
                    if use_proxy:
                        self.proxy_manager.attempts += 1
                        if self.proxy_manager.should_skip():
                            logger.info(f"Proxy attempts exhausted for {url}")
                            break
                    continue
                
                # Handle advanced button
                handle_advanced_button(driver)
                
                # Check for weekly list
                if not driver.find_elements(By.XPATH, "//h1[contains(.,'Weekly List')]"):
                    logger.warning(f"No Weekly List found for {url}")
                    self.html_debugger.save_failed_page(driver, index, None, "no_weekly_list")
                    self.state.mark_url_failed(url, "No Weekly List")
                    return False
                
                # Set filters
                time.sleep(random.randint(3, 5))
                safe_select_by_text(driver, By.ID, "week", DATE_TO_EX)
                logger.debug(f"Week filter set to {DATE_TO_EX}")
                
                safe_click(driver, By.XPATH, '//input[@value="DC_Decided"]')
                safe_click(driver, By.XPATH, '//input[@value="Search"]')
                logger.debug("Search submitted")
                
                time.sleep(random.randint(3, 6))
                
                # Set results per page
                select_element = safe_find(driver, By.ID, 'resultsPerPage')
                Select(select_element).select_by_value("100")
                logger.debug("Results per page set to 100")
                
                safe_click(driver, By.XPATH, '//input[@value="Go"]')
                time.sleep(random.randint(5, 8))
                
                # Save HTML for debugging
                self.html_debugger.save_page(driver, LINK_HTML_DIR, f"row_{index}_listing")
                
                # Get all links
                element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                total_links = len(element_links)
                logger.info(f"Found {total_links} links for {url}")
                
                # Process each link from start_index
                success_count = 0
                for link_idx in range(start_index, total_links):
                    logger.info(f"Processing link {link_idx + 1}/{total_links}")
                    
                    # Process the link
                    result = self._process_link(driver, index, link_idx, url)
                    
                    # Update state periodically
                    if (link_idx - start_index + 1) % CHECKPOINT_EVERY == 0:
                        self.state.mark_url_partial(url, link_idx + 1)
                        self.checkpoint.save_temp(self.results)
                    
                    success_count += 1
                
                # All links processed
                self.state.mark_url_completed(url)
                self.df.at[index, 'Status'] = 'Completed'
                self.df.at[index, 'Processed'] = datetime.now().isoformat()
                
                # Save checkpoint
                self.checkpoint.save_temp(self.results)
                self.checkpoint.save_final(self.results)
                
                logger.info(f"Successfully processed {url}")
                return True
                
            except Exception as e:
                logger.error(f"Attempt {attempt + 1} failed for {url}: {e}")
                logger.exception("Full traceback:")
                
                if self.browser.driver:
                    self.html_debugger.save_failed_page(
                        self.browser.driver, index, None, str(e)
                    )
                
                # Check if it's a bot/rate limit issue
                if self.browser.driver and "too many requests" in str(e).lower():
                    logger.warning("Rate limit detected, restarting browser")
                    self.browser.close()
                    time.sleep(random.randint(10, 20))
                
                # Try with proxy
                if attempt >= MAX_RETRIES_PER_ROW - 1:
                    self.state.mark_url_failed(url, str(e))
                    self.df.at[index, 'Status'] = 'Failed'
                    return False
                
                time.sleep(5 * (attempt + 1))
        
        return False
    
    def run(self):
        """Main run method"""
        logger.info("=" * 70)
        logger.info("Starting scraper run")
        logger.info("=" * 70)
        
        if self.df.empty:
            logger.error("No data to process")
            return
        
        with BrowserManager(TEMP_DIR) as browser:
            self.browser = browser
            
            for index, row in self.df.iterrows():
                # Skip already completed
                if self.state.is_url_completed(row["Url"]):
                    logger.info(f"Skipping completed URL: {row['Url']}")
                    continue
                
                # Process row
                success = self._process_row(index, row)
                
                # Save updated DataFrame
                self.df.to_excel(self.input_file, index=False)
                
                # Log final status
                if success:
                    logger.info(f"Row {index} completed successfully")
                else:
                    logger.error(f"Row {index} failed after all attempts")
        
        # Final save
        self.checkpoint.save_final(self.results)
        self.checkpoint.save_temp(self.results)
        
        logger.info("=" * 70)
        logger.info("Scraper run completed")
        logger.info(f"Total results: {len(self.results)}")
        logger.info(f"App type counts: {self.app_type_count}")
        logger.info("=" * 70)

# ---------------------------------------------------------------------------
# Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        scraper = PlanningScraper("Url.xlsx")
        scraper.run()
    except Exception as e:
        logger.critical(f"Fatal error: {e}")
        logger.exception("Full traceback:")
    finally:
        logger.info("Scraper finished")