from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
    ElementClickInterceptedException,
)
import re, scrapy, time, pandas as pd, fitz
import random, zipfile, functools, shutil, os, glob, csv, traceback
import logging
from logging.handlers import RotatingFileHandler
from random import randint
from pathlib import Path
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from commons import user_agent_list
import traceback
# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
date_to_ex = "10 Aug 2026"

LOG_DIR = Path("Logs") / date_to_ex
LOG_DIR.mkdir(parents=True, exist_ok=True)

_run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

LOG_FILE = LOG_DIR / f"scraper_{_run_stamp}.log"

logger = logging.getLogger("planning_scraper")
logger.setLevel(logging.DEBUG)
logger.propagate = False


if not logger.handlers:
    _fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(funcName)s:%(lineno)d | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        LOG_FILE, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(_fmt)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(_fmt)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

logger.info("=" * 70)
logger.info("Scraper run starting - log file: %s", LOG_FILE)
logger.info("=" * 70)


application_types = [
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


zip_dir = Path("Data") / date_to_ex / "zip"
zip_dir.mkdir(parents=True, exist_ok=True)

DOWNLOAD_DIR = Path("Data") / date_to_ex / "temp_dir"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

pdf_dir = Path("Data") / date_to_ex / "pdf"
pdf_dir.mkdir(parents=True, exist_ok=True)

res_col = [
    "Applicant Name", "Full Postal Address", "Planning Reference Number", "Brief Project Description"
]

results = []
RESULTS_CSV = "results_checkpoint.csv"
FAILED_CSV = "failed_items.csv"
CHECKPOINT_EVERY = 10  # rows

# --- NEW: new_df is now built from a list of per-case dicts, not by
# repeatedly calling .at[len(df), col] = value on a growing DataFrame.
# One dict == one case/link. Fields are filled in as that case is
# processed, then the whole dict is appended once. The DataFrame itself
# is only materialised when we need to save/checkpoint it.
new_df_rows = []
LINKS_CSV = "Links Data.csv"
LINKS_CSV_BACKUP = "Links Data Backup.csv"

driver = None  # global, guarded at shutdown


def save_new_df():
    """Build the DataFrame from new_df_rows and write it to disk."""
    try:
        pd.DataFrame(new_df_rows).to_csv(LINKS_CSV, index=False)
    except Exception:
        pd.DataFrame(new_df_rows).to_csv(LINKS_CSV_BACKUP, index=False)


# ---------------------------------------------------------------------------
# Generic retry decorator
# ---------------------------------------------------------------------------
def retry(max_attempts=3, delay=2, backoff=2, exceptions=(Exception,), on_retry=None):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            attempt = 1
            current_delay = delay
            last_exc = None
            while attempt <= max_attempts:
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exc = e
                    if attempt == max_attempts:
                        logger.error(
                            "[RETRY FAILED] %s gave up after %s attempts: %s: %s",
                            func.__name__, max_attempts, type(e).__name__, e
                        )
                        raise
                    logger.warning(
                        "[RETRY %s/%s] %s raised %s: %s. Retrying in %ss...",
                        attempt, max_attempts, func.__name__, type(e).__name__, e, current_delay
                    )
                    if on_retry:
                        try:
                            on_retry(*args, **kwargs)
                        except Exception as cb_err:
                            logger.error("[on_retry callback error] %s", cb_err)
                    time.sleep(current_delay)
                    current_delay *= backoff
                    attempt += 1
            raise last_exc
        return wrapper
    return decorator


RETRYABLE_SELENIUM_EXC = (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementClickInterceptedException,
)

file_path = r'Url.xlsx'

def log_failed_item(row_index, url, case_ref, stage, error):
    is_new = not Path(FAILED_CSV).exists()
    with open(FAILED_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(["timestamp", "row_index", "url", "case_ref", "stage", "error"])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            row_index, url, case_ref, stage, str(error)
        ])
    logger.error(
        "[LOGGED FAILURE] row=%s stage=%s case_ref=%s url=%s error=%s",
        row_index, stage, case_ref, url, error
    )


def checkpoint_results():
    if not results:
        return
    try:
        pd.DataFrame(results).to_csv(RESULTS_CSV, index=False)
        logger.info("Checkpoint written: %s rows -> %s", len(results), RESULTS_CSV)
    except Exception as e:
        logger.error("[checkpoint_results failed] %s", e)

def open_chrome():
    global driver
    try:
        driver.quit()
        # open_vpn()
    except:
        pass

    options = uc.ChromeOptions()

    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }

    logger.debug("Download directory resolved to: %s", DOWNLOAD_DIR.resolve())
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument(f"--user-agent={random.choice(user_agent_list)}")
    options.add_experimental_option("prefs", prefs)

    @retry(max_attempts=3, delay=5, backoff=2, exceptions=(WebDriverException,))
    def _launch():
        drv = uc.Chrome(
            options=options,
            headless=False,
            version_main=150
        )
        return drv

    logger.info("Launching Chrome...")
    driver = _launch()

    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {
            "behavior": "allow",
            "downloadPath": str(DOWNLOAD_DIR.resolve())
        }
    )

    driver.maximize_window()
    driver.set_page_load_timeout(100)
    logger.info("Chrome launched and configured successfully")


def save_html_by_id(id, base_dir="script_html_pages"):
    id = str(id + 1)
    folder_path = Path(base_dir) / id
    folder_path.mkdir(parents=True, exist_ok=True)

    html_file = folder_path / f"{id}.html"
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(driver.page_source)

    logger.info("HTML saved for debugging: %s", html_file)

    return html_file


def random_scroll(min_scroll=200, max_scroll=800, min_pause=0.5, max_pause=2, iterations=None):

    if iterations is None:
        iterations = randint(5, 15)

    for _ in range(iterations):
        logger.debug("Scrolling ... iteration %s", _)
        direction = 1 if random.random() < 0.8 else -1

        pixels = random.randint(min_scroll, max_scroll) * direction

        driver.execute_script(
            "window.scrollBy({top: arguments[0], behavior: 'smooth'});",
            pixels
        )

        time.sleep(random.uniform(min_pause, max_pause))


def check_for_bot():
    time.sleep(randint(5, 8))

    page_source = driver.page_source

    if "One moment, we're checking you're not a bot." in page_source or "Performing security function" in page_source:
        logger.warning("Bot check detected - performing evasive scrolling")

        random_scroll(
            min_scroll=200,
            max_scroll=800,
            min_pause=0.5,
            max_pause=2,
            iterations=random.randint(1, 4))

        time.sleep(randint(3, 8))
    else:
        logger.debug("No bot check found")

# ---------------------------------------------------------------------------
# Safe navigation / interaction helpers
# ---------------------------------------------------------------------------
def safe_get(url, wait_ready=True, timeout=20):
    logger.debug("Navigating to: %s", url)
    try:
        driver.get(url)

    except TimeoutException:
        logger.error("Page load timed out: %s", url)
        return False

    except WebDriverException as e:
        error = str(e).lower()

        browser_errors = [
            "err_name_not_resolved",
            "err_connection_refused",
            "err_connection_timed_out",
            "err_connection_reset",
            "err_connection_closed",
            "err_internet_disconnected",
            "err_address_unreachable",
            "err_ssl_protocol_error",
            "err_ssl_version_or_cipher_mismatch",
            "err_tunnel_connection_failed",
        ]

        if any(err in error for err in browser_errors):
            logger.error("Browser/network error while opening %s: %s", url, e)
            return False

        raise

    if driver.current_url.startswith("chrome-error://"):
        logger.error("Chrome error page: %s", url)
        return False

    title = (driver.title or "").lower()
    source = (driver.page_source or "").lower()

    error_strings = [
        "this site can't be reached",
        "err_name_not_resolved",
        "err_connection_refused",
        "err_connection_timed_out",
        "dns_probe_finished",
        "access denied",
        "service unavailable",
        "temporarily unavailable",
        "under maintenance",
        "maintenance",
        "502 bad gateway",
        "503 service unavailable",
        "504 gateway timeout",
        "internal server error",
        "just a moment",
        "checking your browser",
    ]

    if any(err in title or err in source for err in error_strings):
        logger.error("Website returned an error page: %s", url)
        return False

    if not title and len(source.strip()) < 50:
        logger.error("Blank page: %s", url)
        return False

    if wait_ready:
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            logger.error("document.readyState timeout: %s", url)
            return False

    logger.debug("Page loaded successfully: %s", url)
    return True


@retry(max_attempts=4, delay=1.5, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def safe_find(by, value, timeout=10):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def safe_click(by, value, timeout=10):
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )
    el.click()
    logger.debug("Clicked element: %s=%s", by, value)
    return el


@retry(max_attempts=3, delay=1.5, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def safe_select_by_text(by, value, text, timeout=10):
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    Select(el).select_by_visible_text(text)
    logger.debug("Selected option '%s' on %s=%s", text, by, value)


@retry(max_attempts=3, delay=1, backoff=2, exceptions=(IndexError, WebDriverException))
def safe_switch_to_window(index=-1):
    driver.switch_to.window(driver.window_handles[index])
    logger.debug("Switched to window index %s", index)


# ---------------------------------------------------------------------------
# Extraction / download logic
# ---------------------------------------------------------------------------

@retry(max_attempts=3, delay=2, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def extract_via_web(case_ref=None):
    time.sleep(randint(5,12))
    details = safe_find(By.XPATH, '//div[@class="addressCrumb"]')
    try:
        app_name = details.find_element(By.XPATH, "//th[contains(text(),'Applicant Name')]/following-sibling::td").text
    except:
        app_name = ""
    try:
        address = details.find_element(By.CLASS_NAME, 'address').text
    except:
        address = ""
    try:
        ref_num = details.find_element(By.CLASS_NAME, 'caseNumber').text
    except:
        ref_num = ""
    try:
        desc = details.find_element(By.CLASS_NAME, 'description').text
    except:
        desc = ""
    try:
        agent_name = details.find_element(By.XPATH, "//th[contains(text(),'Agent Name')]/following-sibling::td").text
    except:
        agent_name = ""
    try:
        agent_address = details.find_element(By.XPATH, "//th[contains(text(),'Agent Address')]/following-sibling::td").text
    except:
        agent_address = ""

    item_ext = {
        "Address": address,
        "Applicant Name": app_name,
        # "Applicant address": address.replace(ref_number, ""),
        "Agent Name":agent_name,
        "Agent Address":agent_address,
        "Applicant = Agent": agent_name==app_name,
        "Reference": ref_num,
        "Application Form PDF Available":"No",
        "Proposal": desc,
        "Case Ref (internal)": case_ref,
    }
    
    logger.info("Data extracted from website for case_ref=%s: ref=%s", case_ref, ref_num)
    logger.debug("Web-extracted data: %s", item_ext)
    return item_ext


def wait_for_download(before_files, timeout=60, stable_checks=2, poll=1.0):
    start = time.time()
    stable_path = None
    last_size = None
    stable_count = 0

    while time.time() - start < timeout:
        try:
            current = set(DOWNLOAD_DIR.iterdir())
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
                logger.info("Download confirmed stable (zip): %s", candidate)
                return "zip", candidate
            if suffix == ".pdf":
                logger.info("Download confirmed stable (pdf): %s", candidate)
                return "pdf", candidate
            raise ValueError(f"Unexpected downloaded file type: {candidate.name}")

        time.sleep(poll)

    raise TimeoutError(f"No new stable download appeared in {DOWNLOAD_DIR} within {timeout}s")


def extract_via_pdf(pdf_path, case_ref=None):

    doc = fitz.open(pdf_path)
    logger.info("Opening PDF for extraction: %s", pdf_path)

    pdf_text = ""
    for page in doc:
        pdf_text += page.get_text()
    doc.close()

    pdf_text = re.sub(r'\r\n?', '\n', pdf_text)
    pdf_text = re.sub(r'[ \t]+', ' ', pdf_text)

    if pdf_text:
        print("Text found")
    else:
        logger.warning("No text in pdf: %s", pdf_path)
        raise ValueError(f"No extractable text in PDF: {pdf_path}")

    ref_regex = r'(Planning Portal Reference:\s*PP-\d+)'
    ref = re.search(ref_regex, pdf_text, re.S)
    ref_number = ref.group(1) if ref else ""

    address_regex = (
        r'Address line 1\s*(.*?)\s*'
        r'Address line 2(.*?)\s*'
        r'Address line 3(.*?)\s*'
        r'Town/City\s*(.*?)\s*'
        r'County\s*(.*?)\s*'
        r'Country(.*?)'
        r'Postcode\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})'
    )

    m = re.search(address_regex, pdf_text, re.S)
    if m:
        address = f"{m.group(1) or ''}, {m.group(2) or ''}, {m.group(3) or ''} {m.group(4) or ''}  {m.group(5) or ''} {m.group(6) or ''}"
        address = ", ".join([part.strip() for part in address.split(",") if part.strip()])
    else:
        # NOTE: was None before, which crashes .replace() below on any PDF
        # that doesn't match this exact template.
        address = ""

    agent_pattern = re.compile(
        r"Agent Details.*?"
        r"Title\s*\n(?P<title>.*?)\n"
        r"First name\s*\n(?P<first_name>.*?)\n"
        r"Surname\s*\n(?P<surname>.*?)\n"
        r"Company Name\s*\n(?P<company>.*?)\n"
        r"Address",
        re.DOTALL
    )
    match = agent_pattern.search(pdf_text)

    if match:
        agent_name = " ".join([
            match.group("title").strip(),
            match.group("first_name").strip(),
            match.group("surname").strip()
        ])
        company_name = match.group("company").strip()
        logger.debug("Agent found: %s (company: %s)", agent_name, company_name)
    else:
        agent_name = ""
        logger.debug("Agent name not found in PDF: %s", pdf_path)

    applicant_block = re.search(
        r"Name/Company(.*?)Are you an agent acting on behalf of the applicant\?",
        pdf_text,
        re.S
    )

    applicant_name = ""
    company_name = ""

    if applicant_block:
        block = applicant_block.group(1)

        title = re.search(r"Title\s*\n([^\n]*)", block)
        first = re.search(r"First name\s*\n([^\n]*)", block)
        surname = re.search(r"Surname\s*\n([^\n]*)", block)
        company = re.search(
            r"Company Name\s*\n(.*?)(?=\nAddress|\nAddress line 1)",
            block,
            re.S
        )

        title = title.group(1).strip() if title else ""
        first = first.group(1).strip() if first else ""
        surname = surname.group(1).strip() if surname else ""

        company_name = company.group(1).strip() if company else ""

        if first and surname:
            applicant_name = " ".join(filter(None, [title, first, surname]))

    logger.debug("Applicant name parsed: %s", applicant_name)

    site_block = re.search(
        r"Site Location(.*?)Description of site location must be completed",
        pdf_text,
        re.S
    )

    site_block = site_block.group(1) if site_block else ""

    def get_field(text, start_label, end_label):
        m = re.search(
            rf"{re.escape(start_label)}\s*\n(.*?)(?=\n{re.escape(end_label)})",
            text,
            re.S,
        )
        return m.group(1).strip() if m else ""

    number = get_field(site_block, "Number", "Suffix")
    property_name = get_field(site_block, "Property Name", "Address Line 1")
    address1 = get_field(site_block, "Address Line 1", "Address Line 2")
    address2 = get_field(site_block, "Address Line 2", "Address Line 3")
    address3 = get_field(site_block, "Address Line 3", "Town/city")
    town = get_field(site_block, "Town/city", "Postcode")

    postcode_match = re.search(
        r'Postcode\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})?',
        site_block
    )
    postcode = postcode_match.group(1) if postcode_match else ""

    site_address = " ".join(
        filter(
            None,
            [
                number,
                property_name,
                address1,
                address2,
                address3,
                town,
                postcode,
            ],
        )
    )

    logger.debug("Site address parsed: %s", site_address)

    agent_block = re.search(
        r"Agent Details(.*?)(?=Contact Details|Description of Proposed Works|Site Area)",
        pdf_text,
        re.S
    )

    agent_address = ""

    if agent_block:
        block = agent_block.group(1)

        line1 = re.search(r"Address line 1\s*\n(.*?)(?=\nAddress line 2)", block, re.S)
        line2 = re.search(r"Address line 2\s*\n(.*?)(?=\nAddress line 3)", block, re.S)
        line3 = re.search(r"Address line 3\s*\n(.*?)(?=\nTown/City)", block, re.S)
        town = re.search(r"Town/City\s*\n(.*?)(?=\nCounty)", block, re.S)
        county = re.search(r"County\s*\n(.*?)(?=\nCountry)", block, re.S)
        country = re.search(r"Country\s*\n(.*?)(?=\nPostcode)", block, re.S)
        postcode = re.search(r"Postcode\s*\n([A-Z0-9 ]*)", block)

        agent_address = " ".join(
            filter(
                None,
                [
                    line1.group(1).strip() if line1 else "",
                    line2.group(1).strip() if line2 else "",
                    line3.group(1).strip() if line3 else "",
                    town.group(1).strip() if town else "",
                    county.group(1).strip() if county else "",
                    country.group(1).strip() if country else "",
                    postcode.group(1).strip() if postcode else "",
                ],
            )
        )
        logger.debug("Agent address parsed: %s", agent_address)
    else:
        agent_address = ""
        logger.debug("Agent address not found in PDF: %s", pdf_path)

    is_same = True if agent_name == applicant_name else False

    item_ext = {
        "Address": site_address.replace(ref_number, ""),
        "Applicant Name": applicant_name.replace("First name", "").replace("Surname", "").strip(),
        "Applicant Address": address.replace(ref_number, ""),
        "Agent Name": agent_name,
        "Agent Address": agent_address.replace(ref_number, ""),
        "Applicant = Agent": is_same,
        "Application Form PDF Available": "Yes",
        "Case Ref (internal)": case_ref,
        "Source File": str(pdf_path),
    }

    logger.info("Data extracted from PDF for case_ref=%s: applicant=%s", case_ref, applicant_name)
    logger.debug("PDF-extracted data: %s", item_ext)
    return item_ext


@retry(max_attempts=3, delay=3, backoff=2, exceptions=(RETRYABLE_SELENIUM_EXC + (TimeoutError, ValueError)))
def download_and_extract(pdf_new_name, case_ref=None):
    check_for_bot()

    app_form_xpath = (
        "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
        "'abcdefghijklmnopqrstuvwxyz'), 'application form')]"
        "/../td/input[@class='bulkCheck']"
    )

    if not driver.find_elements(By.XPATH, app_form_xpath):
        logger.info("Application form checkbox not found - nothing to download (case_ref=%s)", case_ref)
        return None

    try:
        pdf_link = driver.find_element(By.XPATH, "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]/../td/input[@class='bulkCheck']/../..//a").get_attribute('href')
    except:
        pdf_link = ""

    driver.find_element(By.XPATH, app_form_xpath).click()
    time.sleep(randint(5, 12))

    before = set(DOWNLOAD_DIR.iterdir())
    safe_click(By.ID, "downloadFiles")
    logger.info("Download triggered for case_ref=%s, waiting for file...", case_ref)

    file_type, downloaded_path = wait_for_download(before, timeout=60)
    logger.info("Confirmed new download: type=%s path=%s", file_type, downloaded_path)

    if file_type == "zip":
        zip_copy_path = zip_dir / f"{pdf_new_name}.zip"
        shutil.copy2(downloaded_path, zip_copy_path)
        downloaded_path.unlink(missing_ok=True)
        logger.info("Copied ZIP to: %s", zip_copy_path)

        try:
            with zipfile.ZipFile(zip_copy_path, "r") as zf:
                zf.extractall(pdf_dir)
        except zipfile.BadZipFile as e:
            raise ValueError(f"Could not unzip {zip_copy_path}: {e}")

        extracted_pdfs = list(pdf_dir.glob("*.pdf"))
        if not extracted_pdfs:
            raise FileNotFoundError(f"No PDF found inside ZIP: {zip_copy_path}")

        newest_extracted = max(extracted_pdfs, key=lambda p: p.stat().st_mtime)
        final_pdf_path = pdf_dir / f"{pdf_new_name}.pdf"
        if final_pdf_path.exists():
            final_pdf_path.unlink()
        newest_extracted.rename(final_pdf_path)
        logger.info("PDF extracted from ZIP: %s", final_pdf_path)

    elif file_type == "pdf":
        final_pdf_path = pdf_dir / f"{pdf_new_name}.pdf"
        shutil.copy2(downloaded_path, final_pdf_path)
        downloaded_path.unlink(missing_ok=True)
        logger.info("Copied PDF to: %s", final_pdf_path)

    else:
        raise ValueError(f"Unsupported downloaded file type: {file_type}")

    item_ext = extract_via_pdf(final_pdf_path, case_ref=case_ref)
    item_ext["Pdf Link"] = pdf_link

    return item_ext


@retry(max_attempts=3, delay=2, backoff=2, exceptions=(FileNotFoundError, ValueError, OSError))
def file_rename(file_new_name, before_files=None):
    candidates = [
        f for f in DOWNLOAD_DIR.iterdir()
        if f.is_file() and f.suffix.lower() not in (".crdownload", ".part", ".tmp")
    ]

    if before_files is not None:
        candidates = [f for f in candidates if f not in before_files]

    if not candidates:
        raise FileNotFoundError("No new downloaded files found.")

    latest_file = max(candidates, key=lambda f: f.stat().st_ctime)
    new_name = DOWNLOAD_DIR / f"{file_new_name}{latest_file.suffix}"
    latest_file.rename(new_name)

    logger.info("Renamed %s -> %s", latest_file.name, new_name.name)
    return new_name


def is_site_not_working():
    err_xpaths = [
        "//*[contains(text(),'This site can’t be reached')]",
        "//meta[@name='ROBOTS']"
        "//*[contains(text(),'Access denied')]",
        "//*[contains(text(),'404 - File or directory not found.')]",
        "//*[contains(text(),'Page not found')]",
    ]

    for err_x in err_xpaths:
        if driver.find_elements(By.XPATH, err_x):
            print("Matched Tag")
            return True
    return False


def adv_pro():
    logger.debug("Handling browser 'Advanced' bypass interstitial")
    try:
        driver.find_element(By.XPATH, "//button[@id='details-button'][contains(text(),'Advanced')]").click()
        time.sleep(5)
        driver.find_element(By.XPATH, "//button[@id='details-button'][contains(text(),'Advanced')]").click()
        driver.find_element(By.XPATH, "//a[@id='proceed-link']").click()
    except:
        logger.error("Failed ADV PRO")


def main():

    df = pd.read_excel(file_path)

    # df = df.loc[df['Status'] != 'Completed']

    logger.info("Loaded input file %s - %s rows pending", file_path, df.shape[0])

    for index, row in df.iterrows():

        logger.info("=== Row %s starting ===", index)

        url = "https://" + row["Url"] if "https://" not in row["Url"] else row["Url"]
        count = 1

        websites = row["Url"]

        if "weeklyList" not in url:
            url = url.split("uk/")[0] + "uk/online-applications/search.do?action=weeklyList"

        logger.info("Row %s redirecting to: %s", index, url)

        row_attempts = 0
        row_max_attempts = 3

        while row_attempts < row_max_attempts:
            row_attempts += 1

            try:
                if not safe_get(url):
                    logger.warning("Row %s: skipping URL that failed to load: %s", index, url)
                    log_failed_item(index, url, None, "safe_get", "page failed to load")
                    break

                check_for_bot()
                # checking this xpath in case not found possiblity of blocking or need VPN or site not working so used retry with re-opening of browser
                if driver.find_elements(By.XPATH, "//button[@id='details-button'][contains(text(),'Advanced')]"):
                    adv_pro()

                if not driver.find_elements(By.XPATH,"//h1[contains(.,'Weekly List')]"):
                    logger.warning("Row %s: page not found: %s", index, url)
                    # for i in range(3):
                    #     logger.warning(f"Retrying Attempt : {i+1}")
                    #     open_chrome()
                    #     driver.get(url)
                    #     time.sleep(randint(5,12))
                    #     if driver.find_elements(By.XPATH,"//h1[contains(.,'Weekly List')]"):
                    #         break
                    # else:
                    save_html_by_id(index)
                    df = df.loc[df['Status'] != 'Failed']
                    df.at[index, 'Start'] = 0
                    log_failed_item(index, url, None, "listing_page", "page not found")
                    break

                # if is_site_not_working():

                time.sleep(randint(3, 5))
                #selecting drop down 

                safe_select_by_text(By.ID, "week", date_to_ex)
                logger.debug(f"Week filter set to {date_to_ex}")

                safe_click(By.XPATH, '//input[@value="DC_Decided"]')
                safe_click(By.XPATH, '//input[@value="Search"]')
                logger.debug("Search submitted, listing page redirected")

                time.sleep(randint(3, 6))

                select_element = safe_find(By.ID, 'resultsPerPage')
                Select(select_element).select_by_value("100")
                logger.debug("Results-per-page set to 100")

                safe_click(By.XPATH, '//input[@value="Go"]')
                time.sleep(randint(5, 8))

                save_html_by_id(index,base_dir="script_html_pages/Links_Dir")

                element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                n = len(element_links)
                link_index = 0
                # n = 
                logger.info("Row %s: total result links found: %s", index, n)

                success = False
                while link_index < n:
                    logger.info("Running for %s",link_index)
                    case_ref = None
                    temp_dict = {
                        "S.No.": link_index,
                        "Websites": websites,
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

                    link_row = {}
                    try:
                        link_row["Total"] = n
                        link_row["Row Index"] = index
                        link_row["Link Index"] = link_index
                        element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                        element = element_links[link_index]
                        case_ref = element.text.strip() if element.text else f"row{index}_link{link_index}"
                        link_row["Case Ref"] = case_ref
                        element.send_keys(Keys.CONTROL + Keys.RETURN)

                        safe_switch_to_window(-1)
                        app_url = driver.current_url

                        if not driver.find_elements(By.XPATH,"//h1[contains(.,'Application Summary')]"):
                            logger.warning(
                                "Application Summary heading not found - skipping this link "
                                "(case_ref=%s, url=%s)", case_ref, app_url
                            )
                            save_html_by_id(index)
                            if len(driver.window_handles) > 1:
                                driver.close()
                            safe_switch_to_window(0)

                            link_row["Status"] = "Not Working Site"
                            log_failed_item(
                                index, url, case_ref, "detail_page",
                                "Application Summary heading not found (blocked / rate-limited / no VPN)"
                            )
                            df.at[index, 'Start'] = link_index + 1
                            link_index += 1
                            continue

                        check_for_bot()

                        safe_click(By.XPATH,
                                   "//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                                   "'abcdefghijklmnopqrstuvwxyz'), 'information')]")
                        
                        time.sleep(randint(5,12))

                        handled = False

                        for j in range(1, 5):
                            if not driver.find_elements(By.XPATH, "//th[contains(text(),'Application Type')]"):
                                logger.debug("Application Type not visible yet, retry %s/4 (case_ref=%s)", j, case_ref)
                                time.sleep(randint(5, 15))
                                continue

                            app_type = driver.find_element(
                                By.XPATH,
                                "//th[contains(text(),'Application Type')]/following-sibling::td"
                            ).text

                            link_row["Application Type"] = app_type
                            logger.debug("Application Type read: %s (case_ref=%s)", app_type, case_ref)

                            if app_type not in application_types:
                                link_row["Status"] = "Not Found"
                                # logger.info("Case_ref=%s: application type not in target list, skipping: %s", case_ref, app_type)
                                logger.warning("%s not in out app_type list",app_type)
                                # handled = True
                                link_row["Status"] = "Failed"
                                link_row["Error"] = "Application Type field not in our list"
                                # breakpoint()
                                # log_failed_item(index, url, case_ref, "application_type","Application Type field not in our list")
                                break
                            # breakpoint()
                            logger.info("Case_ref=%s: matching application type found: %s", case_ref, app_type)
                            temp_dict["Url"] = driver.current_url
                            temp_dict["Application Type"] = app_type
                            link_row["Url"] = driver.current_url
                            link_row["Status"] = "Found"

                            try:
                                web_data = extract_via_web(case_ref=case_ref)
                                temp_dict.update(web_data)
                                link_row["Extracted"] = "Web Data"
                                link_row["Status"] = "Completed"

                            except Exception as e:
                                link_row["Extracted"] = "Not Extracted"
                                logger.error("Web extraction failed for case_ref=%s: %s", case_ref, e)
                                log_failed_item(index, url, case_ref, "extract_via_web", e)

                            app_typ_count[app_type] += 1

                            pdf_name = f'{index}_{count}'
                            
                            if driver.find_elements(By.XPATH,"//a[@id='tab_documents']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', ""'abcdefghijklmnopqrstuvwxyz'), 'document')]"):
                                doc_link = safe_find(By.XPATH, '//a[@id="tab_documents"]').get_attribute("href")
                                logger.debug("Navigating to documents tab: %s", doc_link)
                                safe_get(doc_link)
                                check_for_bot()

                                try:
                                    item_ext = download_and_extract(pdf_name, case_ref=case_ref)
                                    temp_dict.update({k: v for k, v in item_ext.items() if v != ''})
                                    link_row["PDF Data"] = "Extracted"
                                    link_row["Status"] = "Completed"
                                except Exception as e:
                                    link_row["PDF Data"] = "Not Extracted"
                                    logger.error("[download_and_extract failed after retries] case_ref=%s: %s", case_ref, e)
                                    log_failed_item(index, url, case_ref, "download_and_extract", e)
                                count += 1
                                handled = True

                            elif driver.find_elements(By.XPATH,"//a[@id='tab_externalDocuments']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'documents')]"):
                                driver.find_element(By.XPATH,"//a[@id='tab_externalDocuments']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'documents')]").click()

                                time.sleep(randint(5, 8))
                                safe_click(By.XPATH, '//a[contains(text(),"documents")]')

                                driver.close()
                                safe_switch_to_window(-1)
                                check_for_bot()

                                try:
                                    pdf_link = ""
                                    try:
                                        select_element = driver.find_element(By.XPATH, '//select[@name="searchResult_length"]')
                                        Select(select_element).select_by_value("100")
                                    except NoSuchElementException:
                                        logger.debug("No results-per-page dropdown on external documents page (case_ref=%s)", case_ref)
                                    # breakpoint()
                                    try:
                                        time.sleep(randint(5,8))
                                        pdf_link = driver.find_element(By.XPATH,'//a[@data-bind="click: OpenDocument, href: Link"]//div[contains(text(),"ApplicationForm") or contains(text(),"ApplicationForm")]//ancestor::a').get_attribute('href')
                                        driver.find_element(By.XPATH,'//a[@data-bind="click: OpenDocument, href: Link"]//div[contains(text(),"ApplicationForm") or contains(text(),"ApplicationForm")]').click()
                                        # time.sleep(5)
                                    except NoSuchElementException:
                                        try:
                                            driver.find_element(By.XPATH,"//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]").click()
                                            try:
                                                time.sleep(randint(5,8))
                                                driver.find_element(By.XPATH, "//a[contains(text(),'Download')]")
                                            except NoSuchElementException:
                                                driver.find_element(
                                                    By.XPATH, '//tr[@class="selected"]//a[@class="viewDocument"]')
                                        except:
                                            raise("Error in downloading : %s",str(traceback.format_exc()))

                                    before = set(DOWNLOAD_DIR.iterdir())
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
                                    temp_dict["Pdf Link"] = pdf_link

                                    file_type, downloaded_path = wait_for_download(before, timeout=60)
                                    if file_type != "pdf":
                                        raise ValueError(
                                            f"Expected a PDF from external documents, got {file_type}")

                                    final_pdf_path = pdf_dir / f"{pdf_name}.pdf"
                                    shutil.copy2(downloaded_path, final_pdf_path)
                                    downloaded_path.unlink(missing_ok=True)
                                    logger.info("External-documents PDF saved: %s (case_ref=%s)", final_pdf_path, case_ref)
                                    item_ext = extract_via_pdf(final_pdf_path, case_ref=case_ref)
                                    temp_dict.update({k: v for k, v in item_ext.items() if v != ''})

                                    count += 1
                                    link_row["Extracted"] = "PDF Data"

                                    handled = True
                                except Exception as e:
                                    logger.error("External document extraction failed for case_ref=%s: %s", case_ref, e)
                                    log_failed_item(index, url, case_ref, "external_documents", e)
                                    link_row["Extracted"] = "Not Extracted"

                            else:
                                logger.warning("No document tab found for %s (case_ref=%s)", driver.current_url, case_ref)
                                log_failed_item(index, url, case_ref, "no_documents_tab", "no document tab present")
                                link_row["PDF Data"] = "Not Extracted"

                            if len(driver.window_handles) > 1:
                                driver.close()

                            safe_switch_to_window(0)

                            results.append(dict(temp_dict))
                            
                            break

                        if not handled:
                            if len(driver.window_handles) > 1:
                                driver.close()
                            safe_switch_to_window(0)
                        success = True 
                    except Exception as link_err:
                        link_row["Status"] = "Failed"
                        link_row["Error"] = link_err
                        logger.error("[link %s failed, skipping] case_ref=%s: %s", link_index, case_ref, link_err)
                        log_failed_item(index, url, case_ref, "link_loop", str(traceback.format_exc()))
                        if len(driver.window_handles) > 1:
                            try:
                                driver.close()
                            except Exception:
                                pass
                        try:
                            safe_switch_to_window(0)
                        except Exception:
                            pass

                    finally:
                        new_df_rows.append(link_row)
                        save_new_df()
                        logger.info("Appended the new DF")

                    link_index += 1

                    if link_index % CHECKPOINT_EVERY == 0:
                        checkpoint_results()
                if success:
                    df.at[index, 'Status'] = 'Completed'
                else:
                    df.at[index,'Status'] = 'Failed'

                break

            except Exception as e:
                logger.error("[row %s attempt %s/%s failed] url=%s error=%s", index, row_attempts, row_max_attempts, url, str(traceback.print_exc()))
                if row_attempts >= row_max_attempts:
                    save_html_by_id(index)
                    log_failed_item(index, url, None, "row_retry_exhausted", e)
                else:
                    time.sleep(3 * row_attempts)
            df.at[index, 'Status'] = 'Failed'

            checkpoint_results()

        df.to_excel(file_path, index=False)
        logger.info("=== Row %s complete ===", index)


app_typ_count = {app: 0 for app in application_types}

count_ext = 0

for val in app_typ_count.values():
    count_ext += val

logger.info("Initial application-type count total: %s", count_ext)
driver = None
win = None
try:
    open_chrome()
    main()

except Exception as e:
    logger.critical("Fatal error in main run: %s", e)
    logger.exception("Full traceback for fatal error:")
finally:
    logger.info("Final application-type counts: %s", app_typ_count)
    res_df = pd.DataFrame(results)
    # for i in range(1, 1000):
    try:
        out_path = f"Output/res_df_{datetime.today().strftime('%Y%m%d')}.csv"
        res_df.to_csv(out_path, index=False)
        save_new_df()
        logger.info("Results saved: %s (%s rows)", out_path, len(res_df))
        # break
    except Exception:
        pass
    if driver is not None:
        try:
            driver.quit()
            logger.info("Chrome driver closed cleanly")
        except Exception as e:
            logger.error("Error while closing driver: %s", e)
    logger.info("Scraper run finished. Log file: %s", LOG_FILE)

    if win:
        win[0].close()