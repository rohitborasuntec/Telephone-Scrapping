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
from random import randint
from pathlib import Path
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


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

zip_dir = Path(r"Data\zip")
zip_dir.mkdir(parents=True, exist_ok=True)

DOWNLOAD_DIR = Path(r"Data\temp_dir")
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

pdf_dir = Path(r"Data\pdf")
pdf_dir.mkdir(parents=True, exist_ok=True)

res_col = [
    "Applicant Name", "Full Postal Address", "Planning Reference Number", "Brief Project Description"
]

results = []
RESULTS_CSV = "results_checkpoint.csv"
FAILED_CSV = "failed_items.csv"
CHECKPOINT_EVERY = 10  # rows

driver = None  # global, guarded at shutdown

# ---------------------------------------------------------------------------
# Generic retry decorator
# ---------------------------------------------------------------------------
def retry(max_attempts=3, delay=2, backoff=2, exceptions=(Exception,), on_retry=None):
    """
    Retries the wrapped function on the given exception types.

    max_attempts : total number of tries (including the first one)
    delay        : seconds to wait before the first retry
    backoff      : multiplier applied to `delay` after each failed attempt
    exceptions   : tuple of exception classes that should trigger a retry
    on_retry     : optional callback(*args, **kwargs) run before each retry
                   (e.g. to refresh the page or reset state)
    """
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
                        print(f"[RETRY FAILED] {func.__name__} gave up after "
                              f"{max_attempts} attempts: {type(e).__name__}: {e}")
                        raise
                    print(f"[RETRY {attempt}/{max_attempts}] {func.__name__} "
                          f"raised {type(e).__name__}: {e}. Retrying in {current_delay}s...")
                    if on_retry:
                        try:
                            on_retry(*args, **kwargs)
                        except Exception as cb_err:
                            print(f"[on_retry callback error] {cb_err}")
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

file_path = r'c:\Users\Rohit\Documents\Telephone Scrap Project\Url.xlsx'


# ---------------------------------------------------------------------------
# Failure logging (so nothing fails silently)
# ---------------------------------------------------------------------------
def log_failed_item(row_index, url, case_ref, stage, error):
    """
    Append a structured failure record. This is the single place every
    'something went wrong and we could not extract data' path reports to,
    so a run's failures are auditable afterwards instead of only living
    in console scrollback.
    """
    is_new = not Path(FAILED_CSV).exists()
    with open(FAILED_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(["timestamp", "row_index", "url", "case_ref", "stage", "error"])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            row_index, url, case_ref, stage, str(error)
        ])
    print(f"[LOGGED FAILURE] row={row_index} stage={stage} error={error}")


def checkpoint_results():
    """Write current results to disk so a crash mid-run doesn't lose everything."""
    if not results:
        return
    try:
        pd.DataFrame(results).to_csv(RESULTS_CSV, index=False)
    except Exception as e:
        print(f"[checkpoint_results failed] {e}")


# ---------------------------------------------------------------------------
# Browser setup
# ---------------------------------------------------------------------------
def open_chrome():

    options = uc.ChromeOptions()

    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }

    print(str(DOWNLOAD_DIR.resolve()))
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-background-timer-throttling")
    options.add_experimental_option("prefs", prefs)

    @retry(max_attempts=3, delay=5, backoff=2, exceptions=(WebDriverException,))
    def _launch():
        drv = uc.Chrome(
            options=options,
            headless=False,
            version_main=150
        )
        return drv

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

    return driver


def save_html_by_id(id, base_dir="script_html_pages"):
    id = str(id + 1)
    folder_path = Path(base_dir) / id
    folder_path.mkdir(parents=True, exist_ok=True)

    html_file = folder_path / f"{id}.html"
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(driver.page_source)

    print(f"HTML saved: {html_file}")

    return html_file


def random_scroll(min_scroll=200, max_scroll=800, min_pause=0.5, max_pause=2, iterations=None):

    if iterations is None:
        iterations = randint(5, 15)

    for _ in range(iterations):
        print("Scrolling ... ", _)
        direction = 1 if random.random() < 0.8 else -1

        pixels = random.randint(min_scroll, max_scroll) * direction

        driver.execute_script(
            "window.scrollBy({top: arguments[0], behavior: 'smooth'});",
            pixels
        )

        time.sleep(random.uniform(min_pause, max_pause))


def check_for_bot():
    time.sleep(randint(3, 6))

    page_source = driver.page_source

    if "One moment, we're checking you're not a bot." in page_source or "Performing security function" in page_source:
        print("*" * 10 + " Bot check detected " + "*" * 10)

        random_scroll(
            min_scroll=200,
            max_scroll=800,
            min_pause=0.5,
            max_pause=2,
            iterations=random.randint(1, 4))

        time.sleep(randint(3, 8))
    else:
        print("*" * 10 + " No bot check found " + "*" * 10)


# ---------------------------------------------------------------------------
# Safe navigation / interaction helpers
# ---------------------------------------------------------------------------
def safe_get(url, wait_ready=True, timeout=20):
    """
    Returns True if the page appears to have loaded successfully.
    Returns False for common browser/network/server failures.
    """
    try:
        driver.get(url)

    except TimeoutException:
        print(f"Page load timed out: {url}")
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
            print(f"Browser/network error while opening {url}")
            return False

        raise

    if driver.current_url.startswith("chrome-error://"):
        print(f"Chrome error page: {url}")
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
        "just a moment",            # Cloudflare
        "checking your browser",    # Cloudflare
    ]

    if any(err in title or err in source for err in error_strings):
        print(f"Website returned an error page: {url}")
        return False

    if not title and len(source.strip()) < 50:
        print(f"Blank page: {url}")
        return False

    if wait_ready:
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            print(f"document.readyState timeout: {url}")
            return False

    return True


@retry(max_attempts=4, delay=1.5, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def safe_find(by, value, timeout=10):
    """Waits for and returns an element, retrying on stale/not-found/timeout."""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def safe_click(by, value, timeout=10):
    """Waits for an element to be clickable and clicks it, retrying on failure."""
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )
    el.click()
    return el


@retry(max_attempts=3, delay=1.5, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def safe_select_by_text(by, value, text, timeout=10):
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    Select(el).select_by_visible_text(text)


@retry(max_attempts=3, delay=1, backoff=2, exceptions=(IndexError, WebDriverException))
def safe_switch_to_window(index=-1):
    driver.switch_to.window(driver.window_handles[index])


# ---------------------------------------------------------------------------
# Extraction / download logic
# ---------------------------------------------------------------------------

@retry(max_attempts=3, delay=2, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def extract_via_web(case_ref=None):
    time.sleep(3)
    details = safe_find(By.XPATH, '//div[@class="addressCrumb"]')
    app_name = details.find_element(By.XPATH, "//th[contains(text(),'Applicant Name')]/following-sibling::td").text
    address = details.find_element(By.CLASS_NAME, 'address').text
    ref_num = details.find_element(By.CLASS_NAME, 'caseNumber').text
    desc = details.find_element(By.CLASS_NAME, 'description').text

    item_ext = {
        "Applicant Name": app_name,
        "Address": address,
        "Reference": ref_num,
        "Proposal": desc,
        "Case Ref (internal)": case_ref,
    }
    print("Data extracted from the website : ", item_ext)
    # results.append(item_ext)
    return item_ext


def wait_for_download(before_files, timeout=60, stable_checks=2, poll=1.0):
    """
    Watches DOWNLOAD_DIR for a NEW file (relative to `before_files`, a set of
    Path objects captured immediately before triggering the download) and
    waits for its size to stop changing before returning it.

    This avoids the old bug where any leftover file already sitting in
    DOWNLOAD_DIR (from a prior failed/incomplete run) could be picked up
    and mistaken for the current download.

    Raises TimeoutError if no new, stable file shows up in time.
    Raises ValueError if the new file is neither .zip nor .pdf.
    """
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
            # file got renamed/moved mid-check; reset and keep waiting
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
            if suffix == ".pdf":
                return "pdf", candidate
            raise ValueError(f"Unexpected downloaded file type: {candidate.name}")

        time.sleep(poll)

    raise TimeoutError(f"No new stable download appeared in {DOWNLOAD_DIR} within {timeout}s")


def extract_via_pdf(pdf_path, case_ref=None):

    doc = fitz.open(pdf_path)
    print("Opening PDF >> ", pdf_path)

    pdf_text = ""
    for page in doc:
        pdf_text += page.get_text()
    doc.close()

    pdf_text = re.sub(r'\r\n?', '\n', pdf_text)
    pdf_text = re.sub(r'[ \t]+', ' ', pdf_text)


    if pdf_text:
        print("Text found")
    else:
        print("No text in pdf")
        raise ValueError(f"No extractable text in PDF: {pdf_path}")

    # ref_regex = r'Planning Portal Reference:\s*(PP-\d+)'

    # applicant_regex = (
    #     r'Applicant Details.*?'
    #     r'First name\s*([A-Za-z\'\-]+).*?'
    #     r'Surname\s*([A-Za-z\'\-]+)'
    # )

    address_regex = (
        r'Address line 1\s*(.*?)\s*'
        r'Address line 2.*?'
        r'Town/City\s*(.*?)\s*'
        r'County\s*(.*?)\s*'
        r'Country.*?'
        r'Postcode\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})'
    )
    # description_regex = (
    #     r'Description\s*'
    #     r'Please describe details of the proposed development or works.*?\n'
    #     r'(.*?)\s*'
    #     r'Yes\s*\nNo'
    # )

    # ref = re.search(ref_regex, pdf_text, re.S)
    # ref_num = ref.group(1) if ref else None
    # is_same = False

    # m = re.search(applicant_regex, pdf_text, re.S)
    # app_name = f"{m.group(1)} {m.group(2)}" if m else None

    m = re.search(address_regex, pdf_text, re.S)
    address = f"{m.group(1)}, {m.group(2)}, {m.group(3)} {m.group(4)}" if m else None

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

        print("Agent:", agent_name)
        print("Company:", company_name)
    
    else:
        agent_name = ""
        print("Agent name not found")

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

        # Only create applicant name if both first and surname exist
        if first and surname:
            applicant_name = " ".join(filter(None, [title, first, surname]))

    print(applicant_name)

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
        r"Postcode\s*\n([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})",
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

    print(site_address)   

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
        print(agent_address)

    else:
        agent_address = ""
        print("Agent address not found")

    # m = re.search(description_regex, pdf_text, re.S)
    # desc = m.group(1).strip() if m else None

    is_same = True if agent_name == applicant_name else False

    item_ext = {
        "Address":site_address,
        "Applicant Name": applicant_name,
        "Applicant address": address,
        "Agent name" : agent_name,
        "Agent Address" : agent_address,
        "Applicant = Agent" : is_same,
        "Application Form PDF Available" : "Yes",
        # "Planning Reference Number": ref_num,
        # "Brief Project Description": desc,
        # "Source": "PDF",
        "Case Ref (internal)": case_ref,
        "Source File": str(pdf_path),
    }
    results.append(item_ext)

    print("Data extracted from the pdf : ", item_ext)
    return item_ext


@retry(max_attempts=3, delay=3, backoff=2, exceptions=(RETRYABLE_SELENIUM_EXC + (TimeoutError, ValueError)))
def download_and_extract(pdf_new_name, case_ref=None):
    """
    Single entry point for 'click the application-form checkbox, download
    whatever the site gives us (zip or pdf), get an actual PDF out of it,
    and run text extraction on it'.

    Rules enforced here (per requirements):
      - The download is only considered successful if a NEW file (not a
        leftover from a previous case) is found and its size is stable.
      - A file is only renamed/moved AFTER that success check passes.
      - If unzip fails, or no PDF is found inside a zip, or extraction
        fails, an exception is raised rather than silently continuing -
        the caller is responsible for logging it and moving to the next
        item. Nothing is left "half done" and treated as success.
    """
    check_for_bot()

    app_form_xpath = (
        "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
        "'abcdefghijklmnopqrstuvwxyz'), 'application form')]"
        "/../td/input[@class='bulkCheck']"
    )

    if not driver.find_elements(By.XPATH, app_form_xpath):
        print("Application form checkbox not found - nothing to download")
        return None

    try:
        pdf_link = driver.find_element(By.XPATH,"//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]/../td/input[@class='bulkCheck']/../..//a").get_attribute('href')
    except:
        pdf_link = ""

    driver.find_element(By.XPATH, app_form_xpath).click()
    time.sleep(randint(1, 3))

    before = set(DOWNLOAD_DIR.iterdir())
    safe_click(By.ID, "downloadFiles")
    print("Download triggered, waiting for file...")

    # Raises TimeoutError / ValueError on failure - caller must handle.
    file_type, downloaded_path = wait_for_download(before, timeout=60)
    print(f"Confirmed new download: type={file_type} path={downloaded_path}")

    if file_type == "zip":
        zip_copy_path = zip_dir / f"{pdf_new_name}.zip"
        shutil.copy2(downloaded_path, zip_copy_path)
        downloaded_path.unlink(missing_ok=True)
        print(f"Copied ZIP to: {zip_copy_path}")

        try:
            with zipfile.ZipFile(zip_copy_path, "r") as zf:
                zf.extractall(pdf_dir)
        except zipfile.BadZipFile as e:
            # explicit exception, not a silent None-return
            raise ValueError(f"Could not unzip {zip_copy_path}: {e}")

        extracted_pdfs = list(pdf_dir.glob("*.pdf"))
        if not extracted_pdfs:
            raise FileNotFoundError(f"No PDF found inside ZIP: {zip_copy_path}")

        newest_extracted = max(extracted_pdfs, key=lambda p: p.stat().st_mtime)
        final_pdf_path = pdf_dir / f"{pdf_new_name}.pdf"
        if final_pdf_path.exists():
            final_pdf_path.unlink()
        newest_extracted.rename(final_pdf_path)
        print(f"PDF extracted from ZIP: {final_pdf_path}")

    elif file_type == "pdf":
        final_pdf_path = pdf_dir / f"{pdf_new_name}.pdf"
        shutil.copy2(downloaded_path, final_pdf_path)
        downloaded_path.unlink(missing_ok=True)
        print(f"Copied PDF to: {final_pdf_path}")

    else:
        # wait_for_download already guards this, but keep the check explicit here too
        raise ValueError(f"Unsupported downloaded file type: {file_type}")

    # This is the step that was previously missing: actually parse the PDF.
    item_ext = extract_via_pdf(final_pdf_path, case_ref=case_ref)
    item_ext["Pdf Link"] = pdf_link

    return item_ext


@retry(max_attempts=3, delay=2, backoff=2, exceptions=(FileNotFoundError, ValueError, OSError))
def file_rename(file_new_name, before_files=None):
    """
    Kept for the 'external documents' flow where the site doesn't go through
    the bulk-checkbox + downloadFiles button. Uses the same 'only accept a
    NEW file' rule as wait_for_download - if before_files isn't supplied it
    falls back to scanning the whole directory (legacy behaviour), but
    callers should now always pass before_files.
    """
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

    print(f"Renamed {latest_file.name} -> {new_name.name}")
    return new_name


def main():

    df = pd.read_excel(file_path)
    
    df = df.loc[df['Status'] != 'Completed']

    print(f"Only {df.shape[0]} Pending")

    for index, row in df.iterrows():

        print(f"{index} is running ..... ")

        url = "https://" + row["Url"] if "https://" not in row["Url"] else row["Url"]
        count = 1

        websites = row["Url"]

        if "weeklyList" not in url:
            url = url.split("uk/")[0] + "uk/online-applications/search.do?action=weeklyList"

        print("Redirecting to the url : ", url)

        row_attempts = 0
        row_max_attempts = 3

        temp_dict = {
            "S.No." : count,
            "Websites":websites,
            "Url" : "",
            "Reference" : "",
            "Address":"",
            "Proposal":"",
            "Application Type":"",
            "Application Form PDF Available" : "",
            "Applicant = Agent" : "",
            "Agent name" : "",
            "Agent Address" : "",
            "Applicant Name" : "",
            "Applicant address" : "",
            "Pdf Link" : ""            
        }

        while row_attempts < row_max_attempts:
            row_attempts += 1
            try:
                if not safe_get(url):
                    print("Skipping this URL...")
                    log_failed_item(index, url, None, "safe_get", "page failed to load")
                    break

                check_for_bot()
                page_source = driver.page_source

                if ("page can't be found" in page_source.lower()
                        or "site can't be reached" in page_source.lower()
                        or "page cant be found" in page_source.lower()):
                    print("Page not found : ", url)
                    save_html_by_id(index)
                    log_failed_item(index, url, None, "listing_page", "page not found")
                    break  # no point retrying a genuinely missing page

                safe_select_by_text(By.ID, "week", "29 Jun 2026")
                print("Date Changed")

                safe_click(By.XPATH, '//input[@value="DC_Decided"]')
                safe_click(By.XPATH, '//input[@value="Search"]')
                print("Search Page Redirection")

                time.sleep(3)

                select_element = safe_find(By.ID, 'resultsPerPage')
                Select(select_element).select_by_value("100")
                print("100 results per page selected")

                safe_click(By.XPATH, '//input[@value="Go"]')
                time.sleep(randint(1, 3))

                element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                n = len(element_links)
                i = 0
                print(f"Total links found : {n}")

                while i < n:
                    case_ref = None
                    try:
                        element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                        element = element_links[i]
                        case_ref = element.text.strip() if element.text else f"row{index}_link{i}"
                        element.send_keys(Keys.CONTROL + Keys.RETURN)

                        safe_switch_to_window(-1)
                        # refresh page_source for the NEW tab - previously this stayed
                        # bound to the listing page and never actually checked the
                        # detail page's content.
                        detail_page_source = driver.page_source

                        if ("page can't be found" in detail_page_source.lower()
                                or "page cant be found" in detail_page_source.lower()):
                            print("Page not found for link", i)
                            save_html_by_id(index)
                            log_failed_item(index, url, case_ref, "detail_page", "page not found")
                            if len(driver.window_handles) > 1:
                                driver.close()
                            safe_switch_to_window(0)
                            i += 1
                            continue

                        check_for_bot()

                        safe_click(By.XPATH,
                                   "//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                                   "'abcdefghijklmnopqrstuvwxyz'), 'information')]")
                        time.sleep(3)

                        handled = False
                        for j in range(1, 5):
                            if not driver.find_elements(By.XPATH, "//th[contains(text(),'Application Type')]"):
                                print(f"Application Type not visible yet, retry {j}/4 ...")
                                time.sleep(randint(3, 5))
                                continue

                            app_type = driver.find_element(
                                By.XPATH,
                                "//th[contains(text(),'Application Type')]/following-sibling::td"
                            ).text
                            print(app_type)

                            if app_type not in application_types:
                                print("Application type not in target list, skipping:", app_type)
                                handled = True
                                break

                            print("Found matching type : ", app_type)
                            temp_dict["Url"] = driver.current_url
                            temp_dict["Application Type"] = app_type
                            try:
                                web_data = extract_via_web(case_ref=case_ref)

                                temp_dict.update(web_data)
                                
                                # results.append(web_data)
                            except Exception as e:
                                # web_data = 
                                print("Web extraction failed : ", e)
                                log_failed_item(index, url, case_ref, "extract_via_web", e)

                            app_typ_count[app_type] += 1
                            pdf_name = f'{index}_{count}'

                            if driver.find_elements(
                                By.XPATH,
                                "//a[@id='tab_documents']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                                "'abcdefghijklmnopqrstuvwxyz'), 'document')]"
                            ):
                                doc_link = safe_find(By.XPATH, '//a[@id="tab_documents"]').get_attribute("href")
                                print(f"Going to {doc_link}")
                                safe_get(doc_link)
                                check_for_bot()
                                
                                try:
                                    item_ext = download_and_extract(pdf_name, case_ref=case_ref)
                                    temp_dict.update(item_ext)
                                except Exception as e:
                                    print(f"[download_and_extract failed after retries] {e}")
                                    log_failed_item(index, url, case_ref, "download_and_extract", e)
                                count += 1

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

                                time.sleep(randint(3, 6))
                                safe_click(By.XPATH, '//a[contains(text(),"documents")]')
                                driver.close()
                                safe_switch_to_window(-1)
                                check_for_bot()

                                try:
                                    try:
                                        select_element = driver.find_element(
                                            By.XPATH, '//select[@name="searchResult_length"]')
                                        Select(select_element).select_by_value("100")
                                    except NoSuchElementException:
                                        print("No results-per-page dropdown on external documents page")
                                        time.sleep(5)
                                    try:
                                        driver.find_element(
                                            By.XPATH,
                                            '//a[@data-bind="click: OpenDocument, href: Link"]'
                                            '//div[contains(text(),"Application Form")]'
                                        ).click()
                                    except NoSuchElementException:
                                        driver.find_element(
                                            By.XPATH,
                                            "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                                            "'abcdefghijklmnopqrstuvwxyz'), 'application form')]"
                                        ).click()
                                    try:
                                        time.sleep(5)
                                        driver.find_element(By.XPATH, "//a[contains(text(),'Download')]")
                                    except NoSuchElementException:
                                        driver.find_element(
                                            By.XPATH, '//tr[@class="selected"]//a[@class="viewDocument"]')

                                    before = set(DOWNLOAD_DIR.iterdir())
                                    driver.find_element(
                                        By.XPATH,
                                        '//tr[@class="selected"]//a[@class="viewDocument"] | '
                                        '//a[contains(text(),"Download")] | '
                                        '//a[@data-bind="click: OpenDocument, href: Link"]'
                                        '//div[contains(text(),"Application Form")]/..'
                                    ).click()
                                    pdf_link = ""
                                    temp_dict["Pdf Link"] = pdf_link

                                    file_type, downloaded_path = wait_for_download(before, timeout=60)
                                    if file_type != "pdf":
                                        raise ValueError(
                                            f"Expected a PDF from external documents, got {file_type}")

                                    final_pdf_path = pdf_dir / f"{pdf_name}.pdf"
                                    shutil.copy2(downloaded_path, final_pdf_path)
                                    downloaded_path.unlink(missing_ok=True)
                                    print(f"PDF saved: {final_pdf_path}")

                                    item_ext = extract_via_pdf(final_pdf_path, case_ref=case_ref)
                                    temp_dict.update(item_ext)
                                    count += 1

                                except Exception as e:
                                    print("External document extraction failed : ", e)
                                    log_failed_item(index, url, case_ref, "external_documents", e)
                            else:
                                print("No document tab found for ", driver.current_url)
                                log_failed_item(index, url, case_ref, "no_documents_tab", "no document tab present")

                            if len(driver.window_handles) > 1:
                                driver.close()
                            safe_switch_to_window(0)
                            handled = True

                            results.append(temp_dict)

                            break

                        if not handled:
                            print("Could not read Application Type after retries for link", i)
                            log_failed_item(index, url, case_ref, "application_type",
                                             "Application Type field never appeared")
                            if len(driver.window_handles) > 1:
                                driver.close()
                            safe_switch_to_window(0)

                    except Exception as link_err:
                        print(f"[link {i} failed, skipping] {link_err}")
                        log_failed_item(index, url, case_ref, "link_loop", link_err)
                        if len(driver.window_handles) > 1:
                            try:
                                driver.close()
                            except Exception:
                                pass
                        try:
                            safe_switch_to_window(0)
                        except Exception:
                            pass

                    i += 1

                    if i % CHECKPOINT_EVERY == 0:
                        checkpoint_results()

                break  # row processed successfully, exit retry loop

            except Exception as e:
                print(f"[row {index} attempt {row_attempts}/{row_max_attempts} failed] {e}")
                print(url)
                if row_attempts >= row_max_attempts:
                    save_html_by_id(index)
                    log_failed_item(index, url, None, "row_retry_exhausted", e)
                else:
                    time.sleep(3 * row_attempts)

        checkpoint_results()

        df.at[index,'Status'] = 'Completed'


app_typ_count = {app: 0 for app in application_types}

count_ext = 0

for val in app_typ_count.values():
    count_ext += val

print("Total Count : " ,count_ext)

try:
    driver = open_chrome()
    main()
except Exception as e:
    print("Fatal error in main run:", e)
    traceback.print_exc()
finally:
    print(app_typ_count)
    res_df = pd.DataFrame(results)
    for i in range(1, 1000):
        try:
            res_df.to_csv(f"Output/res_df{i}.csv", index=False)
            print("Res_DF saved")
            break
        except Exception:
            pass
    if driver is not None:
        try:
            driver.quit()
        except Exception:
            pass