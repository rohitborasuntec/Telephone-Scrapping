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
import re, scrapy, time, pandas as pd,fitz
import random, zipfile, functools,shutil,os,glob
from random import randint
from pathlib import Path
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


application_types = [ "Full Planning Application", "Householder", "Full Planning Permission", "Listed Building Consent (S8 P&LBC 1990)", "Fast Track FLH (Householder App)", "Full Planning Permission (Householder)", "Householder Planning Application","Householder Application", "Full Application", "Lawful Development Certificate -Proposed", "Listed Building Consent", "Householders Extensions Prior Approval", "Lawful Development Certificate proposed", "Householder Prior Approval", "General Permitted Development - Extns", "Full planning", "Lawful Development - Proposed Use", "Full", "Outline", "Householder Planning Permission", "Planning Application", "Listed Building", "Planning Permission in Principle", "Prior Approval - Larger Household Extension", "Householder Planning Consent", "Proposed Lawful Development", "Prior Approval Larger Home Extension", "Domestic Application (Householder)", "C of Lawfulness for Proposed Use or Dev", "Residential Extensions", "Pre-Application", "Lawful Development Certificate", "Outline Application", "Full App via planning portal", "Certificate Proposed Development", "Certificate of Lawfulness (Proposed)", "Full Application (8 Weeks)", "Outline Application (8 Weeks)" ]

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

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    html_file = folder_path / f"{id}.html"
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(driver.page_source)

    print(f"📄 HTML saved: {html_file}")

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
        print("*" * 10 + "Bot Aaya" + "*" * 10)

        random_scroll(
            min_scroll=200,
            max_scroll=800,
            min_pause=0.5,
            max_pause=2,
            iterations=random.randint(1, 4))

        time.sleep(randint(3, 8))
    else:
        print("*" * 10 + "No Bot Found" + "*" * 10)


# ---------------------------------------------------------------------------
# Safe navigation / interaction helpers (used inside main loop)
# ---------------------------------------------------------------------------
# @retry(max_attempts=3, delay=3, backoff=2, exceptions=(WebDriverException, TimeoutException))
def safe_get(url, wait_ready=True, timeout=20):
    """
    Returns True if the page appears to have loaded successfully.
    Returns False for common browser/network/server failures.
    """

    try:
        # driver.set_page_load_timeout(timeout)
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

    # Chrome internal error page
    if driver.current_url.startswith("chrome-error://"):
        print(f"Chrome error page: {url}")
        return False

    title = (driver.title or "").lower()
    source = (driver.page_source or "").lower()

    # Common browser/server error pages
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

    # Blank page
    if not title and len(source.strip()) < 50:
        print(f"Blank page: {url}")
        return False

    # Wait for DOM ready
    if wait_ready:
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script(
                    "return document.readyState"
                ) == "complete"
            )
        except TimeoutException:
            print(f"document.readyState timeout: {url}")
            return False

    return True

@retry(max_attempts=4, delay=1.5, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
def safe_find(by, value, timeout=10):
    """Waits for and returns an element, retrying on stale/not-found/timeout."""
    # print(driver.current_url)
    # print(driver.title)
    # driver.save_screenshot("timeout.png")
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


# @retry(max_attempts=4, delay=1.5, backoff=2, exceptions=RETRYABLE_SELENIUM_EXC)
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
def extract_via_web():
    time.sleep(3)
    details = safe_find(By.XPATH, '//div[@class="addressCrumb"]')
    app_name = details.find_element(By.XPATH, "//th[contains(text(),'Applicant Name')]/following-sibling::td").text
    address = details.find_element(By.CLASS_NAME, 'address').text
    ref_num = details.find_element(By.CLASS_NAME, 'caseNumber').text
    desc = details.find_element(By.CLASS_NAME, 'description').text

    item_ext = {
        "Applicant Name": app_name,
        "Full Postal Address": address,
        "Planning Reference Number": ref_num,
        "Brief Project Description": desc,
        "Source" : "Web"
    }
    print("Data extracted from the website : ",item_ext)
    results.append(item_ext)
    # return item_ext

def wait_for_download(timeout=60):
    start = time.time()

    while time.time() - start < timeout:
        if list(DOWNLOAD_DIR.glob("*.crdownload")):
            time.sleep(1)
            continue

        zips = list(DOWNLOAD_DIR.glob("*.zip"))
        if zips:
            return "zip", max(zips, key=lambda p: p.stat().st_ctime)

        pdfs = list(DOWNLOAD_DIR.glob("*.pdf"))
        if pdfs:
            return "pdf", max(pdfs, key=lambda p: p.stat().st_ctime)
        time.sleep(1)

    raise TimeoutError("Download timed out")

def extract_via_pdf(pdf_path):

    doc = fitz.open(pdf_path)
    print("Openging PDF >> ")

    pdf_text = ""

    for page in doc:
        pdf_text += page.get_text()

    if pdf_text:
        print("Text found")
    else:
        print("No text in pdf")

    ref_regex = r'Planning Portal Reference:\s*(PP-\d+)'

    applicant_regex = (
        r'Applicant Details.*?'
        r'First name\s*([A-Za-z\'\-]+).*?'
        r'Surname\s*([A-Za-z\'\-]+)'
    )

    # Applicant Address
    address_regex = (
        r'Address line 1\s*(.*?)\s*'
        r'Address line 2.*?'
        r'Town/City\s*(.*?)\s*'
        r'County\s*(.*?)\s*'
        r'Country.*?'
        r'Postcode\s*([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})'
    )
    description_regex = (
        r'Description\s*'
        r'Please describe details of the proposed development or works.*?\n'
        r'(.*?)\s*'
        r'Yes\s*\nNo'
    )

    ref = re.search(ref_regex, pdf_text, re.S)
    ref_num = ref.group(1) if ref else None

    # Applicant
    m = re.search(applicant_regex, pdf_text, re.S)
    app_name = f"{m.group(1)} {m.group(2)}" if m else None

    # Address
    m = re.search(address_regex, pdf_text, re.S)
    if m:
        address = f"{m.group(1)}, {m.group(2)}, {m.group(3)} {m.group(4)}"
    else:
        address = None

    # Description
    m = re.search(description_regex, pdf_text, re.S)
    desc = m.group(1).strip() if m else None

    # res = extract_via_web()

    item_ext = {
        "Applicant Name": app_name,
        "Full Postal Address": address,
        "Planning Reference Number": ref_num,
        "Brief Project Description": desc,
        "Source" : "PDF"
    }
    results.append(item_ext)

    print("Data extracted from the pdf : ",item_ext)

@retry(max_attempts=3, delay=3, backoff=2, exceptions=(RETRYABLE_SELENIUM_EXC + (TimeoutError,)))
# def extract_pdf(pdf_new_name):

#     print("Extrating PDF..")

#     check_for_bot()

#     app_form_xpath = "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]/../td/input[@class='bulkCheck']"

#     if driver.find_elements(By.XPATH,app_form_xpath):

#         driver.find_element(By.XPATH,app_form_xpath).click()

#         time.sleep(randint(1, 3))

#         safe_click(By.ID, "downloadFiles")

#         print("ZIP DOWN")


#         zip_path = wait_for_download(timeout=60)  # should return downloaded zip path
#         print("Zip Path : ",zip_path)
#         # Extract ZIP
#         extract_dir = os.path.dirname(zip_path)
#         print("extract_dir : ",extract_dir)

#         with zipfile.ZipFile(zip_path, "r") as zf:
#             zf.extractall(extract_dir)

#         # Find extracted PDF
#         pdf_files = glob.glob(os.path.join(extract_dir, "*.pdf"))

#         if not pdf_files:
#             raise FileNotFoundError("No PDF found inside ZIP")

#         extracted_pdf = max(pdf_files, key=os.path.getmtime)

#         final_pdf_path = os.path.join(extract_dir, pdf_new_name)
#         shutil.move(extracted_pdf, final_pdf_path)

#         print(f"PDF extracted: {final_pdf_path}")

#         extract_via_pdf(final_pdf_path)

#     else:
#         print("Application not found")
#         return None
# def extract_pdf(pdf_new_name):

#     print("Extrating PDF..")

#     check_for_bot()

#     app_form_xpath = "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]/../td/input[@class='bulkCheck']"
    
#     if driver.find_elements(By.XPATH, app_form_xpath):

#         try:
#             driver.find_element(By.XPATH, app_form_xpath).click()

#             time.sleep(randint(1, 3))

#             safe_click(By.ID, "downloadFiles")

#             print("ZIP DOWN")
#             # breakpoint()
#             zip_path = wait_for_download(timeout=60)
#             print("Zip Path :", zip_path)
#             print("Zip Path Type :", type(zip_path))

#         # Handle tuple return like:
#         # ('zip', WindowsPath('PDF/files/26-W-00029.zip'))
#             # if isinstance(zip_path, tuple):
#             #     print("Tuple detected, extracting actual path...")
#             #     print("Tuple contents:", zip_path)
#             #     zip_path = zip_path[1]
#         # except:
#         #     breakpoint()
#             # zip_path = str(zip_path)

#             # print("Normalized Zip Path :", zip_path)

#             # # Extract ZIP
#             # extract_dir = os.path.dirname(zip_path)
#             # print("extract_dir :", extract_dir)

#             # with zipfile.ZipFile(zip_path, "r") as zf:
#             #     zf.extractall(extract_dir)

#             # print("ZIP extracted successfully")

#             # # Find extracted PDF
#             # pdf_files = glob.glob(os.path.join(extract_dir, "*.pdf"))

#             # print("PDF files found:", pdf_files)

#             # if not pdf_files:
#             #     raise FileNotFoundError(f"No PDF found inside ZIP: {zip_path}")

#             # extracted_pdf = max(pdf_files, key=os.path.getmtime)

#             # print("Latest extracted PDF:", extracted_pdf)
            
#             # final_pdf_path = os.path.join(extract_dir, pdf_new_name)

#             # # Remove existing file if present
#             # if os.path.exists(final_pdf_path):
#             #     os.remove(final_pdf_path)

#             # shutil.move(extracted_pdf, final_pdf_path)

#             # print(f"PDF extracted: {final_pdf_path}")

#             # extract_via_pdf(final_pdf_path)

#             # return final_pdf_path

#         except:
#             print("Breakpoint due to errorrrr.....")
#             breakpoint()
#     else:
#         print("Application not found")
#         return None


# from pathlib import Path
# import shutil
# import zipfile
# import time
# from random import randint
# from selenium.webdriver.common.by import By

def extract_pdf(pdf_new_name):

    print("Extrating PDF..")

    check_for_bot()

    app_form_xpath = "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]/../td/input[@class='bulkCheck']"

    if driver.find_elements(By.XPATH, app_form_xpath):

        try:
            driver.find_element(By.XPATH, app_form_xpath).click()

            time.sleep(randint(1, 3))

            safe_click(By.ID, "downloadFiles")

            print("ZIP DOWN")

            # file_ = Path(wait_for_download(timeout=60))

            # print("Zip Path :", file_)
            # print("Zip Path Type :", type(file_))

            file_type, file_ = wait_for_download(timeout=60)
            file_ = Path(file_)

            print("File Type :", file_type)
            print("File Path :", file_)

            if file_type == "zip":

                zip_new_name = f"{pdf_new_name}.zip"

                renamed_zip_path = file_.with_name(zip_new_name)
                file_.rename(renamed_zip_path)

                print("Renamed ZIP :", renamed_zip_path)

                zip_copy_path = zip_dir / zip_new_name
                shutil.copy2(renamed_zip_path, zip_copy_path)

                print("Copied ZIP to :", zip_copy_path)

                with zipfile.ZipFile(zip_copy_path, "r") as zip_ref:
                    zip_ref.extractall(pdf_dir)

                print("ZIP extracted to :", pdf_dir)

                renamed_zip_path.unlink(missing_ok=True)

                print("Deleted temp ZIP :", renamed_zip_path)

                return zip_copy_path

            else:

                pdf_name = f"{pdf_new_name}.pdf"

                renamed_pdf_path = file_.with_name(pdf_name)
                file_.rename(renamed_pdf_path)

                pdf_copy_path = pdf_dir / pdf_name
                shutil.copy2(renamed_pdf_path, pdf_copy_path)

                print("Copied PDF to :", pdf_copy_path)

                renamed_pdf_path.unlink(missing_ok=True)

                print("Deleted temp PDF :", renamed_pdf_path)

                return pdf_copy_path

        except Exception as e:
            print(f"[extract_pdf] Error: {e}")
            return None

    else:
        print("Application not found")
        return None


@retry(max_attempts=3, delay=2, backoff=2, exceptions=(FileNotFoundError, ValueError, OSError))
def file_rename(file_new_name):
    # Ignore temporary download files
    files = [
        f for f in DOWNLOAD_DIR.iterdir()
        if f.is_file() and f.suffix.lower() not in [".crdownload", ".part", ".tmp"]
    ]

    if not files:
        raise FileNotFoundError("No downloaded files found.")

    # Get the most recently created/modified file
    latest_file = max(files, key=lambda f: f.stat().st_ctime)

    # Preserve the original extension
    new_name = DOWNLOAD_DIR / f"{file_new_name}{latest_file.suffix}"

    latest_file.rename(new_name)

    print(f"Renamed {latest_file.name} -> {new_name.name}")

    return new_name

def main():

    df = pd.read_excel(file_path)

    for index, row in df.iterrows():
        url = "https://" + row["Url"] if "https://" not in row["Url"] else row["Url"]

        count = 1

        if "weeklyList" not in url:
            url = url.split("uk/")[0] + "uk/online-applications/search.do?action=weeklyList"

        print("Redirecting to the url : ", url)

        row_attempts = 0
        row_max_attempts = 3
        while row_attempts < row_max_attempts:
            row_attempts += 1
            try:
                if not safe_get(url):
                    print("Skipping this URL...")
                    break
                check_for_bot()
                page_source = driver.page_source
                
                if "page can’t be found" in page_source.lower() or "site can’t be reached" in page_source.lower()  or "request " in page_source.lower():
                    print("Page not found : ", url)
                    save_html_by_id(index)
                    break  # no point retrying a genuinely missing page

                safe_select_by_text(By.ID, "week", "29 Jun 2026")
                print("Date Changed")

                safe_click(By.XPATH, '//input[@value="DC_Decided"]')
                safe_click(By.XPATH, '//input[@value="Search"]')
                print("Search Page Redirection")

                time.sleep(3)

                safe_select_by_text(By.ID, 'resultsPerPage', "100") \
                    if False else None  # visible-text selector not applicable here; use value selector below

                select_element = safe_find(By.ID, 'resultsPerPage')
                Select(select_element).select_by_value("100")
                print("100 Pages selection")

                safe_click(By.XPATH, '//input[@value="Go"]')
                go = "GO," * 10
                print(go.rstrip())

                time.sleep(randint(1, 3))

                element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                n = len(element_links)
                i = 0
                print(f"Total links found : {n}")

                while i < n:
                    try:
                        element_links = driver.find_elements(By.XPATH, '//li[@class="searchresult"]/a[not(@href="#")]')
                        element = element_links[i]
                        element.send_keys(Keys.CONTROL + Keys.RETURN)

                        safe_switch_to_window(-1)

                        if "page can’t be found" in page_source:
                            print("Page not found")
                            print(url)
                            save_html_by_id(index)
                            i += 1
                            continue

                        check_for_bot()

                        safe_click(By.XPATH, "//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'information')]")

                        time.sleep(3)

                        try:
                            for j in range(1, 5):
                                if driver.find_elements(By.XPATH, "//th[contains(text(),'Application Type')]"):
                                    app_type = driver.find_element(
                                        By.XPATH,
                                        "//th[contains(text(),'Application Type')]/following-sibling::td"
                                    ).text
                                    print(app_type)
                                    if app_type in application_types:
                                        print("Found : ", app_type)
                                        
                                        try:
                                            extract_via_web()
                                        except Exception as e:
                                            print("Ex in extraction : ", e)

                                        app_typ_count[app_type] += 1
                                        pdf_name = f'{index}_{count}.pdf'
                                            #                                     elif driver.find_elements(By.XPATH , '//a[@data-bind="click: OpenDocument, href: Link"]//div[contains(text(),"Application Form")]'):
                                            # safe_click(By.XPATH , '//a[@data-bind="click: OpenDocument, href: Link"]//div[contains(text(),"Application Form")]')
                                            # time.sleep(5)
                                            # print("PDF SAVED Suceessfully")
                                            # try:
                                            #     file_rename(pdf_name)
                                            # except Exception as e:
                                            #     print(f"[pdf_rename failed after retries] {e}")
                                            # count += 1
                                        if driver.find_elements(
                                            By.XPATH,
                                            "//a[@id='tab_documents']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'document')]"
                                        ):
                                            doc_link = safe_find(By.XPATH, '//a[@id="tab_documents"]').get_attribute("href")
                                            print(f"Going to {doc_link}")
                                            safe_get(doc_link)
                                            check_for_bot()
                                            try:
                                                extract_pdf(pdf_name)

                                            except Exception as e:
                                                print(f"[extract_pdf failed after retries] {e}")
                                            count += 1

                                        elif driver.find_elements(
                                            By.XPATH,
                                            "//a[@id='tab_externalDocuments']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'documents')]"
                                        ):
                                            driver.find_element(
                                                By.XPATH,
                                                "//a[@id='tab_externalDocuments']/span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'documents')]"
                                            ).click()

                                            time.sleep(randint(3, 6))
                                            safe_click(By.XPATH, '//a[contains(text(),"documents")]')
                                            driver.close()
                                            safe_switch_to_window(-1)
                                            check_for_bot()
                                            try:
                                                try:
                                                    select_element = driver.find_element(By.XPATH, '//select[@name="searchResult_length"]')
                                                    Select(select_element).select_by_value("100")
                                                except:
                                                    print("no dropdown..")
                                                    try:
                                                        driver.find_element(By.XPATH , '//a[@data-bind="click: OpenDocument, href: Link"]//div[contains(text(),"Application Form")]')
                                                    except:
                                                        driver.find_element(By.XPATH , "//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'application form')]")
                                                        try:
                                                            driver.find_element(By.XPATH,"//a[contains(text(),'Download')]")
                                                        except:    
                                                            driver.find_element(By.XPATH, '//tr[@class="selected"]//a[@class="viewDocument"]')

                                                time.sleep(5)
                                                try:
                                                    file_rename(pdf_name)
                                                    print("PDF SAVED Suceessfully")
                                                except Exception as e:
                                                    print(f"[pdf_rename failed after retries] {e}")
                                                count += 1
                                            except:
                                                print("PDF not found : ",driver.current_url)
                                        else:
                                            print("NOT FOUND FOR ", driver.current_url)

                                    if len(driver.window_handles) > 1:
                                        driver.close()
                                    safe_switch_to_window(0)
                                    break
                            else:
                                print("Something's wrong .... retyinggg...", {j + 1})
                                time.sleep(randint(3,5))
                        except Exception as e:
                            print("Exception : ", e)
                            if len(driver.window_handles) > 1:
                                driver.close()
                            safe_switch_to_window(0)

                    except Exception as link_err:
                        # One bad link shouldn't kill the whole listing pass.
                        print(f"[link {i} failed, skipping] {link_err}")
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

                break  # row processed successfully, exit retry loop

            except Exception as e:
                print(f"[row {index} attempt {row_attempts}/{row_max_attempts} failed] {e}")
                print(url)
                if row_attempts >= row_max_attempts:
                    save_html_by_id(index)
                else:
                    time.sleep(3 * row_attempts)

app_typ_count = {app: 0 for app in application_types}

try:
    driver = open_chrome()
    main()
except Exception as e:
    print(e)
finally:
    print(app_typ_count)
    res_df = pd.DataFrame(results)
    for i in range(2, 1000):
        try:
            res_df.to_csv(f"res_df{i}.csv", index=False)
            print("Res_DF saved")
            break
        except:
            pass
    driver.quit()