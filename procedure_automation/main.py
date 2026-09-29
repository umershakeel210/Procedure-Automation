import os
import time
from datetime import datetime

from openpyxl import load_workbook, Workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException

from webdriver_manager.chrome import ChromeDriverManager

import config
from field_mapping import DEFAULT_VALUES


# =====================================================
# Folder Setup
# =====================================================

os.makedirs("logs", exist_ok=True)
os.makedirs("screenshots/failed_rows", exist_ok=True)


# =====================================================
# Excel Helper Functions
# =====================================================

def clean_value(value):
    if value is None:
        return ""

    value = str(value).strip()

    if value.lower() == "nan":
        return ""

    return value


def clean_amount(value):
    if value is None:
        return None

    value = str(value).strip()

    if value == "" or value == "-" or value.lower() == "nan":
        return None

    value = value.replace(",", "")

    try:
        return float(value)
    except ValueError:
        return None


def load_excel_rows(file_path):
    if not os.path.exists(file_path):
        raise Exception(f"Excel file not found: {file_path}")

    workbook = load_workbook(file_path)
    sheet = workbook.active

    headers = []

    for cell in sheet[1]:
        if cell.value is None:
            headers.append("")
        else:
            headers.append(str(cell.value).strip())

    rows = []

    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        row_data = {}

        for index, value in enumerate(row):
            if index < len(headers):
                column_name = headers[index]

                if column_name:
                    row_data[column_name] = value

        row_data["_excel_row_number"] = row_number

        has_data = False

        for key, value in row_data.items():
            if key != "_excel_row_number" and value not in [None, ""]:
                has_data = True
                break

        if has_data:
            rows.append(row_data)

    return rows


def write_log(log_records):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Automation Log"

    headers = [
        "Excel Row",
        "Procedure Code",
        "Procedure Name",
        "Status",
        "Message",
        "Date Time",
        "Screenshot"
    ]

    sheet.append(headers)

    for record in log_records:
        sheet.append([
            record.get("Excel Row", ""),
            record.get("Procedure Code", ""),
            record.get("Procedure Name", ""),
            record.get("Status", ""),
            record.get("Message", ""),
            record.get("Date Time", ""),
            record.get("Screenshot", "")
        ])

    workbook.save(config.LOG_FILE_PATH)


# =====================================================
# Selenium Helper Functions
# =====================================================

def wait_for_element(driver, selector, timeout=25):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
        )
    except TimeoutException:
        raise Exception(f"Element not found or not visible. Selector: {selector}")


def wait_for_present_element(driver, selector, timeout=25):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
        )
    except TimeoutException:
        raise Exception(f"Element not found in DOM. Selector: {selector}")


def clear_and_type(driver, selector, value):
    element = wait_for_element(driver, selector)

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        element
    )

    time.sleep(0.3)

    element.click()
    element.send_keys(Keys.CONTROL, "a")
    element.send_keys(Keys.BACKSPACE)

    if value is not None:
        element.send_keys(str(value))

    return element


def force_set_value(driver, selector, value):
    element = wait_for_present_element(driver, selector)

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        element
    )

    time.sleep(0.5)

    try:
        element.click()
        element.send_keys(Keys.CONTROL, "a")
        element.send_keys(Keys.BACKSPACE)
        element.send_keys(str(value))
        time.sleep(0.5)
    except Exception:
        pass

    driver.execute_script(
        """
        var el = arguments[0];
        var val = arguments[1];

        el.removeAttribute('disabled');
        el.removeAttribute('readonly');

        el.value = val;

        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
        """,
        element,
        str(value)
    )

    time.sleep(0.5)


def click_by_xpath_list(driver, xpaths, timeout=15):
    for xpath in xpaths:
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                element
            )

            time.sleep(0.5)

            try:
                element.click()
            except Exception:
                driver.execute_script("arguments[0].click();", element)

            time.sleep(1)
            return True

        except Exception:
            continue

    return False


def generate_procedure_code(index):
    number = config.START_PROCEDURE_CODE + index
    return f"{config.PROCEDURE_CODE_PREFIX}{number}"


def save_failed_screenshot(driver, row_number):
    screenshot_path = f"screenshots/failed_rows/row_{row_number}.png"

    try:
        driver.save_screenshot(screenshot_path)
    except Exception:
        screenshot_path = ""

    return screenshot_path


# =====================================================
# Browser Setup
# =====================================================

def setup_browser():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    return driver


# =====================================================
# Auto Login + Navigation
# =====================================================

def auto_login(driver):
    print("Opening login page...")
    driver.get(config.LOGIN_URL)
    time.sleep(3)

    print("Finding username field...")

    username_selectors = [
        "input[name='username']",
        "input[placeholder*='Username']",
        "input[ng-model*='user']",
        "input[type='text']",
        "input"
    ]

    username_input = None

    for selector in username_selectors:
        try:
            username_input = WebDriverWait(driver, 8).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
            )
            print(f"Username field found: {selector}")
            break
        except Exception:
            continue

    if username_input is None:
        raise Exception("Username field not found on login page.")

    username_input.click()
    username_input.send_keys(Keys.CONTROL, "a")
    username_input.send_keys(Keys.BACKSPACE)
    username_input.send_keys(config.LOGIN_USERNAME)

    print("Finding password field...")

    password_selectors = [
        "input[type='password']",
        "input[name='password']",
        "input[placeholder*='Password']",
        "input[ng-model*='password']"
    ]

    password_input = None

    for selector in password_selectors:
        try:
            password_input = WebDriverWait(driver, 8).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
            )
            print(f"Password field found: {selector}")
            break
        except Exception:
            continue

    if password_input is None:
        raise Exception("Password field not found on login page.")

    password_input.click()
    password_input.send_keys(Keys.CONTROL, "a")
    password_input.send_keys(Keys.BACKSPACE)
    password_input.send_keys(config.LOGIN_PASSWORD)

    time.sleep(1)

    print("Clicking SIGN IN button...")

    try:
        login_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button#loginButton"))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            login_button
        )

        time.sleep(0.5)

        try:
            login_button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", login_button)

        print("SIGN IN button clicked")

    except Exception:
        print("SIGN IN button direct click failed. Trying ENTER on password field...")
        password_input.send_keys(Keys.ENTER)

    print("Waiting after login...")
    time.sleep(7)


def open_create_procedure_page(driver):
    print("Opening Create Procedure page...")

    create_url = getattr(config, "CREATE_PROCEDURE_URL", "")

    if create_url:
        try:
            driver.get(create_url)
            time.sleep(5)

            wait_for_element(driver, config.PROCEDURE_NAME_INPUT, timeout=20)
            print("Create Procedure page opened by direct URL.")
            return

        except Exception:
            print("Direct Create Procedure URL failed. Trying menu navigation...")

    manage_procedure_xpaths = [
        "//*[normalize-space()='Manage Procedures']",
        "//*[contains(normalize-space(), 'Manage Procedures')]",
        "//*[normalize-space()='Manage Procedure']",
        "//*[contains(normalize-space(), 'Manage Procedure')]"
    ]

    create_procedure_xpaths = [
        "//*[normalize-space()='Create Procedure']",
        "//*[contains(normalize-space(), 'Create Procedure')]",
        "//*[normalize-space()='Add Procedure']",
        "//*[contains(normalize-space(), 'Add Procedure')]"
    ]

    print("Clicking Manage Procedures menu...")

    manage_clicked = click_by_xpath_list(
        driver,
        manage_procedure_xpaths,
        timeout=20
    )

    if not manage_clicked:
        raise Exception("Manage Procedures menu not found.")

    time.sleep(2)

    print("Clicking Create Procedure...")

    create_clicked = click_by_xpath_list(
        driver,
        create_procedure_xpaths,
        timeout=20
    )

    if not create_clicked:
        raise Exception("Create Procedure option not found.")

    time.sleep(4)

    wait_for_element(driver, config.PROCEDURE_NAME_INPUT, timeout=25)

    print("Create Procedure page opened by menu navigation.")


# =====================================================
# Dropdown Helper
# =====================================================

def select_smart_dropdown(driver, selector, value, field_name):
    value = clean_value(value)

    if value == "":
        print(f" - {field_name} skipped because value is empty")
        return

    print(f" - {field_name}: {value}")

    try:
        element = wait_for_element(driver, selector, timeout=20)
    except Exception:
        raise Exception(f"{field_name} field not found. Selector: {selector}")

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        element
    )

    time.sleep(0.5)

    tag_name = element.tag_name.lower()

    if tag_name == "select":
        try:
            Select(element).select_by_visible_text(value)
            time.sleep(0.5)
            return
        except Exception as e:
            raise Exception(
                f"{field_name} normal select failed. Value: {value}. Error: {str(e)}"
            )

    try:
        element.click()
        time.sleep(1)

        option_xpaths = [
            f"//md-option//div[normalize-space()='{value}']",
            f"//md-option//div[contains(normalize-space(), '{value}')]",
            f"//md-option//*[normalize-space()='{value}']",
            f"//md-option//*[contains(normalize-space(), '{value}')]",
            f"//*[contains(@class,'md-text') and normalize-space()='{value}']",
            f"//*[contains(@class,'md-text') and contains(normalize-space(), '{value}')]",
            f"//*[normalize-space()='{value}']",
            f"//*[contains(normalize-space(), '{value}')]"
        ]

        for xpath in option_xpaths:
            try:
                option = WebDriverWait(driver, 7).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});",
                    option
                )

                time.sleep(0.3)

                try:
                    option.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", option)

                time.sleep(0.7)
                return

            except Exception:
                continue

        raise Exception(f"{field_name} option not found. Value: {value}")

    except Exception as e:
        raise Exception(
            f"{field_name} dropdown selection failed. Selector: {selector}. Value: {value}. Error: {str(e)}"
        )


# =====================================================
# Attach Report
# =====================================================

def attach_procedure_report(driver):
    print(" - Opening Attach Report")

    main_window = driver.current_window_handle
    all_windows_before = set(driver.window_handles)

    attach_clicked = click_by_xpath_list(
        driver,
        [
            "//*[normalize-space()='ATTACH REPORT']",
            "//*[contains(normalize-space(), 'ATTACH REPORT')]",
            "//*[normalize-space()='Attach Report']",
            "//*[contains(normalize-space(), 'Attach Report')]"
        ],
        timeout=15
    )

    if not attach_clicked:
        raise Exception("ATTACH REPORT button not found.")

    time.sleep(2)

    try:
        WebDriverWait(driver, 10).until(
            lambda d: len(d.window_handles) > len(all_windows_before)
        )
        new_window = list(set(driver.window_handles) - all_windows_before)[0]
        driver.switch_to.window(new_window)
        print(" - Switched to new window")
        time.sleep(1.5)
    except Exception:
        print(" - No new window, modal in same window")

    print(" - Finding report search input...")

    search_input = None

    for selector in [
        "input#searchBar",
        "input[type='search']",
        "input[placeholder*='Search']",
        "input[placeholder*='Report']",
        "md-dialog input"
    ]:
        try:
            search_input = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            )
            print(f" - Found search input: {selector}")
            break
        except Exception:
            continue

    if search_input is None:
        raise Exception("Report search input not found.")

    print(" - Searching Procedure Report")

    driver.execute_script("arguments[0].removeAttribute('disabled');", search_input)
    driver.execute_script("arguments[0].removeAttribute('readonly');", search_input)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_input)

    time.sleep(0.5)

    search_input.click()
    search_input.send_keys(Keys.CONTROL, "a")
    search_input.send_keys(Keys.BACKSPACE)
    search_input.send_keys("Procedure")

    time.sleep(3)

    option_selected = False

    option_xpaths = [
        "//ul[contains(@id,'ul-')]//li[contains(normalize-space(), 'Form000082')]",
        "//ul[contains(@id,'ul-')]//li[contains(normalize-space(), 'Procedure Report')]",
        "//ul[contains(@id,'ul-')]//li",
        "//*[contains(@class,'md-virtual-repeat-container')]//li[contains(normalize-space(), 'Form000082')]",
        "//*[contains(@class,'md-virtual-repeat-container')]//li[contains(normalize-space(), 'Procedure Report')]",
        "//*[contains(@class,'md-virtual-repeat-container')]//li",
        "//*[contains(@class,'md-autocomplete-suggestion') and contains(normalize-space(), 'Form000082')]",
        "//*[contains(@class,'md-autocomplete-suggestion') and contains(normalize-space(), 'Procedure Report')]",
        "//*[contains(@class,'md-autocomplete-suggestion')]",
        "//li[contains(normalize-space(), 'Form000082')]",
        "//li[contains(normalize-space(), 'Procedure Report')]"
    ]

    for xpath in option_xpaths:
        try:
            option = WebDriverWait(driver, 6).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                option
            )

            time.sleep(0.4)

            try:
                option.click()
            except Exception:
                driver.execute_script("arguments[0].click();", option)

            option_selected = True
            time.sleep(1.5)
            break

        except Exception:
            continue

    if not option_selected:
        try:
            search_input.click()
            search_input.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.7)
            search_input.send_keys(Keys.ENTER)
            time.sleep(1.5)
            option_selected = True
        except Exception:
            pass

    if not option_selected:
        raise Exception("Form000082-Procedure Report option select nahi ho saka.")

    print(" - Clicking report DONE")

    done_clicked = click_by_xpath_list(
        driver,
        [
            "//button[normalize-space()='DONE']",
            "//button[contains(normalize-space(),'DONE')]",
            "//*[@ng-click=\"answer('done')\"]",
            "//md-dialog//button[last()]"
        ],
        timeout=10
    )

    if not done_clicked:
        for selector in [
            "button[ng-click=\"answer('done')\"]",
            "button.md-primary.md-button",
            "button.md-primary",
            "md-dialog-actions button.md-primary"
        ]:
            try:
                done_btn = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                driver.execute_script("arguments[0].click();", done_btn)
                done_clicked = True
                time.sleep(2)
                break
            except Exception:
                continue

    if not done_clicked:
        raise Exception("Report DONE button click nahi ho saka.")

    try:
        if len(driver.window_handles) > 1:
            driver.switch_to.window(main_window)
            time.sleep(1)
    except Exception:
        pass

    print(" - Report attached successfully")


# =====================================================
# Components Open
# =====================================================

def open_components_section(driver):
    print(" - Opening Components section")

    clicked = click_by_xpath_list(
        driver,
        [
            "//*[normalize-space()='CREATE COMPONENTS']",
            "//*[contains(normalize-space(), 'CREATE COMPONENTS')]",
            "//*[normalize-space()='Create Components']",
            "//*[contains(normalize-space(), 'Create Components')]",
            "//*[contains(normalize-space(), 'CREATE COMPONENT')]",
            "//*[contains(normalize-space(), 'Create Component')]"
        ],
        timeout=15
    )

    if not clicked:
        raise Exception("CREATE COMPONENTS button not found or not clickable.")

    time.sleep(2)
    print(" - Components section opened")


# =====================================================
# Main Form Helpers
# =====================================================

def get_excel_or_default(row, excel_column, default_key):
    excel_value = clean_value(row.get(excel_column, ""))

    if excel_value != "":
        return excel_value

    return DEFAULT_VALUES.get(default_key, "")


def get_procedure_code(row, index):
    excel_code = clean_value(row.get("Procedure Code", ""))

    if excel_code != "":
        return excel_code

    return generate_procedure_code(index)


# =====================================================
# Main Form Fill
# =====================================================

def fill_main_form(driver, row, procedure_code):
    procedure_name = clean_value(row.get("Procedure Name", ""))
    package_rate = clean_amount(row.get("Package Rate", ""))

    defined_as = get_excel_or_default(row, "Defined As", "defined_as")
    department = get_excel_or_default(row, "Department", "department")
    rate_list = get_excel_or_default(row, "Rate List", "rate_list")
    procedure_type = get_excel_or_default(row, "Procedure Type", "procedure_type")

    description = procedure_name

    if procedure_code == "":
        raise Exception("Procedure Code is empty in Excel row.")

    if procedure_name == "":
        raise Exception("Procedure Name is empty in Excel row.")

    print("Filling main form...")

    print(" - Procedure Code")
    force_set_value(driver, config.PROCEDURE_CODE_INPUT, procedure_code)

    print(" - Procedure Name")
    clear_and_type(driver, config.PROCEDURE_NAME_INPUT, procedure_name)

    time.sleep(1)

    select_smart_dropdown(driver, config.DEFINED_AS_DROPDOWN, defined_as, "Defined As")
    select_smart_dropdown(driver, config.DEPARTMENT_DROPDOWN, department, "Department")
    time.sleep(2)
    select_smart_dropdown(driver, config.RATE_LIST_DROPDOWN, rate_list, "Rate List")
    select_smart_dropdown(driver, config.PROCEDURE_TYPE_DROPDOWN, procedure_type, "Procedure Type")

    print(" - Length of Stay")
    clear_and_type(driver, config.LENGTH_OF_STAY_INPUT, 1)

    print(" - Nights skipped")

    print(" - Description same as Procedure Name")
    clear_and_type(driver, config.DESCRIPTION_INPUT, description)

    attach_procedure_report(driver)

    open_components_section(driver)

    if package_rate is not None:
        print(" - Price")
        clear_and_type(driver, config.PRICE_INPUT, int(package_rate))
    else:
        print(" - Price skipped because Package Rate is empty")

    print(" - Notes skipped")

    print(" - Active skipped because field is not visible or already default")


# =====================================================
# Components Fill Based on Excel Amount + Billing Rule
# =====================================================

FIXED_BILLING_HEADS = [
    "procedure fee",
    "assistant surgeon fee",
    "anesthetist fee",
    "assistant anesthetist fee",
]


def normalize_text_for_match(text):
    return clean_value(text).lower().replace("  ", " ").strip()


def get_billing_type_for_head(head_name):
    head_clean = normalize_text_for_match(head_name)

    if head_clean in FIXED_BILLING_HEADS:
        return "Fixed"

    return "Estimated"


def get_component_table_rows(driver):
    print("Finding component table rows...")

    row_xpaths = [
        "//tr[.//td and (.//*[contains(normalize-space(), 'Fee')] or .//*[contains(normalize-space(), 'Charges')] or .//*[contains(normalize-space(), 'Pharmacy')] or .//*[contains(normalize-space(), 'Laboratory')] or .//*[contains(normalize-space(), 'Radiology')] or .//*[contains(normalize-space(), 'Cardiology')] or .//*[contains(normalize-space(), 'Visits')])]",
        "//table//tr[td]"
    ]

    for xpath in row_xpaths:
        try:
            rows = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.XPATH, xpath))
            )

            valid_rows = []

            for row in rows:
                row_text = clean_value(row.text)

                if row_text == "":
                    continue

                if "Heads" in row_text and "Billing Type" in row_text:
                    continue

                if "Procedure Name" in row_text or "Package Rate" in row_text:
                    continue

                valid_rows.append(row)

            if valid_rows:
                print(f"Component rows found: {len(valid_rows)}")
                return valid_rows

        except Exception:
            continue

    raise Exception("Component table rows not found.")


def get_head_name_from_row(row):
    try:
        cells = row.find_elements(By.XPATH, ".//td")

        if cells:
            head_name = clean_value(cells[0].text)

            if head_name:
                return head_name

        row_text = clean_value(row.text)

        if row_text:
            return row_text.split("\n")[0].strip()

    except Exception:
        pass

    return ""


def find_component_row_by_head(driver, head_name):
    row_xpaths = [
        f"//tr[.//td//*[normalize-space()='{head_name}']]",
        f"//tr[.//td[contains(normalize-space(), '{head_name}')]]",
        f"//*[normalize-space()='{head_name}']/ancestor::tr[1]",
        f"//*[contains(normalize-space(), '{head_name}')]/ancestor::tr[1]",
    ]

    for xpath in row_xpaths:
        try:
            row = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                row
            )

            time.sleep(0.3)
            return row

        except Exception:
            continue

    return None


def get_excel_amount_by_head(row_data, head_name):
    head_clean = normalize_text_for_match(head_name)

    for excel_column, value in row_data.items():
        if excel_column.startswith("_"):
            continue

        excel_clean = normalize_text_for_match(excel_column)

        if excel_clean == head_clean:
            return clean_amount(value), excel_column

    common_map = {
        "procedure fee": "Procedure Fee",
        "assistant surgeon fee": "Assistant Surgeon Fee",
        "anesthetist fee": "Anesthetist Fee",
        "assistant anesthetist fee": "Assistant Anesthetist Fee",
        "procedure room time charges": "Procedure Room Charges",
        "procedure room charges": "Procedure Room Charges",
        "stay charges": "Stay Charges",
        "procedure room pharmacy": "Procedure Room Pharmacy",
        "ward pharmacy": "Ward Pharmacy",
        "laboratory": "Laboratory",
        "radiology": "Radiology",
        "cardiology": "Cardiology",
        "visits": "Visits",
        "mo/nursing and allied charges": "MO/Nursing and Allied Charges",
    }

    mapped_column = common_map.get(head_clean)

    if mapped_column:
        return clean_amount(row_data.get(mapped_column, "")), mapped_column

    for excel_column, value in row_data.items():
        if excel_column.startswith("_"):
            continue

        excel_clean = normalize_text_for_match(excel_column)

        if excel_clean in head_clean or head_clean in excel_clean:
            return clean_amount(value), excel_column

    return None, ""

def select_billing_type_in_row(driver, head_name, billing_type):
    """
    Sirf us Head row ka Billing Type select karega jiska Excel mein amount available hai.

    Rule:
    - Procedure Fee / Assistant Surgeon Fee / Anesthetist Fee / Assistant Anesthetist Fee = Fixed
    - Baaki amount wale heads = Estimated

    Important:
    Ye function page ke kisi random Fixed/Estimated text ko click nahi karega.
    Sirf opened md-select dropdown ke visible md-option ko click karega.
    """

    print(f"   - Selecting Billing Type {billing_type} for: {head_name}")

    component_row = find_component_row_by_head(driver, head_name)

    if component_row is None:
        print(f"   - Row not found for billing type: {head_name}")
        return False

    try:
        md_selects = component_row.find_elements(By.TAG_NAME, "md-select")

        if not md_selects:
            print(f"   - Billing Type dropdown not found for: {head_name}")
            return False

        billing_dropdown = md_selects[0]

        current_text = clean_value(billing_dropdown.text)

        if billing_type.lower() in current_text.lower():
            print(f"   - {billing_type} already selected for: {head_name}")
            return True

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            billing_dropdown
        )
        time.sleep(0.4)

        # Open the dropdown
        try:
            billing_dropdown.click()
        except Exception:
            driver.execute_script("arguments[0].click();", billing_dropdown)

        time.sleep(0.8)

        # =====================================================
        # Method 1: Click only VISIBLE opened md-option
        # =====================================================
        try:
            clicked = driver.execute_script(
                """
                var target = arguments[0].trim().toLowerCase();

                var options = Array.from(document.querySelectorAll('md-option'));

                for (var i = 0; i < options.length; i++) {
                    var opt = options[i];
                    var rect = opt.getBoundingClientRect();
                    var style = window.getComputedStyle(opt);
                    var txt = (opt.innerText || opt.textContent || '').trim().toLowerCase();

                    var isVisible =
                        rect.width > 0 &&
                        rect.height > 0 &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        txt.length > 0;

                    if (isVisible && (txt === target || txt.indexOf(target) !== -1)) {
                        opt.click();
                        return true;
                    }
                }

                return false;
                """,
                billing_type
            )

            if clicked:
                time.sleep(0.8)

                # Verify after click
                component_row_after = find_component_row_by_head(driver, head_name)
                if component_row_after is not None:
                    md_after = component_row_after.find_elements(By.TAG_NAME, "md-select")
                    if md_after:
                        selected_text = clean_value(md_after[0].text)
                        if billing_type.lower() in selected_text.lower():
                            print(f"   - {billing_type} selected successfully for: {head_name}")
                            return True

                print(f"   - {billing_type} clicked for: {head_name}")
                return True

        except Exception as e:
            print(f"   - Visible option JS failed for {head_name}: {str(e)}")

        # =====================================================
        # Method 2: XPath only inside active md-select menu container
        # =====================================================
        option_xpaths = [
            f"//div[contains(@class,'md-select-menu-container') and not(contains(@class,'ng-hide'))]//md-option[contains(normalize-space(), '{billing_type}')]",
            f"//md-select-menu//md-option[contains(normalize-space(), '{billing_type}')]",
            f"//md-option[contains(normalize-space(), '{billing_type}')]"
        ]

        for xpath in option_xpaths:
            try:
                option = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )

                try:
                    option.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", option)

                time.sleep(0.8)

                print(f"   - {billing_type} selected by XPath for: {head_name}")
                return True

            except Exception:
                continue

        # =====================================================
        # Method 3: Keyboard fallback
        # Type first letter then Enter
        # Fixed = F, Estimated = E
        # =====================================================
        try:
            active = driver.switch_to.active_element
            active.send_keys(billing_type[0])
            time.sleep(0.3)
            active.send_keys(Keys.ENTER)
            time.sleep(0.8)

            print(f"   - {billing_type} selected by keyboard for: {head_name}")
            return True

        except Exception:
            pass

        # Close dropdown if still open
        try:
            driver.switch_to.active_element.send_keys(Keys.ESCAPE)
            time.sleep(0.3)
        except Exception:
            pass

        print(f"   - {billing_type} not selected for: {head_name}")
        return False

    except Exception as e:
        try:
            driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        except Exception:
            pass

        print(f"   - Billing Type error for {head_name}: {str(e)}")
        return False



def fill_amount_in_component_row(driver, head_name, amount):
    print(f"   - Filling Component Amount for {head_name}: {amount}")

    component_row = find_component_row_by_head(driver, head_name)

    if component_row is None:
        print(f"   - Row not found for amount: {head_name}")
        return False

    try:
        amount_input = None

        exact_selectors = [
            "input[ng-model='component.Amount']",
            "input[name='componentAmount']",
            "input[name*='componentAmount']",
            "input[type='number']",
        ]

        for selector in exact_selectors:
            inputs = component_row.find_elements(By.CSS_SELECTOR, selector)

            for inp in inputs:
                try:
                    if inp.is_displayed():
                        amount_input = inp
                        break
                except Exception:
                    continue

            if amount_input is not None:
                break

        if amount_input is None:
            cells = component_row.find_elements(By.XPATH, ".//td")

            if len(cells) >= 3:
                amount_cell = cells[2]

                inputs = amount_cell.find_elements(
                    By.XPATH,
                    ".//input[not(@type) or @type='text' or @type='number']"
                )

                for inp in inputs:
                    try:
                        if inp.is_displayed():
                            amount_input = inp
                            break
                    except Exception:
                        continue

        if amount_input is None:
            print(f"   - Component Amount input not found for: {head_name}")
            return False

        amount_value = str(int(amount))

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            amount_input
        )

        time.sleep(0.3)

        try:
            driver.execute_script(
                """
                arguments[0].removeAttribute('disabled');
                arguments[0].removeAttribute('readonly');
                """,
                amount_input
            )
        except Exception:
            pass

        try:
            amount_input.click()
            time.sleep(0.2)
            amount_input.send_keys(Keys.CONTROL, "a")
            amount_input.send_keys(Keys.BACKSPACE)
            amount_input.send_keys(amount_value)

            driver.execute_script(
                """
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
                """,
                amount_input
            )

            time.sleep(0.5)

            print(f"   - Amount typed successfully for: {head_name}")
            return True

        except Exception:
            pass

        driver.execute_script(
            """
            var el = arguments[0];
            var val = arguments[1];

            el.removeAttribute('disabled');
            el.removeAttribute('readonly');

            el.value = val;

            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
            """,
            amount_input,
            amount_value
        )

        time.sleep(0.5)

        print(f"   - Amount set by JS for: {head_name}")
        return True

    except Exception as e:
        print(f"   - Amount fill error for {head_name}: {str(e)}")
        return False



def fill_components(driver, row_data):
    """
    Final required flow:

    1. Sirf Heads read karega.
    2. Excel mein us Head ka amount check karega.
    3. Agar amount nahi hai:
       - Billing Type touch nahi karega.
       - Amount field touch nahi karega.
       - Next row.
    4. Agar amount hai:
       - First 4 fee heads = Fixed
       - Baaki heads = Estimated
       - Component Amount fill karega
    """

    print("Filling components based on Excel amounts only...")

    component_rows = get_component_table_rows(driver)

    for component_row in component_rows:
        head_name = get_head_name_from_row(component_row)

        if head_name == "":
            continue

        amount, excel_column = get_excel_amount_by_head(row_data, head_name)

        if amount is None:
            print(f" - Skip Head: {head_name} | Excel amount empty/not found")
            continue

        billing_type = get_billing_type_for_head(head_name)

        print("")
        print(f" - Processing Head: {head_name}")
        print(f"   Excel Column: {excel_column}")
        print(f"   Billing Type: {billing_type}")
        print(f"   Amount: {amount}")

        # Step 1: select Billing Type only for amount heads
        billing_selected = select_billing_type_in_row(driver, head_name, billing_type)

        if not billing_selected:
            print(f"   - Warning: Billing Type not selected for {head_name}. Amount fill will still continue.")

        # Step 2: fill amount in same head row
        amount_filled = fill_amount_in_component_row(driver, head_name, amount)

        if not amount_filled:
            print(f"   - Warning: Amount not filled for {head_name}.")

        time.sleep(0.5)

    print("Components fill completed.")


# =====================================================
# Save
# =====================================================

def save_form(driver):
    print("Clicking Save button...")

    try:
        button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, config.SAVE_BUTTON))
        )

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            button
        )

        time.sleep(0.5)

        try:
            button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", button)

    except Exception as e:
        return False, f"Save button click failed: {str(e)}"

    try:
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, config.SUCCESS_MESSAGE))
        )
        return True, "Saved successfully"

    except Exception:
        try:
            error_element = driver.find_element(By.CSS_SELECTOR, config.ERROR_MESSAGE)
            return False, error_element.text
        except Exception:
            return False, "Save failed or success message not found"


# =====================================================
# Main
# =====================================================

def main():
    log_records = []

    print("Reading Excel file...")

    try:
        excel_rows = load_excel_rows(config.EXCEL_FILE_PATH)
    except Exception as excel_error:
        print("Excel file read error:")
        print(excel_error)
        return

    print(f"Total records found: {len(excel_rows)}")

    if len(excel_rows) == 0:
        print("No records found in Excel.")
        return

    driver = setup_browser()

    try:
        auto_login_enabled = getattr(config, "AUTO_LOGIN", True)
        auto_open_enabled = getattr(config, "AUTO_OPEN_CREATE_PROCEDURE", True)

        if auto_login_enabled:
            auto_login(driver)
        else:
            driver.get(config.LOGIN_URL)
            input("Manual login aur Create Procedure page open karne ke baad ENTER press karein...")

        if auto_open_enabled:
            open_create_procedure_page(driver)
        else:
            input("Create Procedure page open hone ke baad ENTER press karein...")

        print("")
        print("Checking current page...")
        print("Current browser URL:", driver.current_url)
        print("Current page title:", driver.title)

        wait_for_element(driver, config.PROCEDURE_NAME_INPUT)

        for index, row in enumerate(excel_rows):
            excel_row_number = row.get("_excel_row_number", index + 2)
            procedure_name = clean_value(row.get("Procedure Name", ""))
            procedure_code = get_procedure_code(row, index)

            print("")
            print("=====================================================")
            print(f"Processing Excel row {excel_row_number}: {procedure_name}")
            print(f"Procedure Code: {procedure_code}")
            print("=====================================================")

            try:
                wait_for_element(driver, config.PROCEDURE_NAME_INPUT)

                fill_main_form(driver, row, procedure_code)

                fill_components(driver, row)

                # Save enabled
                success, message = save_form(driver)

                if success:
                    print(f"SUCCESS: {procedure_name}")
                    screenshot_path = ""
                    status = "Success"
                else:
                    print(f"FAILED: {procedure_name} - {message}")
                    screenshot_path = save_failed_screenshot(driver, excel_row_number)
                    status = "Failed"

                log_records.append({
                    "Excel Row": excel_row_number,
                    "Procedure Code": procedure_code,
                    "Procedure Name": procedure_name,
                    "Status": status,
                    "Message": message,
                    "Date Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Screenshot": screenshot_path
                })

                write_log(log_records)

                time.sleep(3)

                if index < len(excel_rows) - 1:
                    print("Opening next Create Procedure form...")
                    open_create_procedure_page(driver)

            except Exception as row_error:
                error_message = str(row_error)

                print(f"ERROR in Excel row {excel_row_number}: {error_message}")

                screenshot_path = save_failed_screenshot(driver, excel_row_number)

                log_records.append({
                    "Excel Row": excel_row_number,
                    "Procedure Code": procedure_code,
                    "Procedure Name": procedure_name,
                    "Status": "Failed",
                    "Message": error_message,
                    "Date Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Screenshot": screenshot_path
                })

                write_log(log_records)

                if index < len(excel_rows) - 1:
                    print("Error ke baad next form open kar raha hoon...")
                    open_create_procedure_page(driver)

                continue

    finally:
        write_log(log_records)

        total = len(log_records)
        success_count = len([x for x in log_records if x["Status"] == "Success"])
        failed_count = len([x for x in log_records if x["Status"] == "Failed"])

        print("")
        print("Automation Finished")
        print("-------------------")
        print(f"Total Processed: {total}")
        print(f"Success: {success_count}")
        print(f"Failed: {failed_count}")
        print(f"Log File: {config.LOG_FILE_PATH}")


if __name__ == "__main__":
    main()