from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.devtools.v145.runtime import run_script
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# DRIVER SETUP
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium import webdriver

def setup_driver():

    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=options)

    wait = WebDriverWait(driver, 20)

    return driver, wait



# NAVIGATION FUNCTION

def navigate_to_page(driver, wait,financial_year,expenditure_wise, grant_no, scheme_code, city):
    driver.get("https://koshvani.up.nic.in/")

    # WAIT for financial year dropdown
    fy_dropdown = wait.until(
        EC.presence_of_element_located((By.ID, "ddlFinYear"))
    )

    Select(fy_dropdown).select_by_visible_text(financial_year)


    # Click Grant-wise Expenditure
    wait.until(EC.element_to_be_clickable(
        (By.PARTIAL_LINK_TEXT, expenditure_wise))
    ).click()

    # Click Grant Number
    wait.until(EC.element_to_be_clickable(
        (By.PARTIAL_LINK_TEXT, grant_no))
    ).click()

    # Click Scheme Code
    wait.until(EC.element_to_be_clickable(
        (By.PARTIAL_LINK_TEXT, scheme_code))
    ).click()

    # Wait for treasury list
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

    # # Click Treasury (City)
    treasury = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH,
             f"//a[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{city.upper()}')]")
        )

    )
    driver.execute_script("arguments[0].click();", treasury)

    # Wait for data table to load (many rows)
    wait.until(
        lambda d: len(d.find_elements(By.XPATH, "//table//tr")) > 5
    )

    print(f"Navigation completed for {city}")





# ==============================
# TABLE EXTRACTION FUNCTION
# ==============================
def extract_final_page_table(driver, wait):

    # Wait for Voucher header to appear
    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//th[contains(text(),'Voucher')]")
        )
    )

    # Select table that contains Voucher header
    main_table = driver.find_element(
        By.XPATH,
        "//table[.//th[contains(text(),'Voucher')]]//div[@class='myDiv']"
    )

    rows = main_table.find_elements(By.XPATH, ".//tr")

    data = []

    for row in rows:
        cols = row.find_elements(By.XPATH, ".//th | .//td")
        row_data = [c.text.strip() for c in cols]
        if row_data:
            data.append(row_data)

    print(f"Extracted {len(data)} rows from final page")

    return data


# ==============================
# SAVE TO EXCEL FUNCTION
# ==============================
def save_to_excel(data, city):

    df = pd.DataFrame(data)

    # First row is header
    header = df.iloc[0]

    # Remaining rows
    df = df.iloc[1:].copy()

    # Remove unwanted rows dynamically
    df = df[
        (df[0].astype(str).str.strip() != "") &   # remove blank
        (~df[0].astype(str).str.startswith("("))  # remove numbering row
    ]

    # Assign header back
    df.columns = header

    df = df.reset_index(drop=True)

    df.to_excel(f"{city}.xlsx", index=False)

    print(f"{city}.xlsx created successfully")


# ==============================
# MAIN EXECUTION
# ==============================
if __name__ == "__main__":

    driver, wait = setup_driver()

    city = "Ayodhya"
    expenditure_wise="Grant-wise"
    grant_no = "001"
    scheme_code = "2039000010300"
    try:
        navigate_to_page(driver,wait,expenditure_wise,grant_no,scheme_code,city=city)

        data = extract_final_page_table(driver, wait)
        if data:
            save_to_excel(data, city)

    finally:
        driver.quit()

