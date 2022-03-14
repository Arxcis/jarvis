import os
import re
import time

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait       
from selenium.webdriver.support import expected_conditions as EC

from getpass import getpass
from openpyxl import load_workbook
from datetime import datetime

def main():
    #
    # Config
    #
    EMAIL = "jon.filip.ultvedt@neat.no"
    NETSUITE_URL = "https://5677765.app.netsuite.com/app/center/card.nl"
    INPUT = "./input.xlsx"
    OUTPUT_DIR = "./output"
    YEAR = 2022
    MONTH = 2

    #
    # Program
    #
    if os.path.exists(OUTPUT_DIR):
        os.rmdir(OUTPUT_DIR)
    os.mkdir(OUTPUT_DIR)

    invoices = read_input(INPUT, YEAR, MONTH)
    browser = make_browser(OUTPUT_DIR)
    authenticate(browser, NETSUITE_URL, EMAIL)

    for i, invoice in enumerate(invoices):
        download_invoice(browser, invoice, NETSUITE_URL, OUTPUT_DIR)
        print(f"{i}/{len(invoice)}: {invoice}")


def read_input(INPUT, YEAR, MONTH):
    sheet = load_workbook(filename=INPUT).active

    i = 1
    invoices = []
    while True:
        try:
            A,B,C,D = sheet[f"A{i}"].value, sheet[f"B{i}"].value, sheet[f"C{i}"].value, sheet[f"D{i}"].value
        except ValueError:
            break

        i += 1
        if A is None:
            continue
        
        match = re.search("^[A-Z]+\d+$", A) 
        if match is None:
            continue
        
        date = datetime.strptime(B, "%m/%d/%Y")
        if date.month != MONTH:
            continue

        if date.year != YEAR:
            continue
        
        invoices.append(A)
        continue

    print("")
    print(f"Found {len(invoices)} invoices from {YEAR}-{MONTH} in {INPUT}!")

    return invoices
    

def make_browser(OUTPUT_DIR):
    DOWNLOADS_DIR = f"{OUTPUT_DIR}/downloads"

    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
        "download.default_directory": DOWNLOADS_DIR, #Change default directory for downloads
        "download.prompt_for_download": False, #To auto download the file
        #"download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome 
    })
    browser = webdriver.Chrome(ChromeDriverManager().install(), options=options)

    return browser


def authenticate(browser, NETSUITE_URL, EMAIL):
    # 0. Get username and password
    print("")
    print(f"--- Sign into {NETSUITE_URL} ---")
    print("")
    print(f"Email: {EMAIL}")
    password = getpass("Password:")
    secret_answer = getpass("Secret answer:")
    print("")

    # 1. Open browser window
    browser.maximize_window()
    browser.get(NETSUITE_URL)

    # 2. Fill in email and password
    input_email = browser.find_element(by=By.CSS_SELECTOR, value="input#email")
    input_password = browser.find_element(by=By.CSS_SELECTOR, value="input#password")
    input_email.send_keys(EMAIL)
    input_password.send_keys(password + Keys.ENTER)

    # 3. Answer to jesus
    answer = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="answer"]')))
    answer.send_keys(secret_answer + Keys.ENTER)


def download_invoice(browser, invoice, NETSUITE_URL, OUTPUT_DIR):
    DOWNLOADS_DIR = f"{OUTPUT_DIR}/downloads"

    # 0. Refresh browser
    browser.get(NETSUITE_URL)

    # 1. Prepare output folder
    print(f"Downloading {invoice}...")
    if os.path.exists(DOWNLOADS_DIR):
        os.rmdir(DOWNLOADS_DIR)
    os.mkdir(DOWNLOADS_DIR)

    # 2. Search
    search = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#_searchstring")))
    search.send_keys(f"{invoice}{Keys.ENTER}")

    # 3. View Invoice
    view = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#row0 > td.listtextctr > a.dottedlink.viewitem")))
    view.click()

    # 4. Print invoice pdf
    browser.execute_script("NLInvokeButton(getButton('print'))")

    # 5. Rename file
    filename = None
    while True:
        try:
            filename = os.listdir(DOWNLOADS_DIR)[0]
        except BaseException:
            time.sleep(1)
            continue
        break
    os.rename(f"{DOWNLOADS_DIR}/{filename}", f"{OUTPUT_DIR}/{invoice}.pdf")



if __name__ == "__main__":
    main()
