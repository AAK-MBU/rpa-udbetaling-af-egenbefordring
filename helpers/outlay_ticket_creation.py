"""This module contains the logic for creating an outlay ticket in OPUS."""

import sys
import os
import time

import logging

from pynput.keyboard import Key, Controller

from selenium import webdriver

from selenium.common.exceptions import TimeoutException

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

from mbu_dev_shared_components.utils.fernet_encryptor import Encryptor

from mbu_rpa_core.exceptions import BusinessError

from helpers.ticket_creation_helpers import wait_and_click, enter_text, switch_to_frame

logger = logging.getLogger(__name__)


def initialize_browser(opus_username, opus_password, headless=False):
    """Initialize the Selenium Chrome WebDriver."""
    chrome_options = Options()

    # Chrome prefs
    prefs = {"safebrowsing.enabled": False}
    chrome_options.add_experimental_option("prefs", prefs)

    ### REMOVE ###
    # Debug: stay open even on exit
    # chrome_options.add_experimental_option("detach", not headless)
    ### REMOVE ###

    # Common browser flags
    chrome_options.add_argument("test-type")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    chrome_options.add_argument("--incognito")

    # Enable headless mode
    if headless:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--window-size=1920,1080")

    browser = webdriver.Chrome(options=chrome_options)

    login_to_opus(browser, opus_username, opus_password)

    return browser


def login_to_opus(browser, username, password):
    """Login to OPUS."""

    browser.get("https://portal.kmd.dk/irj/portal")

    wait_and_click(browser, By.ID, 'logonuidfield')

    enter_text(browser, By.ID, 'logonuidfield', {username})
    enter_text(browser, By.ID, 'logonpassfield', {password})

    wait_and_click(browser, By.ID, 'buttonLogon')


def handle_opus(item_data, path, browser, headless):
    """Handle the OPUS ticket creation process."""

    attachment_path = os.path.join(path, f'receipt_{item_data["uuid"]}.pdf')

    navigate_to_opus(browser)

    logger.info("Filling form ...")
    fill_form(browser, item_data)

    logger.info("Uploading attachment ...")
    upload_attachment(browser, attachment_path, headless=headless)

    logger.info("Filling out form and controlling ...")
    fill_out_form_and_control(browser=browser, item_data=item_data)

    return

    sys.exit()

    # logger.info("Pressing 'Opret' to create ticket ...")
    # create_ticket(browser=browser)

    # logger.info("Successfully created outlay ticket.")


def navigate_to_opus(browser):
    """Navigate to OPUS page and open required tabs."""
    browser.get("https://portal.kmd.dk/irj/portal")

    wait_and_click(browser, By.XPATH, "//div[text()='Min Økonomi']")

    wait_and_click(browser, By.XPATH, "//div[text()='Bilag og fakturaer']")

    wait_and_click(browser, By.XPATH, "/html/body/div[1]/table/tbody/tr[1]/td/div/div[1]/div[9]/div[2]/span[2]")


def fill_form(browser, item_data):
    """
    Fill out the OPUS ticket form using Selenium.

    Structure of the function:
        1. Navigate to correct frames
        2. Fill creditor (CPR)
        3. Handle "Kreditoren kunne ikke oprettes..." error
        4. Fill main form fields (kommentar, tekst, reference, beløb, agent)
        5. Insert child name in popup window
        6. Return to base frame
    """

    # ---------------------------------------------------------
    # 1. NAVIGATE TO THE CORRECT FRAMES
    # ---------------------------------------------------------
    browser.switch_to.default_content()
    switch_to_frame(browser, "contentAreaFrame")
    switch_to_frame(browser, "ivuFrm_page0ivu0")

    # Shared root xpath for many fields
    root = (
        "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/"
        "tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/"
        "table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[1]/td/"
        "div/div/table/tbody/tr/td[1]/div/div/table/tbody/tr/td/div/"
        "div/table/tbody/"
    )

    # ---------------------------------------------------------
    # 2. FILL CREDITOR CPR
    # ---------------------------------------------------------
    creditor_input = (
        root + "tr[2]/td/div/div/table/tbody/tr/td[1]/div/div/table/"
        "tbody/tr[1]/td[2]/div/div/table/tbody/tr/td[1]/span/input"
    )
    enter_text(browser, By.XPATH, creditor_input, decrypt_cpr(item_data))

    # Click “Hent”
    hent_btn = (
        root + "tr[2]/td/div/div/table/tbody/tr/td[1]/div/div/table/"
        "tbody/tr[1]/td[2]/div/div/table/tbody/tr/td[2]/div"
    )
    wait_and_click(browser, By.XPATH, hent_btn)

    # ---------------------------------------------------------
    # 3. CHECK FOR INVALID CPR ERROR
    # ---------------------------------------------------------
    errorbox = browser.find_elements(By.ID, "WD0324")

    if errorbox and errorbox[0].text == (
        "Kreditoren kunne ikke oprettes automatisk. "
        "Det ikke er et SE/CVR eller CPR nummer."
    ):
        raise BusinessError("Kreditoren ikke oprettet.")

    time.sleep(3)  # OPUS needs time after “Hent”

    # ---------------------------------------------------------
    # 4. FILL MAIN FORM FIELDS
    # ---------------------------------------------------------

    # --- Kommentar ---
    kommentar_xpath = (
        "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/"
        "tbody/tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/"
        "tr[2]/td/table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/"
        "tr[1]/td/div/div/table/tbody/tr/td[2]/table/tbody/tr/td/div/"
        "table/tbody/tr[1]/td/div/div/div/div/table/tbody/tr[2]/td/"
        "div/textarea"
    )
    kommentar_value = item_data.get("evt_kommentar") or ""
    enter_text(browser, By.XPATH, kommentar_xpath, kommentar_value)

    # --- Udbetalingstekst ---
    enter_text(
        browser,
        By.XPATH,
        root + "tr[3]/td/div/div/table/tbody/tr[1]/td[1]/div/div/table/"
               "tbody/tr/td/div/div/table/tbody/tr[1]/td[2]/span/input",
        item_data["posteringstekst"],
    )

    # --- Posteringstekst (duplicate field) ---
    enter_text(
        browser,
        By.XPATH,
        root + "tr[3]/td/div/div/table/tbody/tr[2]/td/div/div/table/"
               "tbody/tr[2]/td[2]/span/input",
        item_data["posteringstekst"],
    )

    # --- Reference ---
    enter_text(
        browser,
        By.XPATH,
        root + "tr[3]/td/div/div/table/tbody/tr[2]/td/div/div/table/"
               "tbody/tr[3]/td[2]/span/input",
        item_data["reference"],
    )

    # --- Beløb ---
    enter_text(
        browser,
        By.XPATH,
        root + "tr[3]/td/div/div/table/tbody/tr[2]/td/div/div/table/"
               "tbody/tr[4]/td[2]/div/div/table/tbody/tr/td[1]/span/input",
        item_data["beloeb"],
    )

    # --- Næste agent ---
    enter_text(
        browser,
        By.XPATH,
        root + "tr[4]/td/div/div/table/tbody/tr[2]/td[2]/div/div/table/"
               "tbody/tr[1]/td[1]/span/input",
        item_data["naeste_agent"],
    )

    # ---------------------------------------------------------
    # 5. OPEN POPUP AND INSERT CHILD NAME
    # ---------------------------------------------------------
    # Click the button next to "udbetalingstekst" that opens popup
    popup_btn_xpath = (
        root + "tr[3]/td/div/div/table/tbody/tr[1]/td[1]/div/div/table/"
               "tbody/tr/td/div/div/table/tbody/tr[1]/td[3]/div"
    )
    wait_and_click(browser, By.XPATH, popup_btn_xpath)

    time.sleep(1)

    # Switch to popup's frame
    browser.switch_to.default_content()
    switch_to_frame(browser, "URLSPW-0")

    # Insert child name
    actions = ActionChains(browser)
    actions.send_keys(item_data["barnets_navn"])
    actions.perform()

    # Click “Gem”
    for button in browser.find_elements(By.CLASS_NAME, "lsButton"):
        if button.text.lower() == "gem":
            button.click()
            break

    # ---------------------------------------------------------
    # 6. RETURN BACK TO MAIN FRAME
    # ---------------------------------------------------------
    browser.switch_to.default_content()
    switch_to_frame(browser, "contentAreaFrame")
    switch_to_frame(browser, "ivuFrm_page0ivu0")


def decrypt_cpr(item_data):
    """Decrypt the CPR number from the element data."""

    encryptor = Encryptor()

    encrypted_cpr = item_data['cpr_encrypted']

    return encryptor.decrypt(encrypted_cpr.encode('utf-8'))


def upload_attachment(browser, attachment_path, headless=False):
    """
    Upload an attachment file into the OPUS form.

    Works in BOTH:
      - headless mode (send_keys to <input type="file">)
      - non-headless mode (Windows file picker via pynput)
    """

    # ---------------------------------------------------------
    # 1. CLICK "VEDHÆFT NYT"
    # ---------------------------------------------------------
    vedhaeft_nyt_xpath = (
        '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/'
        'tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/'
        'tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[1]/td/div/div/'
        'table/tbody/tr/td[2]/table/tbody/tr/td/div/table/tbody/tr[3]/td/div/'
        'span/span/div/span/span[1]/table/thead/tr[2]/th/div/div/div/span/div'
    )
    wait_and_click(browser, By.XPATH, vedhaeft_nyt_xpath)

    # ---------------------------------------------------------
    # 2. WAIT FOR POPUP TO LOAD
    # ---------------------------------------------------------
    WebDriverWait(browser, 20).until(
        lambda driver: driver.execute_script("return document.readyState") == "complete"
    )

    # ---------------------------------------------------------
    # 3. SWITCH TO POPUP FRAME
    # ---------------------------------------------------------
    browser.switch_to.default_content()
    switch_to_frame(browser, "URLSPW-0")

    if headless:
        # ---------------------------------------------------------
        # 4. LOCATE THE REAL <input type='file'>
        # ---------------------------------------------------------
        file_input_xpath = (
            "/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/"
            "tr/td/div/div/span/span[2]/form/input[@type='file']"
        )

        file_input = WebDriverWait(browser, 20).until(
            EC.presence_of_element_located((By.XPATH, file_input_xpath))
        )

        # ---------------------------------------------------------
        # 5. UPLOAD FILE (MODE-DEPENDENT)
        # ---------------------------------------------------------
        # ✅ HEADLESS: Selenium-only (NO OS dialog)
        file_input.send_keys(attachment_path)

    else:
        # ---------------------------------------------------------
        # 4. CLICK THE "VÆLG FIL" (CHOOSE FILE) BUTTON
        # ---------------------------------------------------------
        # OPUS uses a hidden <input type="file"> that can only be triggered by clicking.
        choose_file_xpath = (
            "/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/"
            "tr/td/div/div/span/span[2]/form"
        )
        wait_and_click(browser, By.XPATH, choose_file_xpath)

        # ---------------------------------------------------------
        # 5. TYPE THE FILE PATH INTO THE SYSTEM FILE PICKER
        # ---------------------------------------------------------
        # Selenium CANNOT interact with Windows' native file dialog.
        # So we use pynput's keyboard to type the path into the OS dialog.
        time.sleep(4)  # give OS time to open dialog window

        keyboard = Controller()
        keyboard.type(attachment_path)  # type full file path (C:\path\to\file.pdf)

        time.sleep(2)

        # ---------------------------------------------------------
        # 6. PRESS ENTER TO CONFIRM FILE SELECTION
        # ---------------------------------------------------------
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)

    time.sleep(2)

    # ---------------------------------------------------------
    # 7. CLICK "OK" INSIDE THE POPUP TO CONFIRM UPLOAD
    # ---------------------------------------------------------
    ok_button_xpath = (
        "/html/body/table/tbody/tr/td/div/div[1]/div/div[4]/div/table/"
        "tbody/tr/td[3]/table/tbody/tr/td[1]/div"
    )

    wait_and_click(browser, By.XPATH, ok_button_xpath)

    # Optional small delay for OPUS to process the uploaded file
    time.sleep(2)


def fill_out_form_and_control(browser, item_data):
    """
    Complete the OPUS form and validate the expense ticket.
    SAP-grid-safe, headless-safe, debug-friendly.
    """

    # ---------------------------------------------------------
    # 1. NAVIGATE INTO THE OPUS FORM FRAMES
    # ---------------------------------------------------------
    browser.switch_to.default_content()
    switch_to_frame(browser, "contentAreaFrame")
    switch_to_frame(browser, "ivuFrm_page0ivu0")

    # ---------------------------------------------------------
    # 2. FOCUS FIRST ROW IN POSTING TABLE
    # ---------------------------------------------------------
    first_row_arts_konto_cell = (
        "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/"
        "tbody/tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/"
        "table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[2]/td/div/"
        "span/span[1]/div/span/span[1]/div/div/div/span/span/table/tbody/"
        "tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/table/"
        "tbody/tr[2]/td[3]/table/tbody/tr/td/span"
    )

    WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.XPATH, first_row_arts_konto_cell))
    ).click()

    active = browser.switch_to.active_element
    active.send_keys(item_data["arts_konto"])

    active.send_keys(Keys.TAB)

    active = browser.switch_to.active_element
    active.send_keys(item_data["beloeb"])

    active.send_keys(Keys.TAB)
    active = browser.switch_to.active_element

    active.send_keys(Keys.TAB)
    active = browser.switch_to.active_element

    active.send_keys(Keys.TAB)
    active = browser.switch_to.active_element

    active.send_keys(item_data["psp"])

    active.send_keys(Keys.TAB)
    active = browser.switch_to.active_element

    active.send_keys(item_data["posteringstekst"])

    active.send_keys(Keys.TAB)
    active.send_keys(Keys.TAB)

    # time.sleep(1)

    # ---------------------------------------------------------
    # 5. ACTIVATE 'KONTROLLER'
    # ---------------------------------------------------------
    kontroller_button_xpath = (
        "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/"
        "tbody/tr/td/div/table/tbody/tr[1]/td/div/div[2]/div/div/div/"
        "span[4]/div"
    )

    kontroller = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located((By.XPATH, kontroller_button_xpath))
    )

    browser.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        kontroller
    )

    browser.execute_script("arguments[0].click();", kontroller)

    # ---------------------------------------------------------
    # 6. WAIT FOR VALIDATION RESULT
    # ---------------------------------------------------------
    # WebDriverWait(browser, 30).until(
    #     EC.presence_of_element_located(
    #         (By.XPATH, "//*[contains(text(), 'kontrolleret og OK')]")
    #     )
    # )

    try:
        WebDriverWait(browser, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(), 'kontrolleret og OK')]")
            )
        )

    except TimeoutException as e:
        raise BusinessError(
            "Fejl ved kontrol af udgiftsbilag - 'kontrolleret og OK' blev ikke fundet") from e

    print("\nbilag er kontrolleret ok\n")


# def fill_out_form_and_control(browser, item_data):
#     """
#     Complete the OPUS form and submit the expense ticket.

#     Steps performed:
#         1. Navigate into the correct OPUS frames
#         2. Select the first row of the posting table
#         3. Fill Artskonto, Beløb, PSP, Posteringstekst using keyboard tabbing
#         4. Click 'Kontroller' and ensure OPUS reports "OK"
#         5. Click 'Opret' to submit the ticket
#     """

#     # ---------------------------------------------------------
#     # 1. NAVIGATE INTO THE OPUS FORM FRAMES
#     # ---------------------------------------------------------
#     browser.switch_to.default_content()
#     switch_to_frame(browser, "contentAreaFrame")
#     switch_to_frame(browser, "ivuFrm_page0ivu0")

#     keyboard = Controller()  # used for typing and TAB navigation

#     # ---------------------------------------------------------
#     # 2. CLICK THE FIRST ROW IN THE POSTING TABLE
#     # ---------------------------------------------------------
#     # This focuses the "Artskonto" field in the first line of the table.
#     # (The long XPath points to row 1 -> column for Artskonto)
#     first_row_arts_konto_cell = (
#         "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/"
#         "tbody/tr/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/"
#         "table/tbody/tr/td/div/div[1]/div/div/div/table/tbody/tr[2]/td/div/"
#         "span/span[1]/div/span/span[1]/div/div/div/span/span/table/tbody/"
#         "tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/table/"
#         "tbody/tr[2]/td[3]/table/tbody/tr/td/span"
#     )
#     wait_and_click(browser, By.XPATH, first_row_arts_konto_cell)

#     # ---------------------------------------------------------
#     # 3. FILL FIELDS IN THE POSTINGS TABLE
#     # Using TAB to move between columns
#     # ---------------------------------------------------------

#     # Artskonto
#     keyboard.type(item_data["arts_konto"])

#     # Move to 'Beløb' column
#     press_key(keyboard, Key.tab)
#     keyboard.type(item_data["beloeb"])

#     # Move to PSP column (requires 3 TABs)
#     press_key(keyboard, Key.tab)
#     press_key(keyboard, Key.tab)
#     press_key(keyboard, Key.tab)

#     # PSP value
#     keyboard.type(item_data["psp"])

#     # Move to 'Posteringstekst' column
#     press_key(keyboard, Key.tab)
#     keyboard.type(item_data["posteringstekst"])

#     time.sleep(1)  # small delay so OPUS can update the row

#     # ---------------------------------------------------------
#     # 4. CLICK 'KONTROLLER' TO VALIDATE THE TICKET
#     # ---------------------------------------------------------
#     kontroller_button_xpath = (
#         "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/"
#         "tbody/tr/td/div/table/tbody/tr[1]/td/div/div[2]/div/div/div/"
#         "span[4]/div"
#     )
#     wait_and_click(browser, By.XPATH, kontroller_button_xpath)

#     time.sleep(4)  # OPUS needs time to run validation

#     # Validation check
#     kontrol_ok = browser.find_elements(
#         By.XPATH,
#         "//*[contains(text(), 'Udgiftsbilag er kontrolleret og OK')]"
#     )

#     if not kontrol_ok:
#         raise BusinessError("Fejl ved kontrol af udgiftsbilag.")

#     print("\nbilag er kontrolleret ok\n")


def create_ticket(browser):
    """
    Helper to press the 'Opret' button and create the ticket
    """

    # ---------------------------------------------------------
    # CLICK 'OPRET' TO SUBMIT THE TICKET
    # ---------------------------------------------------------
    opret_button_xpath = (
        "/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/"
        "tbody/tr/td/div/table/tbody/tr[1]/td/div/div[2]/div/div/div/"
        "span[1]/div"
    )
    wait_and_click(browser, By.XPATH, opret_button_xpath)

    time.sleep(4)

    # Confirm ticket was created
    oprettet_ok = browser.find_elements(
        By.XPATH,
        "//*[contains(text(), 'er oprettet')]"
    )

    if not oprettet_ok:
        time.sleep(1)
        raise BusinessError("Fejl ved oprettelse af udgiftsbilag, kontrol OK.")
