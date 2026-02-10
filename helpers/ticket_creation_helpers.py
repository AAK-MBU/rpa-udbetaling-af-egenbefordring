"""This module contains helpers to navigate and leverage OPUS"""

import time

import logging

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logger = logging.getLogger(__name__)


def wait_and_click(browser, by, value):
    """Wait for an element to be clickable, then click it."""

    WebDriverWait(browser, 50).until(EC.presence_of_element_located((by, value)))

    click_element_with_retries(browser, by, value)


def press_key(keyboard, key):
    """Press and release a key on the keyboard."""

    keyboard.press(key)

    keyboard.release(key)


def switch_to_frame(browser, frame):
    """Switch to the required frames to access the form."""

    WebDriverWait(browser, 30).until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame)))


def enter_text(browser, by, value, text):
    """Helper to enter text into a form element."""

    input_element = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((by, value))
    )

    input_element.send_keys(text)


def click_element_with_retries(browser, by, value, retries=4):
    """Click an element with retries and handle common exceptions."""

    for attempt in range(retries):
        try:
            element = WebDriverWait(browser, 2).until(
                EC.element_to_be_clickable((by, value))
            )

            element.click()

            return True

        except Exception as e:  # pylint: disable=broad-except
            logger.info(f"Attempt {attempt + 1} failed: {e}")

            time.sleep(1)

    return False
