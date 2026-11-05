"""This module contains helpers to navigate and leverage OPUS"""

import time

import logging

from selenium.common.exceptions import (
    NoSuchFrameException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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


def switch_to_frame_containing(browser, by, value, timeout=30, max_depth=6):
    """
    Switch into whichever (nested) frame contains the given element.

    Searches depth-first from the top-level document, retrying until the
    timeout, since SAP portal frames load lazily. Leaves the driver switched
    into the matching frame and logs the frame path it took, which is handy
    for learning the frame IDs of a new page.
    """

    def _search(depth, path):
        if browser.find_elements(by, value):
            return path

        if depth >= max_depth:
            return None

        frame_count = len(browser.find_elements(By.XPATH, "//iframe | //frame"))

        for i in range(frame_count):
            # Re-query each time: switching frames invalidates old references.
            frames = browser.find_elements(By.XPATH, "//iframe | //frame")

            if i >= len(frames):
                break

            frame = frames[i]

            try:
                frame_label = frame.get_attribute("id") or frame.get_attribute("name") or f"#{i}"
                browser.switch_to.frame(frame)

            except (StaleElementReferenceException, NoSuchFrameException):
                continue

            found = _search(depth + 1, path + [frame_label])

            if found is not None:
                return found

            browser.switch_to.parent_frame()

        return None

    end_time = time.time() + timeout

    while True:
        browser.switch_to.default_content()

        found_path = _search(0, [])

        if found_path is not None:
            logger.info("Found %s in frame path: %s", value, " > ".join(found_path) or "<top level>")
            return found_path

        if time.time() >= end_time:
            raise TimeoutException(f"No frame contains element {by}={value} after {timeout}s")

        time.sleep(1)


def enter_text(browser, by, value, text):
    """Helper to enter text into a form element."""

    input_element = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((by, value))
    )

    input_element.send_keys(text)


def replace_text(browser, by, value, text):
    """
    Replace the contents of an input. Uses select-all + type instead of
    .clear(), since SAP Web Dynpro does not always register a .clear().
    """

    input_element = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((by, value))
    )

    input_element.click()
    input_element.send_keys(Keys.CONTROL, "a")
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
