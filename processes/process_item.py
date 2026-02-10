"""Module to handle item processing"""
# from mbu_rpa_core.exceptions import ProcessError, BusinessError

import os
import logging

from helpers import outlay_ticket_creation, helper_functions

logger = logging.getLogger(__name__)

DBCONNECTIONSTRING = os.getenv("DBCONNECTIONSTRINGPROD")


def process_item(item_data: dict, item_reference: str, browser, headless, os2_api_key):
    """Function to handle item processing"""

    assert item_data, "Item data is required"
    assert item_reference, "Item reference is required"

    receipts = []

    folder_path, file_content = helper_functions.fetch_receipt(item_data=item_data, os2_api_key=os2_api_key)

    receipts.append(file_content)

    outlay_ticket_creation.handle_opus(item_data=item_data, path=folder_path, browser=browser, headless=headless)

    helper_functions.remove_attachment_if_exists(folder_path=folder_path, item_data=item_data)

    helper_functions.handle_post_process(failed=False, item_data=item_data, item_reference=item_reference)
