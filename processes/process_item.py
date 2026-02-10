"""Module to handle item processing"""
# from mbu_rpa_core.exceptions import ProcessError, BusinessError

import sys
import os
import logging

from mbu_dev_shared_components.database import RPAConnection

from helpers import outlay_ticket_creation, helper_functions

logger = logging.getLogger(__name__)

DBCONNECTIONSTRING = os.getenv("DBCONNECTIONSTRINGPROD")

RPA_CONN = RPAConnection(db_env="PROD", commit=False)
with RPA_CONN:
    OPUS_CREDS = RPA_CONN.get_credential("egenbefordring_udbetaling")
    OPUS_USERNAME = OPUS_CREDS.get("username")
    OPUS_PASSWORD = OPUS_CREDS.get("decrypted_password", "")

    OS2_API_KEY = RPA_CONN.get_credential("os2_api").get("decrypted_password")

print(f"opus_username: {OPUS_USERNAME}")
print(f"opus_password: {OPUS_PASSWORD}")

print(f"OS2_API_KEY: {OS2_API_KEY}")


def process_item(item_data: dict, item_reference: str, browser, headless):
    """Function to handle item processing"""

    assert item_data, "Item data is required"
    assert item_reference, "Item reference is required"

    receipts = []

    folder_path, file_content = helper_functions.fetch_receipt(item_data=item_data, os2_api_key=OS2_API_KEY)

    receipts.append(file_content)

    outlay_ticket_creation.handle_opus(item_data=item_data, path=folder_path, browser=browser, headless=headless)

    helper_functions.remove_attachment_if_exists(folder_path=folder_path, item_data=item_data)

    helper_functions.handle_post_process(failed=False, item_data=item_data, item_reference=item_reference)
