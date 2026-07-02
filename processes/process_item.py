"""Module to handle item processing"""

import os
import time
import logging

from mbu_rpa_core.exceptions import BusinessError

from helpers import config, outlay_ticket_creation, helper_functions

logger = logging.getLogger(__name__)

DBCONNECTIONSTRING = os.getenv("DBCONNECTIONSTRINGPROD")


def process_item(item_data: dict, item_reference: str, browser, headless, os2_api_key):
    """Function to handle item processing"""

    assert item_data, "Item data is required"
    assert item_reference, "Item reference is required"

    # Add database update
    helper_functions.update_process_status(conn_string=DBCONNECTIONSTRING, form_id=item_data.get("uuid"), status="InProgess")

    folder_path = helper_functions.fetch_receipt(item_data=item_data, os2_api_key=os2_api_key)

    handle_opus_with_retry(item_data=item_data, folder_path=folder_path, browser=browser, headless=headless)

    helper_functions.remove_attachment_if_exists(folder_path=folder_path, item_data=item_data)

    helper_functions.handle_post_process(failed=False, item_data=item_data, item_reference=item_reference)

    helper_functions.update_process_status(conn_string=DBCONNECTIONSTRING, form_id=item_data.get("uuid"), status="Successful")


def handle_opus_with_retry(item_data, folder_path, browser, headless):
    """
    Run the OPUS flow for a single item, retrying from a clean
    'navigate to OPUS' state if a transient browser/GUI issue (buffering,
    interrupted click, etc.) occurs.

    BusinessErrors (e.g. duplicate ticket) are never retried - they're
    raised immediately so the item is routed to manual review instead of
    being attempted again.
    """

    for attempt in range(1, config.MAX_ITEM_ATTEMPTS + 1):
        try:
            outlay_ticket_creation.handle_opus(
                item_data=item_data, path=folder_path, browser=browser, headless=headless
            )

            return

        except BusinessError:
            helper_functions.update_process_status(conn_string=DBCONNECTIONSTRING, form_id=item_data.get("uuid"), status="Failed")

            raise

        except Exception as e:
            helper_functions.update_process_status(conn_string=DBCONNECTIONSTRING, form_id=item_data.get("uuid"), status="Failed")

            if attempt == config.MAX_ITEM_ATTEMPTS:
                raise

            logger.warning(
                "OPUS flow failed on attempt %d/%d, retrying from a clean state: %s",
                attempt,
                config.MAX_ITEM_ATTEMPTS,
                e,
            )

            time.sleep(2)
