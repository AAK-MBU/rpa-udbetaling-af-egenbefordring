"""
This is the main entry point for the process
"""

import asyncio
import logging
import sys

from automation_server_client import AutomationServer, Workqueue

from mbu_dev_shared_components.database import RPAConnection

from mbu_rpa_core.exceptions import BusinessError, ProcessError
from mbu_rpa_core.process_states import CompletedState

from helpers import ats_functions, config, outlay_ticket_creation

from processes.application_handler import close, reset, startup
from processes.error_handling import ErrorContext, handle_error
from processes.process_item import process_item
from processes.queue_handler import concurrent_add, retrieve_items_for_queue

logger = logging.getLogger(__name__)


# ╔══════════════════════════════════════════════╗
# ║ 🔥 REMOVE BEFORE DEPLOYMENT (TEMP OVERRIDES) 🔥 ║
# ╚══════════════════════════════════════════════╝
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
_old_request = requests.Session.request
def unsafe_request(self, *args, **kwargs):
    kwargs['verify'] = False
    return _old_request(self, *args, **kwargs)
requests.Session.request = unsafe_request
# ╔══════════════════════════════════════════════╗
# ║ 🔥 REMOVE BEFORE DEPLOYMENT (TEMP OVERRIDES) 🔥 ║
# ╚══════════════════════════════════════════════╝


async def populate_queue(workqueue: Workqueue):
    """Populate the workqueue with items to be processed."""

    logger.info("Populating workqueue...")

    items_to_queue = retrieve_items_for_queue()

    queue_references = {str(r) for r in ats_functions.get_workqueue_items(workqueue)}

    new_items: list[dict] = []
    for item in items_to_queue:
        reference = str(item.get("reference") or "")
        if reference and reference in queue_references:
            logger.info(
                "Reference: %s already in queue. Item: %s not added",
                reference,
                item,
            )
        else:
            new_items.append(item)

    await concurrent_add(workqueue, new_items)
    logger.info("Finished populating workqueue.")


async def process_workqueue(workqueue: Workqueue):
    """Process items from the workqueue."""

    logger.info("Processing workqueue...")

    rpa_conn = RPAConnection(db_env="PROD", commit=False)
    with rpa_conn:
        opus_creds = rpa_conn.get_credential("egenbefordring_udbetaling")
        opus_username = opus_creds.get("username")
        opus_password = opus_creds.get("decrypted_password", "")

    startup()

    error_count = 0

    while error_count < config.MAX_RETRY:
        headless = True

        browser = outlay_ticket_creation.initialize_browser(opus_username=opus_username, opus_password=opus_password, headless=headless)

        for item in workqueue:
            try:
                with item:
                    data, reference = ats_functions.get_item_info(item)

                    try:
                        logger.info("Processing item with reference: %s", reference)
                        process_item(data, reference, browser, headless)

                        completed_state = CompletedState.completed(
                            "Process completed without exceptions"
                        )
                        item.complete(str(completed_state))

                        continue

                    except BusinessError as e:
                        context = ErrorContext(
                            item=item,
                            action=item.pending_user(str(e)),
                            send_mail=True,
                            process_name=workqueue.name,
                        )
                        handle_error(
                            error=e,
                            log=logger.info,
                            context=context,
                            item=item
                        )

                    except Exception as e:
                        pe = ProcessError(str(e))
                        raise pe from e

            except ProcessError as e:
                context = ErrorContext(
                    item=item,
                    action=item.fail,
                    send_mail=True,
                    process_name=workqueue.name
                )
                handle_error(
                    error=e,
                    log=logger.error,
                    context=context,
                    item=item
                )
                error_count += 1
                reset()

        break

    logger.info("Finished processing workqueue.")
    close()

if __name__ == "__main__":
    ats_functions.init_logger()

    ats = AutomationServer.from_environment()

    prod_workqueue = ats.workqueue()
    process = ats.process

    ### REMOVE !!! ###
    prod_workqueue.clear_workqueue()
    ### REMOVE !!! ###

    if "--queue" in sys.argv:
        asyncio.run(populate_queue(prod_workqueue))

    if "--process" in sys.argv:
        asyncio.run(process_workqueue(prod_workqueue))

    sys.exit(0)
