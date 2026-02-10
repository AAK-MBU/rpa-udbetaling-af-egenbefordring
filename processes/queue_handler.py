"""Module to hande queue population"""

import asyncio
import json
import logging

from automation_server_client import Workqueue

from mbu_msoffice_integration.sharepoint_class import Sharepoint

from mbu_dev_shared_components.database.connection import RPAConnection

from helpers import config, helper_functions

logger = logging.getLogger(__name__)


def retrieve_items_for_queue() -> list[dict]:
    """Function to populate queue"""
    data = []
    references = []

    helper_functions.delete_all_files_in_path(config.PATH)

    with RPAConnection(db_env="PROD", commit=False) as rpa_conn:
        egenbefordring_procargs = json.loads(rpa_conn.get_constant(constant_name="egenbefordring_procargs").get("value"))

        naeste_agent = egenbefordring_procargs.get("naeste_agent")

    sharepoint = Sharepoint(**config.SHAREPOINT_KWARGS)

    file_name = helper_functions.fetch_files(folder_name=config.FOLDER_NAME, sharepoint=sharepoint)

    data_df = helper_functions.load_excel_data(file_name=file_name, sharepoint=sharepoint)

    processed_df = helper_functions.process_data(data_df, naeste_agent, file_name)

    approved_df = processed_df[processed_df["is_godkendt"]]

    # Loop through each approved row and build queue data
    for _, row in approved_df.iterrows():
        row_data = {k: helper_functions.nan_to_none(v) for k, v in row.to_dict().items()}

        reference_file_name = str(file_name).replace(".xlsx", "")

        # Reference = posteringstekst + unique UUID
        reference = f"{reference_file_name}_{row_data.get('uuid')}"

        data.append(row_data)

        references.append(reference)

    items = [
        {"reference": ref, "data": d} for ref, d in zip(references, data, strict=True)
    ]

    return items


def create_sort_key(item: dict) -> str:
    """
    Create a sort key based on the entire JSON structure.
    Converts the item to a sorted JSON string for consistent ordering.
    """
    return json.dumps(item, sort_keys=True, ensure_ascii=False)


async def concurrent_add(workqueue: Workqueue, items: list[dict]) -> None:
    """
    Populate the workqueue with items to be processed.
    Uses concurrency and retries with exponential backoff.

    Args:
        workqueue (Workqueue): The workqueue to populate.
        items (list[dict]): List of items to add to the queue.

    Returns:
        None

    Raises:
        Exception: If adding an item fails after all retries.
    """
    sem = asyncio.Semaphore(config.MAX_CONCURRENCY)

    async def add_one(it: dict):
        reference = str(it.get("reference") or "")
        data = {"item": it}

        async with sem:
            for attempt in range(1, config.MAX_RETRIES + 1):
                try:
                    await asyncio.to_thread(workqueue.add_item, data, reference)
                    logger.info("Added item to queue with reference: %s", reference)
                    return True

                except Exception as e:
                    if attempt >= config.MAX_RETRIES:
                        logger.error(
                            "Failed to add item %s after %d attempts: %s",
                            reference,
                            attempt,
                            e,
                        )
                        return False

                    backoff = config.RETRY_BASE_DELAY * (2 ** (attempt - 1))

                    logger.warning(
                        "Error adding %s (attempt %d/%d). Retrying in %.2fs... %s",
                        reference,
                        attempt,
                        config.MAX_RETRIES,
                        backoff,
                        e,
                    )
                    await asyncio.sleep(backoff)

    if not items:
        logger.info("No new items to add.")
        return

    sorted_items = sorted(items, key=create_sort_key)
    logger.info(
        "Processing %d items sorted by complete JSON structure", len(sorted_items)
    )

    results = await asyncio.gather(*(add_one(i) for i in sorted_items))
    successes = sum(1 for r in results if r)
    failures = len(results) - successes

    logger.info(
        "Summary: %d succeeded, %d failed out of %d", successes, failures, len(results)
    )
