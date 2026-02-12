"""Helper module to call some functionality in Automation Server using the API"""

import logging
import os

from datetime import datetime

import requests

from dateutil.parser import isoparse

from automation_server_client import WorkItem, Workqueue
from dotenv import load_dotenv

load_dotenv()

URL = os.getenv("ATS_URL")
TOKEN = os.getenv("ATS_TOKEN")

HEADERS = {"Authorization": f"Bearer {TOKEN}"}


def get_workqueue_items(workqueue: Workqueue, return_data=False):
    """
    Retrieve items from the specified workqueue.
    If the queue is empty, return an empty list.
    """

    if not URL or not TOKEN:
        raise OSError("ATS_URL or ATS_TOKEN is not set in the environment")

    workqueue_items = {} if return_data else set()

    page = 1
    size = 200  # max allowed

    while True:
        full_url = f"{URL}/workqueues/{workqueue.id}/items?page={page}&size={size}"
        response = requests.get(full_url, headers=HEADERS, timeout=60)
        response.raise_for_status()

        res_json = response.json().get("items", [])

        if not res_json:
            break

        for row in res_json:
            ref = row.get("reference")
            if ref:
                workqueue_items.add(ref)

        page += 1

    return workqueue_items


def fetch_run_workqueue_items(file_name: str = ""):
    """
    ATS helper to fetch workqueue items for the current run
    """

    workqueue_name = "bur.befordring.udbetaling_af_egenbefordring"

    workqueue_url = f"{URL}/workqueues/by_name/{workqueue_name}"

    workqueue_respone_json = requests.get(url=workqueue_url, headers=HEADERS, timeout=10).json()

    workqueue_id = workqueue_respone_json.get("id")

    work_items_url = f"{URL}/workqueues/{workqueue_id}/items?page=1&size=200&search={file_name}"

    run_workqueue_items = requests.get(url=work_items_url, headers=HEADERS, timeout=10).json()["items"]

    return run_workqueue_items


def update_work_item_data(item_reference: str, failed: bool):
    """
    ATS helper to update work item data
    """

    url = f"{URL}/workitems/by-reference/{item_reference}"

    raw = requests.get(url=url, headers=HEADERS, timeout=20).json()[0]
    work_item = WorkItem(**raw)

    if failed:
        work_item.data["item"]["data"]["raw_excel_data"]["behandlet_fejl"] = "x"

    else:
        work_item.data["item"]["data"]["raw_excel_data"]["behandlet_ok"] = "x"

    work_item.update(work_item.data)


def get_failed_workqueue_items(workqueue: Workqueue, from_date: datetime, to_date: datetime):
    """
    Function to retrieve failed workqueue items for a given time period
    """

    load_dotenv()

    failed_items = []

    if not URL or not TOKEN:
        raise OSError("ATS_URL or ATS_TOKEN is not set in the environment")

    page = 1
    size = 200  # max allowed

    while True:
        full_url = f"{URL}/workqueues/{workqueue.id}/items?page={page}&size={size}"
        response = requests.get(full_url, headers=HEADERS, timeout=60)
        response.raise_for_status()

        res_items = response.json().get("items", [])

        if not res_items:
            break

        for row in res_items:
            item_created_at_str = row.get("created_at")

            if not item_created_at_str:
                continue  # or handle differently

            item_created_at = isoparse(item_created_at_str)

            if not from_date < item_created_at < to_date:
                continue

            status = row.get("status")

            if status == "failed":
                failed_items.append(row)

        page += 1

    return failed_items


def get_item_info(item: WorkItem):
    """Unpack item"""
    return item.data["item"]["data"], item.data["item"]["reference"]


def init_logger():
    """Initialize the root logger with JSON formatting."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(module)s.%(funcName)s:%(lineno)d — %(message)s",
        datefmt="%H:%M:%S",
    )
