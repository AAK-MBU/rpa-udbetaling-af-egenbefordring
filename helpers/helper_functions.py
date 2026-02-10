"""Module with helper functions"""

import os
import logging

import json

import ast

import shutil

from datetime import datetime

from io import BytesIO

import requests

import pandas as pd

from mbu_dev_shared_components.os2forms import documents

from mbu_dev_shared_components.utils.fernet_encryptor import Encryptor
from mbu_dev_shared_components.database.connection import RPAConnection

from mbu_msoffice_integration.sharepoint_class import Sharepoint

from helpers import config, smtp_util, ats_functions
from processes import finalize_process

logger = logging.getLogger(__name__)


def delete_all_files_in_path(path):
    """Delete all files and directories in the given path."""

    # Check if the path exists and create it if it doesn't
    if not os.path.exists(path):
        logger.info(f"Directory does not exist. Creating: {path}")

        os.makedirs(path)

    for filename in os.listdir(path):
        file_path = os.path.join(path, filename)

        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)

                logger.info(f"File deleted: {file_path}")

            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)

                logger.info(f"Directory deleted: {file_path}")

        except (OSError, shutil.Error) as e:
            logger.info(f"Failed to delete {file_path}. Reason: {e}")


def fetch_files(folder_name, sharepoint):
    """Download Excel files from SharePoint to the specified path."""

    files = sharepoint.fetch_files_list(folder_name)
    logger.info(files)

    if not files:
        logger.info("No files found in the specified SharePoint folder.")

    if len(files) > 1:
        raise Exception(f"More than 1 file in folder - len of files: {len(files)}")

    file_name = files[0].get("Name")

    return file_name


def load_excel_data(file_name: str, sharepoint: Sharepoint) -> pd.DataFrame:
    """
    Load an Excel file from SharePoint into a pandas DataFrame,
    keeping only rows where 'godkendt' contains 'x' (case-insensitive).
    """

    logger.info("Loading Excel data...")

    bytes_data = sharepoint.fetch_file_using_open_binary(
        file_name=file_name,
        folder_name=config.FOLDER_NAME,
    )

    if not bytes_data:
        raise ValueError("No data returned from SharePoint")

    df = pd.read_excel(
        BytesIO(bytes_data),
        dtype={
            "cpr_barnet": str,
            "cpr_nr": str,
            "cpr_nr_paaanden": str,
        },
    )

    logger.info(f"Rows before filtering: {len(df)}")

    df = df[
        df["godkendt"]
        .astype(str)
        .str.lower()
        .str.contains("x", na=False)
    ]

    logger.info(f"Rows after filtering on godkendt='x': {len(df)}")

    return df


def process_data(df: pd.DataFrame, naeste_agent: str, file_name) -> pd.DataFrame:
    """Process the data and return a DataFrame with the required format."""

    encryptor = Encryptor()

    processed_data = []

    for _, row in df.iterrows():
        cpr_nr = (
            str(row["cpr_nr_paaanden"])
            if not pd.isnull(row["cpr_nr_paaanden"])
            else str(row["cpr_nr"])
        )

        attachments_str = str(row.get("attachments", ""))
        url = extract_url_from_attachments(attachments_str)

        skoleliste = (
            str(row["skoleliste"]).lower() if not pd.isnull(row["skoleliste"]) else ""
        )

        barnets_navn = str(row["barnets_navn"])

        month_year = extract_months_and_year(row["test"])
        month_year_child_name = f"{month_year}_{barnets_navn}"

        psp_value = determine_psp_value(skoleliste, row)

        encrypted_cpr = encryptor.encrypt(cpr_nr).decode("utf-8")

        # Ensure that the beloeb value is a string, replace all . with , and keep only the last comma
        beloeb_value = (
            row["aendret_beloeb_i_alt"]
            if not pd.isnull(row["aendret_beloeb_i_alt"])
            else row["beloeb_i_alt"]
        )

        if pd.notnull(beloeb_value):
            # Replace all periods with commas
            beloeb_value = str(beloeb_value).replace(".", ",")

            # If there are multiple commas, keep only the last one
            if beloeb_value.count(",") > 1:
                parts = beloeb_value.split(",")

                # Join all but the last part without commas, then add the last part with a comma
                beloeb_value = "".join(parts[:-1]) + "," + parts[-1]

        raw_excel_data = {k: nan_to_none(v) for k, v in row.to_dict().items()}

        new_row = {
            "file_name": file_name,
            "cpr_encrypted": encrypted_cpr,
            "barnets_navn": barnets_navn,
            "beloeb": beloeb_value,
            "reference": month_year_child_name,
            "arts_konto": "40430002",
            "psp": psp_value,
            "posteringstekst": f"Egenbefordring {month_year}",
            "naeste_agent": naeste_agent,
            "attachment": url,
            "uuid": row.get("uuid", pd.NA),
            "godkendt_af": row.get("godkendt_af", pd.NA),
            "skole": row["skriv_dit_barns_skole_eller_dagtilbud"]
            if not pd.isnull(row["skriv_dit_barns_skole_eller_dagtilbud"])
            else row["skoleliste"],
            "is_godkendt": "x" in str(row.get("godkendt", "")).lower(),
            "evt_kommentar": None if pd.isna(row.get("evt_kommentar")) else row.get("evt_kommentar"),
            "raw_excel_data": raw_excel_data
        }

        new_row = {k: nan_to_none(v) for k, v in new_row.items()}

        processed_data.append(new_row)

    df_processed = pd.DataFrame(processed_data)

    return df_processed


def nan_to_none(value):
    return None if pd.isna(value) else value


def extract_months_and_year(test_str):
    """Extract months and year from the test string."""

    month_map = {
        "January": "Januar",
        "February": "Februar",
        "March": "Marts",
        "April": "April",
        "May": "Maj",
        "June": "Juni",
        "July": "Juli",
        "August": "August",
        "September": "September",
        "October": "Oktober",
        "November": "November",
        "December": "December",
    }

    data = ast.literal_eval(test_str)

    months = set()

    year = None

    for entry in data:
        if isinstance(entry, dict) and "dato" in entry:
            date_str = entry["dato"]
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")

            month_name = date_obj.strftime("%B")
            months.add(month_map.get(month_name, month_name))

            year = date_obj.year

    sorted_months = sorted(months, key=lambda x: list(month_map.values()).index(x))

    month_str = "/".join(sorted_months)

    result = f"{month_str} {year}"

    return result


def extract_url_from_attachments(attachments_str: str) -> str:
    """Extract the URL from the attachments string."""

    if isinstance(attachments_str, str):
        start_index = attachments_str.find("https://")

        if start_index != -1:
            end_index = attachments_str.find("'", start_index)

            if end_index != -1 and end_index > start_index:
                return attachments_str[start_index:end_index]

    return pd.NA


def determine_psp_value(skoleliste: str, row: pd.Series) -> str:
    """Determine PSP value based on school list."""

    if (
        "langagerskolen" in skoleliste
        or "751090#1830" in skoleliste
        or "751090#2471" in skoleliste
    ):
        return "XG-5240220808-00004"

    if (
        "stensagerskolen" in skoleliste
        or "751903#591" in skoleliste
        or "751903#2521" in skoleliste
    ):
        return "XG-5240220808-00005"

    if not pd.isnull(row["skriv_dit_barns_skole_eller_dagtilbud"]):
        return "XG-5240220808-00006"

    # Default PSP value
    return "XG-5240220808-00003"


def get_status_params(journalizing_table: bool = True, form_id: str = ""):
    """
    Generates a set of status parameters for the process, based on the given form_id and JSON arguments.

    Args:
        form_id (str): The unique identifier for the current process.
        case_metadata (dict): A dictionary containing various process-related arguments, including table names.

    Returns:
        tuple: A tuple containing three dictionaries:
            - status_params_inprogress: Parameters indicating that the process is in progress.
            - status_params_success: Parameters indicating that the process completed successfully.
            - status_params_failed: Parameters indicating that the process has failed.
            - status_params_manual: Parameters indicating that the process is handled manually.
    """

    if journalizing_table:
        id_name = "form_id"

    else:
        id_name = "uuid"

    status_params_inprogress = {
        "Status": ("str", "InProgress"),
        f"{id_name}": ("str", f'{form_id}')
    }

    status_params_success = {
        "Status": ("str", "Successful"),
        f"{id_name}": ("str", f'{form_id}')
    }

    status_params_failed = {
        "Status": ("str", "Failed"),
        f"{id_name}": ("str", f'{form_id}')
    }

    status_params_manual = {
        "Status": ("str", "Manual"),
        f"{id_name}": ("str", f'{form_id}')
    }

    return status_params_inprogress, status_params_success, status_params_failed, status_params_manual


def fetch_receipt(item_data, os2_api_key):
    """Fetch a receipt from OS2FORMS and save it to the specified path."""

    file_name = item_data.get("file_name")

    file_name_without_ext = os.path.splitext(file_name)[0]

    attachment_url = item_data.get('attachment')

    form_uuid = item_data.get('uuid')

    if not attachment_url or not form_uuid:
        error_message = "Missing 'attachment' URL or 'uuid' in element data."
        raise ValueError(error_message)

    try:
        # Download the file bytes
        file_content = documents.download_file_bytes(attachment_url, os2_api_key)

        new_path = os.path.join(config.PATH, file_name_without_ext)
        if not os.path.exists(new_path):
            os.makedirs(new_path)

        file_name = f"receipt_{form_uuid}.pdf"
        file_path = os.path.join(new_path, file_name)

        # Save the file content
        with open(file_path, 'wb') as f:
            f.write(file_content)

        logger.info(f"File downloaded and saved successfully to {file_path}.")

    except requests.exceptions.RequestException as e:
        error_message = f"Network error downloading file from OS2FORMS: {e}"
        raise RuntimeError(error_message) from e

    except OSError as e:
        error_message = f"Error saving the file from OS2FORMS: {e}"
        raise RuntimeError(error_message) from e

    return new_path, file_content


def remove_attachment_if_exists(folder_path, item_data):
    """Remove the attachment file if it exists."""

    attachment_path = os.path.join(folder_path, f'receipt_{item_data["uuid"]}.pdf')

    if os.path.exists(attachment_path):
        logger.info(f"Removing attachment file: {attachment_path}")

        os.remove(attachment_path)


def handle_post_process(failed: bool, item_data: str, item_reference: str):
    """Update the Excel file with the status of the element."""

    ats_functions.update_work_item_data(item_reference, failed)

    logger.info(f"Behandlet status updated to {'failed' if failed else 'succeeded'}")

    current_run_file_name = item_data.get("file_name")

    finalize_process.finalize_process(current_run_file_name=current_run_file_name)


def ensure_columns(df: pd.DataFrame, column_order: list) -> pd.DataFrame:
    """Ensure that the Excel file has the necessary columns and clean nulls."""

    string_columns = [
        "behandlet_fejl",
        "behandlet_ok",
    ]

    # 1. Reindex first → ensures all columns exist
    df = df.reindex(columns=column_order)

    # 2. Replace NaN / None with empty cells (Excel-safe)
    df = df.where(pd.notna(df), "")

    for col in string_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    return df


def send_mail(failed_work_items: bool):
    """Function to send email to inputted receiver"""

    logger.info("Sending email")

    with RPAConnection(db_env="PROD", commit=False) as rpa_conn:
        # egenbefordring_procargs = json.loads(rpa_conn.get_constant(constant_name="egenbefordring_procargs").get("value"))
        # email_receivers = egenbefordring_procargs.get("notification_email")

        email_sender = rpa_conn.get_constant("e-mail_noreply").get("value")

        smtp_server = rpa_conn.get_constant("smtp_server").get("value")
        smtp_port = rpa_conn.get_constant("smtp_port").get("value")

    if failed_work_items:
        folder_dest = "Fejlet"

    else:
        folder_dest = "Behandlet"

    folder_url = "/".join(
        [
            config.SHAREPOINT_SITE_URL,
            "teams",
            config.SHAREPOINT_SITE_NAME,
            config.DOCUMENT_LIBRARY,
            config.FOLDER_NAME,
            folder_dest
        ]
    )

    email_subject = "Robotten til egenbefordring er kørt"

    email_body = ('<p>Robotten til egenbefordring er nu kørt '
                  'og oversigten samt eventuelt relevante dokumenter '
                  f'er uploadet til <a href="{folder_url}">{folder_dest}-mappen</a></p>')

    smtp_util.send_email(
        receiver="dadj@aarhus.dk",
        # receiver=email_receivers,
        sender=email_sender,
        subject=email_subject,
        body=email_body,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        html_body=True,
    )
