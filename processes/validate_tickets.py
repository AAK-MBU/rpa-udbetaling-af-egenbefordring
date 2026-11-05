"""
Module for the --finalize stage: validate the number of tickets created in OPUS.

Not to be confused with finalize_process.py, which writes the per-file Excel
result back to SharePoint while items are being processed.
"""

import logging
from datetime import date, datetime

from mbu_dev_shared_components.database import RPAConnection

from helpers import ats_functions, outlay_ticket_creation

logger = logging.getLogger(__name__)


def parse_week_arg(argv: list[str]) -> int:
    """Return the week number given after --finalize, e.g. '--finalize 43'."""

    idx = argv.index("--finalize")

    try:
        week = int(argv[idx + 1])

    except (IndexError, ValueError) as e:
        raise ValueError("--finalize requires a week number, e.g. '--finalize 43'") from e

    if not 1 <= week <= 53:
        raise ValueError(f"Invalid week number: {week}")

    return week


def resolve_year(week: int, today: date) -> int:
    """
    The ISO year the week belongs to: the current year, unless the week is
    later than the current week (e.g. finalizing week 52 in early January),
    in which case it is last year's week.
    """

    current_year, current_week, _ = today.isocalendar()

    return current_year - 1 if week > current_week else current_year


def parse_reference(reference: str) -> tuple[int, int] | None:
    """
    Return (week, iso_year) from a reference like
    'Egenbefordring_43_19102026_25102026_<uuid>', or None if it doesn't match.
    """

    parts = reference.split("_")

    if len(parts) < 3 or parts[0].lower() != "egenbefordring":
        return None

    try:
        week = int(parts[1])
        start_date = datetime.strptime(parts[2], "%d%m%Y").date()

    except ValueError:
        return None

    return week, start_date.isocalendar().year


def count_completed_items_for_week(week: int, year: int) -> int:
    """Count completed ATS work items whose reference belongs to the given week and year."""

    # The ATS search narrows the fetch; the parse below does the exact match
    # (so week 4 doesn't also catch week 43, etc.).
    rows = ats_functions.fetch_run_workqueue_items(file_name=f"Egenbefordring_{week}_")

    count = 0

    for row in rows:
        parsed = parse_reference(str(row.get("reference") or ""))

        if parsed == (week, year) and row.get("status") == "completed":
            count += 1

    return count


def validate_created_tickets(week: int):
    """Compare completed work items for the week with bilag registered in OPUS today."""

    today = date.today()
    year = resolve_year(week, today)

    expected = count_completed_items_for_week(week, year)
    logger.info("ATS: %d completed work items for week %d/%d", expected, week, year)

    rpa_conn = RPAConnection(db_env="PROD", commit=False)
    with rpa_conn:
        opus_creds = rpa_conn.get_credential("egenbefordring_udbetaling")

    browser = outlay_ticket_creation.initialize_browser(
        opus_username=opus_creds.get("username"),
        opus_password=opus_creds.get("decrypted_password", ""),
        headless=False,
    )

    try:
        found = outlay_ticket_creation.count_tickets_registered(
            browser, registreringsdato=today.strftime("%d.%m.%Y")
        )

    finally:
        browser.quit()

    if found == expected:
        logger.info("MATCH: %d tickets in OPUS, %d completed work items", found, expected)

    else:
        logger.warning("MISMATCH: %d tickets in OPUS, %d completed work items", found, expected)

    return expected, found
