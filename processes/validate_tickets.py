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


def parse_week_arg(argv: list[str]) -> int | None:
    """
    Return the week number given after --finalize, e.g. '--finalize 43'.

    Returns None when no week is given ('--finalize' last, or followed by
    another flag), meaning the week should be auto-detected.
    """

    idx = argv.index("--finalize")

    if idx + 1 >= len(argv) or argv[idx + 1].startswith("--"):
        return None

    try:
        week = int(argv[idx + 1])

    except ValueError as e:
        raise ValueError(f"Invalid week number after --finalize: {argv[idx + 1]!r}") from e

    if not 1 <= week <= 53:
        raise ValueError(f"Invalid week number: {week}")

    return week


def detect_week_from_new_items() -> tuple[int, int] | None:
    """
    Auto-detect the (week, year) to validate from the work items still 'new',
    i.e. the ones the upcoming --process run will handle. Must be called
    before --process, since afterwards nothing is 'new' anymore.

    Returns None (and logs why) if there are no new items, if they span more
    than one week, or if ATS can't be reached. Never raises, so a failed
    detection can't block processing.
    """

    try:
        rows = ats_functions.fetch_run_workqueue_items(file_name="Egenbefordring_")

    except Exception as e:  # pylint: disable=broad-except
        logger.warning("Finalize skipped: could not fetch work items to detect week: %s", e)
        return None

    weeks = set()

    for row in rows:
        if row.get("status") != "new":
            continue

        parsed = parse_reference(str(row.get("reference") or ""))

        if parsed is not None:
            weeks.add(parsed)

    if not weeks:
        logger.warning("Finalize skipped: no new work items to detect the week from.")
        return None

    if len(weeks) > 1:
        listed = ", ".join(f"{w}/{y}" for w, y in sorted(weeks, key=lambda t: (t[1], t[0])))
        logger.warning(
            "Finalize skipped: new work items span multiple weeks (%s). "
            "Run '--finalize <week>' for each week separately.",
            listed,
        )
        return None

    week, year = weeks.pop()
    logger.info("Finalize: auto-detected week %d/%d from new work items.", week, year)

    return week, year


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


def is_duplicate(row: dict) -> bool:
    """
    A work item stopped because OPUS reported the bilag as already registered.
    The BusinessError text from fill_out_form_and_control ('... (duplikat).')
    is stored as the item's message when it is set to pending user action.
    """

    return "duplikat" in str(row.get("message") or "").lower()


def summarize_items_for_week(week: int, year: int) -> dict:
    """
    Group the ATS work items for the given week and year by outcome:
      completed  - ticket created in OPUS
      duplicates - no ticket created, OPUS already had the bilag
      other      - failed / pending user action for any other reason, or not yet processed
    """

    # The ATS search narrows the fetch; parse_reference does the exact match
    # (so week 4 doesn't also catch week 43, etc.).
    rows = ats_functions.fetch_run_workqueue_items(file_name=f"Egenbefordring_{week}_")

    summary = {"total": 0, "completed": 0, "duplicates": [], "other": []}

    for row in rows:
        reference = str(row.get("reference") or "")

        if parse_reference(reference) != (week, year):
            continue

        summary["total"] += 1
        status = row.get("status")

        if status == "completed":
            summary["completed"] += 1

        elif is_duplicate(row):
            summary["duplicates"].append(reference)

        else:
            summary["other"].append(f"{reference} ({status})")

    return summary


def log_summary(week: int, year: int, summary: dict, found: int, registreringsdato: str) -> None:
    """Log the ATS breakdown and the OPUS comparison."""

    expected = summary["completed"]
    duplicates = summary["duplicates"]
    other = summary["other"]

    logger.info("===== Ticket validation for week %d/%d =====", week, year)
    logger.info("ATS work items for the week:            %d", summary["total"])
    logger.info("  completed (ticket created):           %d", expected)
    logger.info("  duplicates (no ticket created):       %d", len(duplicates))

    for ref in duplicates:
        logger.info("      - %s", ref)

    logger.info("  other failed / pending / not done:    %d", len(other))

    for ref in other:
        logger.info("      - %s", ref)

    logger.info("OPUS bilag registered %s:          %d", registreringsdato, found)

    duplicate_note = (
        f" {len(duplicates)} duplicate(s) were skipped and are not expected in OPUS."
        if duplicates else ""
    )

    if found == expected:
        logger.info("MATCH: %d tickets in OPUS = %d completed work items.%s", found, expected, duplicate_note)

    else:
        logger.warning(
            "MISMATCH: %d tickets in OPUS vs %d completed work items (difference %+d).%s",
            found, expected, found - expected, duplicate_note,
        )


def validate_created_tickets(week: int, year: int | None = None):
    """Compare completed work items for the week with bilag registered in OPUS today."""

    today = date.today()

    if year is None:
        year = resolve_year(week, today)
    registreringsdato = today.strftime("%d.%m.%Y")

    summary = summarize_items_for_week(week, year)
    expected = summary["completed"]

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
            browser, registreringsdato=registreringsdato
        )

    finally:
        browser.quit()

    log_summary(week, year, summary, found, registreringsdato)

    return expected, found
