"""This module contains configuration constants used across the framework"""

import os

# The number of times the robot retries on an error before terminating.
MAX_RETRY = 3

# ----------------------
# Queue population settings
# ----------------------
MAX_CONCURRENCY = 100  # tune based on backend capacity
MAX_RETRIES = 3  # transient failure retries per item
RETRY_BASE_DELAY = 0.5  # seconds (exponential backoff)

# Whether the robot should be marked as failed if MAX_RETRY_COUNT is reached.
FAIL_ROBOT_ON_TOO_MANY_ERRORS = True

# Error screenshot config
SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
SMTP_PORT = 25
SCREENSHOT_SENDER = "robot@friend.dk"

# Constant/Credential names
ERROR_EMAIL = "Error Email"

SERVICE_NOW_API_DEV_USER = "service_now_dev_user"
SERVICE_NOW_API_PROD_USER = "service_now_prod_user"

# Queue specific configs
# ----------------------

# The limit on how many queue elements to process
MAX_TASK_COUNT = 100

# ----------------------

PATH = "C:\\tmp\\Koerselsgodtgoerelse"

SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com"

# SHAREPOINT_SITE_NAME = "MBU-RPA-Egenbefordring"
SHAREPOINT_SITE_NAME = "MBURPA"

DOCUMENT_LIBRARY = "Delte dokumenter"

SHAREPOINT_KWARGS = {
    "tenant": os.getenv("TENANT"),
    "client_id": os.getenv("CLIENT_ID"),
    "thumbprint": os.getenv("APPREG_THUMBPRINT"),
    "cert_path": os.getenv("GRAPH_CERT_PEM"),
    "site_url": f"{SHAREPOINT_SITE_URL}/",
    "site_name": SHAREPOINT_SITE_NAME,
    "document_library": DOCUMENT_LIBRARY,
}

# FOLDER_NAME = "General/Til udbetaling"
FOLDER_NAME = "Egenbefordring/Til udbetaling"
