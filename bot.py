import os
import re
import base64
import secrets
import threading
from io import BytesIO
from urllib.parse import quote_plus
import warnings
import requests
import msal
from datetime import datetime
from zoneinfo import ZoneInfo

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, FileResponse, Response
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
try:
    from telegram.warnings import PTBUserWarning
except Exception:
    PTBUserWarning = Warning

warnings.filterwarnings(
    "ignore",
    message=".*CallbackQueryHandler will not be tracked for every message.*",
    category=PTBUserWarning,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)
import uvicorn

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

BOT_TOKEN = os.getenv("BOT_TOKEN")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
HELPDESK_CHAT_ID = os.getenv("HELPDESK_CHAT_ID")
SIGNATURE_BASE_URL = os.getenv("SIGNATURE_BASE_URL")
PORT = int(os.getenv("PORT", "8000"))
BUILD_VERSION = "receipt-upload-silent-userlocked-fix-v1"

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"
DAY_LOGS_LIST = "Engineer Day Logs"
BUG_IDEAS_LIST = "Bug Ideas"
BOT_USERS_LIST = "Bot Users"
SITES_LIST = "Sites"


PHOTO_LIBRARY = "Documents"
PHOTO_BASE_FOLDER = "15 - ENGINEER JOB PHOTOS"
SIGNATURE_BASE_FOLDER = "16 - CLIENT SIGNATURES"
VAN_CHECK_PHOTO_BASE_FOLDER = "17 - VAN CHECK PHOTOS"
WORKSHEET_BASE_FOLDER = "18 - JOB WORKSHEETS"
RECEIPT_BASE_FOLDER = "19 - RECEIPTS"

DAY_ACTIVE_STATUS = "Active"
DAY_CLOSED_STATUS = "Closed"

MENU_START_DAY = "🟢 Start Day"
MENU_MY_JOBS = "📋 My Jobs"
MENU_END_DAY = "🏁 End Day"
MENU_BUG_IDEA = "🐞 Bug / Ideas"
MENU_UPLOAD_RECEIPTS = "🧾 Upload Receipts"
MENU_HELPDESK = "🧰 Helpdesk"
MENU_LOG_JOB = "➕ Log Job"
MENU_REASSIGN_JOB = "🔁 Reassign Job"
MENU_OPEN_JOBS = "📋 Open Jobs"
MENU_FIND_JOB = "🔎 Find Job"
MENU_CANCEL_JOB = "❌ Cancel Job"
MENU_DELETE_JOB = "🗑 Delete Job"
MENU_ENGINEER_MENU = "👷 Engineer Menu"


UK_TZ = ZoneInfo("Europe/London")


VAN_CHECK_INTERVAL_DAYS = 14
AWAITING_DEPLOYMENT_STATUS = "Awaiting Dispatch"
LEGACY_AWAITING_DEPLOYMENT_STATUS = "Awaiting Deployment"
ASSIGNED_STATUS = "Assigned"
TRAVELLING_STATUS = "Travelling"
ON_SITE_STATUS = "On Site"
COMPLETED_STATUS = "Completed"

WORK_COMPLETED = 0
MATERIALS_USED = 1
FOLLOW_ON_REQUIRED = 2
FOLLOW_ON_NOTES = 3
ENGINEER_NOTES = 4
PHOTOS = 5
SIGNATURE_REQUIRED = 6
SIGNATURE_WAITING = 7
REVIEW = 8

START_DAY_CONFIRM = 19
START_DAY_VAN_REG = 20
START_DAY_START_MILEAGE = 21
START_DAY_VAN_CHECK = 22
START_DAY_VAN_PHOTOS = 23
END_DAY_CONFIRM = 24
END_DAY_MILEAGE = 25
BUG_IDEA_TEXT = 26

LOGJOB_CDR_NUMBER = 27
LOGJOB_CUSTOMER_NAME = 28
LOGJOB_CUSTOMER_ADDRESS = 29
LOGJOB_SITE_NAME = 30
LOGJOB_SITE_ADDRESS = 31
LOGJOB_CONTACT = 32
LOGJOB_TASK = 33
LOGJOB_NOTES = 34
LOGJOB_DATE = 35
LOGJOB_TIME = 36
LOGJOB_CATEGORY = 37
LOGJOB_ORDER_NUMBER = 38
LOGJOB_ASSIGN_ENGINEERS = 39
LOGJOB_REVIEW = 40

REASSIGN_CDR_NUMBER = 41
REASSIGN_REMOVE_ENGINEERS = 42
REASSIGN_ASSIGN_ENGINEERS = 43
REASSIGN_REASON = 44
REASSIGN_REVIEW = 45

LOGJOB_SITE_CONFIRM = 53
LOGJOB_SITE_NOTES = 54

CANCELJOB_CDR_NUMBER = 55
CANCELJOB_CONFIRM = 56
DELETEJOB_CDR_NUMBER = 57
DELETEJOB_CONFIRM = 58
RECEIPT_ENGINEER_NAME = 59
RECEIPT_DATE = 60
RECEIPT_UPLOADS = 61
ABORTJOB_REASON = 62
ABORTJOB_NOTES = 63

FINDJOB_SEARCH = 46
FINDJOB_SELECT = 47
OPENJOBS_FILTER = 48
OPENJOBS_SELECT = 49
VAN_CHECK_QUESTIONS = [
    "Are the tyres in good condition and correctly inflated? Reply Yes or No.",
    "Are all lights working? Reply Yes or No.",
    "Are windscreen, mirrors and wipers okay? Reply Yes or No.",
    "Is there any new visible damage to the van? Reply No, or describe the damage.",
    "Is the van clean, tidy and safe to work from? Reply Yes or No.",
    "Any defects or issues to report? Reply No, or describe the issue.",
]

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=authority,
    client_credential=CLIENT_SECRET,
)

web_app = FastAPI()


def get_headers(content_type=True):
    token_result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in token_result:
        raise Exception(f"Could not get Microsoft token: {token_result}")

    headers = {"Authorization": f"Bearer {token_result['access_token']}"}

    if content_type:
        headers["Content-Type"] = "application/json"

    return headers


def now_log_time():
    return datetime.now(UK_TZ).strftime("%d/%m/%Y %H:%M")


def graph_datetime_now():
    return datetime.now(UK_TZ).isoformat()


def format_sharepoint_date(value):
    if not value:
        return ""

    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.astimezone(UK_TZ).strftime("%d/%m/%Y")
    except Exception:
        return value


def sharepoint_date_to_uk_date(value):
    if not value:
        return None

    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.astimezone(UK_TZ).date()
    except Exception:
        return None


def get_site_id():
    site_hostname = SHAREPOINT_SITE.split("/")[2]
    site_path = "/" + "/".join(SHAREPOINT_SITE.split("/")[3:])
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"

    response = requests.get(site_url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(f"Could not get SharePoint site: {response.text}")

    return response.json()["id"]


def get_list_id(site_id, list_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    response = requests.get(url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(f"Could not get lists: {response.text}")

    for lst in response.json()["value"]:
        if lst["name"] == list_name:
            return lst["id"]

    raise Exception(f"List not found: {list_name}")


def get_list_items(site_id, list_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields"
    response = requests.get(url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(f"Could not get items: {response.text}")

    return response.json()["value"]


def get_list_columns(site_id, list_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"
    response = requests.get(url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(f"Could not get list columns: {response.text}")

    return response.json().get("value", [])


def normalise_field_name(value):
    return "".join(str(value or "").lower().replace("_x0020_", "").split())


def build_field_payload_for_list(site_id, list_id, fields):
    """
    Build a SharePoint fields payload using real writable internal column names.

    Important:
    - Title must always write to the real editable Title field.
    - SharePoint also exposes read-only display fields such as LinkTitle and
      LinkTitleNoMenu. These must never be written to.
    """
    columns = get_list_columns(site_id, list_id)
    lookup = {"title": "Title"}

    for column in columns:
        internal_name = column.get("name", "")
        display_name = column.get("displayName", "")
        read_only = column.get("readOnly", False)

        if read_only:
            continue

        # Never allow LinkTitle fields, even if Graph does not mark them read-only.
        if internal_name in ["LinkTitle", "LinkTitleNoMenu"]:
            continue

        if internal_name == "Title":
            lookup[normalise_field_name("Title")] = "Title"
            continue

        for key in [internal_name, display_name]:
            normalised = normalise_field_name(key)
            if normalised and normalised not in lookup:
                lookup[normalised] = internal_name

    payload = {}

    for desired_name, value in fields.items():
        if desired_name == "Title":
            internal_name = "Title"
        else:
            internal_name = lookup.get(normalise_field_name(desired_name))

        if internal_name:
            payload[internal_name] = value
        else:
            print(f"WARNING: SharePoint column not found: {desired_name}. Field skipped.")

    return payload


def update_list_item_fields(site_id, list_id, item_id, fields_to_update):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"

    response = requests.patch(
        url,
        headers=get_headers(),
        json=fields_to_update,
    )

    if response.status_code not in [200, 204]:
        raise Exception(f"Could not update item {item_id}: {response.text}")


def create_list_item_fields(site_id, list_id, fields_to_create):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"

    response = requests.post(
        url,
        headers=get_headers(),
        json={"fields": fields_to_create},
    )

    if response.status_code not in [200, 201]:
        raise Exception(f"Could not create list item: {response.text}")

    return response.json()


def delete_list_item(site_id, list_id, item_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}"

    response = requests.delete(url, headers=get_headers())

    if response.status_code not in [200, 202, 204]:
        raise Exception(f"Could not delete list item {item_id}: {response.text}")




def clear_engineer_assignment_payload():
    return {
        "EngineerLookupId@odata.type": "Collection(Edm.Int32)",
        "EngineerLookupId": [],
    }


def remove_current_engineer_assignment_payload(fields, current_lookup_id):
    remaining_ids = []
    engineer_values = fields.get("Engineer", [])

    if isinstance(engineer_values, list):
        for engineer in engineer_values:
            lookup_id = str(engineer.get("LookupId"))

            if lookup_id and lookup_id != str(current_lookup_id):
                remaining_ids.append(int(lookup_id))

    return {
        "EngineerLookupId@odata.type": "Collection(Edm.Int32)",
        "EngineerLookupId": remaining_ids,
    }


def append_engineer_log(fields, engineer_name, action, extra_text=""):
    existing_log = fields.get("EngineerVisitLog", "") or ""
    line = f"{now_log_time()} - {engineer_name} - {action}"

    if extra_text:
        line += f" - {extra_text}"

    if existing_log.strip():
        return existing_log.strip() + "\n" + line

    return line


def get_current_engineer_visit_log_lines(fields, engineer_name):
    """
    Return only the log lines for the engineer's current visit attempt.

    A job can be marked Revisit Required / No Access, returned to the office,
    and later re-dispatched to the same engineer. The old visit history must
    stay in SharePoint for audit, but it must not block the engineer from
    clicking Travelling / On Site again on the new visit.
    """
    log = fields.get("EngineerVisitLog", "") or ""
    lines = [line for line in log.splitlines() if line.strip()]

    reset_actions = ["Completed", "No Access", "Revisit Required"]
    last_reset_index = -1

    for index, line in enumerate(lines):
        if f" - {engineer_name} - " not in line:
            continue

        for reset_action in reset_actions:
            if f" - {engineer_name} - {reset_action}" in line:
                last_reset_index = index
                break

    return lines[last_reset_index + 1:]


def engineer_has_logged(fields, engineer_name, action):
    current_visit_lines = get_current_engineer_visit_log_lines(fields, engineer_name)
    search_text = f" - {engineer_name} - {action}"
    return any(search_text in line for line in current_visit_lines)


def can_click_action(fields, engineer_name, action):
    return validate_job_action(fields, engineer_name, action)


def get_sharepoint_data():
    site_id = get_site_id()
    engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
    jobs_list_id = get_list_id(site_id, JOBS_LIST)

    engineers = get_list_items(site_id, engineers_list_id)
    jobs = get_list_items(site_id, jobs_list_id)

    return site_id, engineers_list_id, jobs_list_id, engineers, jobs


def build_engineer_maps(engineers):
    by_telegram_id = {}
    by_lookup_id = {}

    for engineer in engineers:
        fields = engineer["fields"]

        lookup_id = str(fields.get("id", ""))
        name = fields.get("EngineerName", "")
        telegram_id = str(fields.get("TelegramID", ""))

        if lookup_id and name and telegram_id:
            by_lookup_id[lookup_id] = {
                "name": name,
                "telegram_id": telegram_id,
            }

            by_telegram_id[telegram_id] = {
                "lookup_id": lookup_id,
                "name": name,
            }

    return by_telegram_id, by_lookup_id


def get_engineer_menu(include_helpdesk_menu=False):
    rows = [
        [MENU_START_DAY, MENU_MY_JOBS],
        [MENU_END_DAY, MENU_BUG_IDEA],
        [MENU_UPLOAD_RECEIPTS],
    ]

    # Only Admin users should be able to switch from the Engineer menu
    # into the Helpdesk menu. Helpdesk users stay in the Helpdesk menu,
    # and Engineer users stay in the Engineer menu.
    if include_helpdesk_menu:
        rows.append([MENU_HELPDESK])

    return ReplyKeyboardMarkup(
        rows,
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def get_helpdesk_menu(include_engineer_menu=False):
    rows = [
        [MENU_LOG_JOB, MENU_REASSIGN_JOB],
        [MENU_OPEN_JOBS, MENU_FIND_JOB],
        [MENU_CANCEL_JOB],
        [MENU_BUG_IDEA, MENU_UPLOAD_RECEIPTS],
    ]

    if include_engineer_menu:
        rows.append([MENU_DELETE_JOB])
        rows.append([MENU_ENGINEER_MENU])

    return ReplyKeyboardMarkup(
        rows,
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def get_admin_menu():
    # Admin starts on the Engineer menu, with a switch button into Helpdesk.
    # This keeps the two menus separate instead of showing every option at once.
    return get_engineer_menu(include_helpdesk_menu=True)


def get_main_menu(role="Engineer"):
    role = str(role or "Engineer").strip().lower()

    if role == "admin":
        return get_admin_menu()

    if role == "helpdesk":
        return get_helpdesk_menu()

    if role == "inactive":
        return ReplyKeyboardMarkup([["/start"]], resize_keyboard=True, one_time_keyboard=False)

    return get_engineer_menu()


def get_fallback_admin_ids():
    raw_ids = os.getenv("HELPDESK_ADMIN_TELEGRAM_IDS", "")
    return {value.strip() for value in raw_ids.split(",") if value.strip()}


def get_bot_user_role(site_id, telegram_id):
    """
    Returns Engineer, Helpdesk or Admin for a Telegram user.

    This version uses the existing SharePoint 'Engineers' list as the user table.
    Required columns on Engineers:
    - EngineerName
    - TelegramID
    - Role: Engineer / Helpdesk / Admin
    - Active: Yes/No

    If the Telegram ID is not found, the user is treated as inactive/no access,
    unless the Telegram ID is listed in the HELPDESK_ADMIN_TELEGRAM_IDS Railway variable.
    """
    telegram_id = str(telegram_id).strip()

    if telegram_id in get_fallback_admin_ids():
        return "Admin"

    try:
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        users = get_list_items(site_id, engineers_list_id)
    except Exception as e:
        print(f"Engineers list unavailable while checking role; falling back to Engineer role: {e}")
        return "Engineer"

    for user in users:
        fields = user.get("fields", {})
        user_telegram_id = str(
            get_field_value(fields, "TelegramID", "Telegram ID") or ""
        ).strip()

        if user_telegram_id != telegram_id:
            continue

        active_value = get_field_value(fields, "Active")
        if active_value not in [None, ""] and not bool_field(active_value):
            return "Inactive"

        role = str(get_field_value(fields, "Role") or "Engineer").strip()

        if role.lower() in ["admin", "helpdesk", "engineer"]:
            return role.title()

        return "Engineer"

    print(f"Telegram ID {telegram_id} not found in Engineers list; no access granted.")
    return "Inactive"


def user_can_use_helpdesk(role):
    return str(role or "").strip().lower() in ["helpdesk", "admin"]


JOB_CATEGORY_CHOICES = [
    "Electrical",
    "Mechanical",
    "Plumbing",
    "HVAC",
    "Fire",
    "Building Fabric",
    "Other",
]


def is_blank_or_skip(value):
    return str(value or "").strip().lower() in ["", "skip", "none", "n/a", "na"]


def normalise_site_lookup_text(value):
    """Normalise site names for forgiving matching without being too clever."""
    text = str(value or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def site_match_score(search_text, candidate_text):
    search = normalise_site_lookup_text(search_text)
    candidate = normalise_site_lookup_text(candidate_text)

    if not search or not candidate:
        return 0

    if search == candidate:
        return 100

    if search in candidate or candidate in search:
        return 90

    search_words = set(search.split())
    candidate_words = set(candidate.split())
    if search_words and candidate_words:
        overlap = len(search_words.intersection(candidate_words)) / max(len(search_words), len(candidate_words))
    else:
        overlap = 0

    from difflib import SequenceMatcher
    ratio = SequenceMatcher(None, search, candidate).ratio()
    return int(max(overlap, ratio) * 100)


def extract_site_fields_from_sites_item(item):
    fields = item.get("fields", {})
    return {
        "source": "Sites",
        "item_id": item.get("id"),
        "site_name": str(get_field_value(fields, "SiteName", "Site Name", "Title") or "").strip(),
        "address": str(get_field_value(fields, "Address", "SiteAddress", "Site Address") or "").strip(),
        "customer_name": str(get_field_value(fields, "CustomerName", "Customer Name") or "").strip(),
        "customer_address": str(get_field_value(fields, "CustomerAddress", "Customer Address") or "").strip(),
        "notes": str(get_field_value(fields, "Notes", "SiteNotes", "Site Notes") or "").strip(),
    }


def extract_site_fields_from_job_item(item):
    fields = item.get("fields", {})
    return {
        "source": "Previous Jobs",
        "item_id": item.get("id"),
        "site_name": str(get_field_value(fields, "SiteName", "Site Name") or "").strip(),
        "address": str(get_field_value(fields, "Address", "SiteAddress", "Site Address") or "").strip(),
        "customer_name": str(get_field_value(fields, "CustomerName", "Customer Name") or "").strip(),
        "customer_address": str(get_field_value(fields, "CustomerAddress", "Customer Address") or "").strip(),
        "notes": "",
    }


def find_best_site_candidate(site_id, jobs_list_id, site_name):
    """
    Search the master Sites list first, then fall back to previous Engineer Jobs.
    Returns a dict with source/site_name/address/customer_name/notes, or None.
    """
    best = None
    best_score = 0

    try:
        sites_list_id = get_list_id(site_id, SITES_LIST)
        site_items = get_list_items(site_id, sites_list_id)
        for item in site_items:
            fields = item.get("fields", {})
            active_value = get_field_value(fields, "Active")
            if active_value not in [None, ""] and not bool_field(active_value):
                continue

            candidate = extract_site_fields_from_sites_item(item)
            if not candidate["site_name"] or not candidate["address"]:
                continue

            score = site_match_score(site_name, candidate["site_name"])
            if score > best_score:
                best = candidate
                best_score = score
    except Exception as e:
        print(f"INFO: Sites list lookup skipped: {e}")

    if best and best_score >= 70:
        return best

    try:
        jobs_data = get_list_items(site_id, jobs_list_id)
        for item in jobs_data:
            candidate = extract_site_fields_from_job_item(item)
            if not candidate["site_name"] or not candidate["address"]:
                continue

            score = site_match_score(site_name, candidate["site_name"])
            if score > best_score:
                best = candidate
                best_score = score
    except Exception as e:
        print(f"INFO: Previous job site lookup skipped: {e}")

    if best and best_score >= 70:
        return best

    return None


def format_site_candidate(candidate):
    notes = candidate.get("notes", "")
    notes_block = f"\n\nSite notes:\n{notes}" if notes else ""
    return (
        f"I found a possible saved site from {candidate.get('source', 'records')}:\n\n"
        f"Site: {candidate.get('site_name', '')}\n"
        f"Address: {candidate.get('address', '')}\n"
        f"Customer: {candidate.get('customer_name', '') or 'N/A'}\n"
        f"Customer Address: {candidate.get('customer_address', '') or 'N/A'}"
        f"{notes_block}\n\n"
        "Use this address?\n\n"
        "Reply YES to use it, EDIT to change the address, or NO to enter manually."
    )


def upsert_site_record_from_job(job):
    """Create/update the Sites list after a successful job log. Non-critical: failures only warn."""
    site_id = job.get("site_id")
    if not site_id:
        return

    site_name = str(job.get("site_name", "")).strip()
    address = str(job.get("site_address", "")).strip()
    customer_name = str(job.get("customer_name", "")).strip()
    customer_address = str(job.get("customer_address", "")).strip()
    notes = str(job.get("site_notes", "")).strip()

    if not site_name or not address:
        return

    try:
        sites_list_id = get_list_id(site_id, SITES_LIST)
        site_items = get_list_items(site_id, sites_list_id)

        existing = None
        best_score = 0
        for item in site_items:
            candidate = extract_site_fields_from_sites_item(item)
            score = site_match_score(site_name, candidate.get("site_name", ""))
            if score > best_score:
                existing = item
                best_score = score

        fields = {
            "Title": site_name,
            "SiteName": site_name,
            "Site Name": site_name,
            "Address": address,
            "CustomerName": customer_name,
            "Customer Name": customer_name,
            "CustomerAddress": customer_address,
            "Customer Address": customer_address,
            "Active": True,
        }
        if notes:
            fields["Notes"] = notes

        payload = build_field_payload_for_list(site_id, sites_list_id, fields)

        if existing and best_score >= 90:
            update_list_item_fields(site_id, sites_list_id, existing["id"], payload)
        else:
            create_list_item_fields(site_id, sites_list_id, payload)

    except Exception as e:
        print(f"WARNING: Could not update Sites list for {site_name}: {e}")


def prompt_for_site_notes(existing_notes=""):
    if existing_notes:
        return (
            "Any permanent site notes to save or update?\n\n"
            f"Current notes:\n{existing_notes}\n\n"
            "Reply with updated notes, or type SKIP to leave them as they are."
        )

    return (
        "Any permanent site notes to save for next time?\n\n"
        "Examples: keys at reception, parking instructions, plant room location, asbestos register location.\n\n"
        "Type notes, or type SKIP."
    )


def parse_helpdesk_job_date(value):
    text = str(value or "").strip().lower()
    today = datetime.now(UK_TZ).date()

    if text in ["today", "tod"]:
        return today.isoformat()

    if text in ["tomorrow", "tmr", "tom"]:
        from datetime import timedelta
        return (today + timedelta(days=1)).isoformat()

    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"]:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except Exception:
            pass

    return None


def normalise_helpdesk_time(value):
    text = str(value or "").strip().lower().replace(".", ":")

    if text in ["now", "asap", "urgent"]:
        return datetime.now(UK_TZ).strftime("%H:%M")

    # 800 -> 08:00, 0830 -> 08:30, 1700 -> 17:00
    if text.isdigit() and len(text) in [3, 4]:
        if len(text) == 3:
            text = "0" + text
        text = text[:2] + ":" + text[2:]

    try:
        parsed = datetime.strptime(text, "%H:%M")
        return parsed.strftime("%H:%M")
    except Exception:
        return None


def get_active_assignable_engineers(engineers):
    assignable = []

    for item in engineers:
        fields = item.get("fields", {})
        lookup_id = str(fields.get("id", "") or item.get("id", ""))
        name = get_field_value(fields, "EngineerName", "Engineer Name", "Title") or ""
        telegram_id = str(get_field_value(fields, "TelegramID", "Telegram ID") or "").strip()
        role = str(get_field_value(fields, "Role") or "Engineer").strip().lower()
        active_value = get_field_value(fields, "Active")

        if active_value not in [None, ""] and not bool_field(active_value):
            continue

        # Admins can be assigned if needed, but Helpdesk-only users should not receive engineering jobs.
        if role not in ["engineer", "admin"]:
            continue

        if lookup_id and name and telegram_id:
            assignable.append({
                "lookup_id": lookup_id,
                "name": str(name),
                "telegram_id": telegram_id,
            })

    assignable.sort(key=lambda e: e["name"].lower())
    return assignable


def format_engineer_selection_list(engineers):
    lines = []
    for index, engineer in enumerate(engineers, start=1):
        lines.append(f"{index}. {engineer['name']}")
    return "\n".join(lines)


def parse_engineer_selection(text, engineers):
    selected = []
    seen = set()
    parts = [part.strip() for part in str(text or "").replace(";", ",").split(",") if part.strip()]

    for part in parts:
        if not part.isdigit():
            return None, "Please reply with engineer number(s), for example 1 or 1,3."

        index = int(part)
        if index < 1 or index > len(engineers):
            return None, f"Engineer number {index} is not in the list."

        engineer = engineers[index - 1]
        if engineer["lookup_id"] not in seen:
            selected.append(engineer)
            seen.add(engineer["lookup_id"])

    if not selected:
        return None, "Please select at least one engineer."

    return selected, ""


def build_log_job_review(job):
    assigned_names = ", ".join(engineer["name"] for engineer in job.get("assigned_engineers", [])) or "None"
    return (
        "Please review the new job before I create it in SharePoint:\n\n"
        f"CDR Number: {job.get('cdr_number', '')}\n"
        f"Customer: {job.get('customer_name', '')}\n"
        f"Customer Address: {job.get('customer_address', '')}\n"
        f"Site: {job.get('site_name', '')}\n"
        f"Site Address: {job.get('site_address', '')}\n"
        f"Site Notes: {job.get('site_notes', '') or 'N/A'}\n"
        f"Contact: {job.get('contact', '') or 'N/A'}\n"
        f"Task: {job.get('task', '')}\n"
        f"Notes: {job.get('notes', '') or 'N/A'}\n"
        f"Date: {job.get('date_display', job.get('date', ''))}\n"
        f"Time: {job.get('time', '')}\n"
        f"Category: {job.get('category', '')}\n"
        f"Order Number: {job.get('order_number', '') or 'N/A'}\n"
        f"Assigned Engineer(s): {assigned_names}\n\n"
        "Reply YES to create and send, NO to cancel, or RESTART to start again."
    )


def build_helpdesk_job_fields(site_id, jobs_list_id, job, telegram_notified=False):
    job_notes = job.get("notes", "") or ""
    site_notes = job.get("site_notes", "") or ""

    if site_notes:
        if job_notes:
            job_notes = f"{job_notes}\n\nSite notes: {site_notes}"
        else:
            job_notes = f"Site notes: {site_notes}"

    payload = build_field_payload_for_list(
        site_id,
        jobs_list_id,
        {
            "Title": job["cdr_number"],
            "CDRNumber": job["cdr_number"],
            "CDR Number": job["cdr_number"],
            "Date": job["date"],
            "StartTime": job["time"],
            "Start Time": job["time"],
            "CustomerName": job["customer_name"],
            "Customer Name": job["customer_name"],
            "CustomerAddress": job["customer_address"],
            "Customer Address": job["customer_address"],
            "SiteName": job["site_name"],
            "Site Name": job["site_name"],
            "Address": job["site_address"],
            "ContactName": job.get("contact", ""),
            "Contact Name": job.get("contact", ""),
            "Task": job["task"],
            "Notes": job_notes,
            "CustomerOrderNumber": job.get("order_number", ""),
            "Customer Order Number": job.get("order_number", ""),
            "JobCategory": job.get("category", ""),
            "Job Category": job.get("category", ""),
            "Status": ASSIGNED_STATUS,
            "TelegramNotified": bool(telegram_notified),
            "Telegram Notified": bool(telegram_notified),
            "WorksheetGenerated": False,
            "Worksheet Generated": False,
            "WorksheetSubmitted": False,
            "Worksheet Submitted": False,
            "JobOutcome": "",
            "Job Outcome": "",
            "EngineerVisitLog": f"{now_log_time()} - Helpdesk - Job logged via Telegram",
            "Engineer Visit Log": f"{now_log_time()} - Helpdesk - Job logged via Telegram",
        },
    )

    engineer_ids = [int(engineer["lookup_id"]) for engineer in job.get("assigned_engineers", [])]
    if engineer_ids:
        payload["EngineerLookupId@odata.type"] = "Collection(Edm.Int32)"
        payload["EngineerLookupId"] = engineer_ids


    return payload


def get_current_assigned_engineers_from_job(fields, engineers):
    """Return current assigned engineers with name, lookup_id and telegram_id where possible."""
    assigned = []
    assigned_ids = set(get_assigned_engineer_ids(fields))

    by_lookup = {}
    for item in engineers:
        item_fields = item.get("fields", {})
        lookup_id = str(item_fields.get("id", "") or item.get("id", ""))
        if not lookup_id:
            continue
        by_lookup[lookup_id] = {
            "lookup_id": lookup_id,
            "name": str(get_field_value(item_fields, "EngineerName", "Engineer Name", "Title") or f"Engineer {lookup_id}"),
            "telegram_id": str(get_field_value(item_fields, "TelegramID", "Telegram ID") or "").strip(),
        }

    # Prefer the SharePoint lookup display values on the job, then enrich with Engineers list data.
    engineer_values = fields.get("Engineer", [])
    if isinstance(engineer_values, list):
        for engineer_value in engineer_values:
            lookup_id = str(engineer_value.get("LookupId") or "")
            if not lookup_id:
                continue
            details = by_lookup.get(lookup_id, {})
            assigned.append({
                "lookup_id": lookup_id,
                "name": details.get("name") or str(engineer_value.get("LookupValue") or f"Engineer {lookup_id}"),
                "telegram_id": details.get("telegram_id", ""),
            })

    # Fallback if the lookup field shape changes but IDs are still available.
    seen = {engineer["lookup_id"] for engineer in assigned}
    for lookup_id in assigned_ids:
        if lookup_id not in seen:
            assigned.append(by_lookup.get(lookup_id, {
                "lookup_id": lookup_id,
                "name": f"Engineer {lookup_id}",
                "telegram_id": "",
            }))

    assigned.sort(key=lambda e: e["name"].lower())
    return assigned


def parse_remove_engineer_selection(text, current_engineers):
    value = str(text or "").strip().lower()

    if value in ["all", "remove all", "everyone", "both"]:
        return current_engineers[:], ""

    if value in ["none", "no", "skip", "0"]:
        return [], ""

    selected, error = parse_engineer_selection(text, current_engineers)
    if error:
        return None, "Reply with engineer number(s), ALL, or NONE."
    return selected, ""


def build_reassign_review(data):
    current = ", ".join(e["name"] for e in data.get("current_engineers", [])) or "None"
    removing = ", ".join(e["name"] for e in data.get("remove_engineers", [])) or "None"
    assigning = ", ".join(e["name"] for e in data.get("assign_engineers", [])) or "None"
    final = ", ".join(e["name"] for e in data.get("final_engineers", [])) or "None"
    fields = data.get("job_fields", {})

    return (
        "Please review this reassignment before I update SharePoint:\n\n"
        f"CDR Number: {fields.get('CDRNumber', data.get('cdr_number', ''))}\n"
        f"Site: {fields.get('SiteName', '')}\n"
        f"Current engineer(s): {current}\n"
        f"Remove: {removing}\n"
        f"Assign/send to: {assigning}\n"
        f"Final assigned engineer(s): {final}\n"
        f"Reason: {data.get('reason', '') or 'N/A'}\n\n"
        "Reply YES to reassign and send, NO to cancel, or RESTART to start again."
    )


async def send_created_job_to_engineers(bot, item_id, fields, assigned_engineers):
    sent_to_any = False
    failed = []

    for engineer in assigned_engineers:
        try:
            await bot.send_message(
                chat_id=engineer["telegram_id"],
                text="New job assigned:\n\n" + format_job(fields, engineer["name"]),
                reply_markup=get_job_buttons(item_id, fields.get("Address", "")),
            )
            sent_to_any = True
        except Exception as e:
            failed.append(f"{engineer['name']}: {e}")
            print(f"WARNING: Could not send newly logged job {fields.get('CDRNumber', item_id)} to {engineer['name']}: {e}")

    return sent_to_any, failed


async def get_role_for_update(update):
    try:
        site_id = get_site_id()
        return get_bot_user_role(site_id, update.effective_user.id)
    except Exception as e:
        print(f"Could not determine bot user role: {e}")
        return "Engineer"


def get_today_iso():
    return datetime.now(UK_TZ).date().isoformat()


def safe_folder_name(value):
    cleaned = "".join(ch for ch in str(value).strip().upper() if ch.isalnum() or ch in ["-", "_"])
    return cleaned or "UNKNOWN"


def upload_van_check_photo_to_sharepoint(site_id, work_date, van_reg, file_name, file_bytes):
    folder_name = f"{work_date}/{safe_folder_name(van_reg)}"
    return upload_file_to_sharepoint(
        site_id,
        VAN_CHECK_PHOTO_BASE_FOLDER,
        folder_name,
        file_name,
        file_bytes,
    )


def upload_receipt_to_sharepoint(site_id, receipt_date, engineer_name, file_name, file_bytes):
    folder_name = f"{receipt_date}/{safe_folder_name(engineer_name)}"
    return upload_file_to_sharepoint(
        site_id,
        RECEIPT_BASE_FOLDER,
        folder_name,
        file_name,
        file_bytes,
    )


def clean_receipt_file_name(value):
    cleaned = "".join(ch for ch in str(value or "receipt").strip() if ch.isalnum() or ch in ["-", "_", "."])
    return cleaned or "receipt"


def get_engineer_for_telegram_id(telegram_id):
    site_id = get_site_id()
    engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
    engineers = get_list_items(site_id, engineers_list_id)
    engineers_by_telegram, _ = build_engineer_maps(engineers)
    return site_id, engineers_list_id, engineers, engineers_by_telegram.get(str(telegram_id))


def find_active_day_log(day_logs, telegram_id, work_date=None):
    work_date = work_date or get_today_iso()

    for log in day_logs:
        fields = log.get("fields", {})

        raw_work_date = fields.get("WorkDate", "")
        parsed_work_date = sharepoint_date_to_uk_date(raw_work_date)

        if parsed_work_date:
            log_date = parsed_work_date.isoformat()
        else:
            log_date = str(raw_work_date)[:10]

        status = str(fields.get("Status", ""))
        log_telegram_id = str(fields.get("EngineerTelegramID", ""))

        if (
            log_telegram_id == str(telegram_id)
            and log_date == work_date
            and status == DAY_ACTIVE_STATUS
        ):
            return log

    return None


def get_active_day_for_engineer(site_id, telegram_id):
    day_logs_list_id = get_list_id(site_id, DAY_LOGS_LIST)
    day_logs = get_list_items(site_id, day_logs_list_id)
    return day_logs_list_id, find_active_day_log(day_logs, telegram_id)


def engineer_has_active_day(site_id, telegram_id):
    _, active_day = get_active_day_for_engineer(site_id, telegram_id)
    return active_day is not None


def normalise_mileage(value):
    mileage = str(value or "").strip().replace(",", "")

    try:
        number = float(mileage)
    except Exception:
        return None

    if number < 0:
        return None

    if number.is_integer():
        return str(int(number))

    return str(number)

def get_field_value(fields, *names):
    for name in names:
        if name in fields and fields.get(name) not in [None, ""]:
            return fields.get(name)

    normalised_names = {normalise_field_name(name) for name in names}

    for key, value in fields.items():
        if normalise_field_name(key) in normalised_names and value not in [None, ""]:
            return value

    return None


def parse_sharepoint_datetime(value):
    if not value:
        return None

    if isinstance(value, datetime):
        return value.astimezone(UK_TZ) if value.tzinfo else value.replace(tzinfo=UK_TZ)

    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).astimezone(UK_TZ)
    except Exception:
        return None


def calculate_day_pay_hours(start_dt, end_dt):
    if not start_dt or not end_dt:
        return None

    if end_dt < start_dt:
        return None

    normal_start = start_dt.replace(hour=8, minute=0, second=0, microsecond=0)
    normal_end = start_dt.replace(hour=16, minute=30, second=0, microsecond=0)

    # If the shift ends on a different date, calculate normal window for the start date only.
    # This matches the CDR standard working day rule of 08:00-16:30.
    overlap_start = max(start_dt, normal_start)
    overlap_end = min(end_dt, normal_end)

    normal_gross = 0.0
    if overlap_end > overlap_start:
        normal_gross = (overlap_end - overlap_start).total_seconds() / 3600

    total_hours = (end_dt - start_dt).total_seconds() / 3600
    break_deducted = 0.5 if normal_gross >= 0.5 else normal_gross
    normal_hours = max(0.0, normal_gross - break_deducted)
    ooh_hours = max(0.0, total_hours - normal_gross)

    return {
        "total_hours": round(total_hours, 2),
        "normal_hours": round(normal_hours, 2),
        "ooh_hours": round(ooh_hours, 2),
        "break_deducted": round(break_deducted, 2),
    }


def build_pay_summary(start_dt, end_dt, hours):
    if not hours:
        return "Unable to calculate hours - start or end time missing."

    return (
        f"Start: {start_dt.strftime('%d/%m/%Y %H:%M')}\n"
        f"End: {end_dt.strftime('%d/%m/%Y %H:%M')}\n"
        f"Total hours: {hours['total_hours']}\n"
        f"Normal hours: {hours['normal_hours']}\n"
        f"OOH hours: {hours['ooh_hours']} at 1.5x\n"
        f"Break deducted: {hours['break_deducted']}"
    )


def update_active_day_live_status(site_id, telegram_id, status, current_job=""):
    """
    Live CurrentStatus/CurrentJob tracking has been retired.
    Kept as a no-op so older call sites cannot break the bot.
    """
    return


def get_job_reference(fields):
    cdr_number = fields.get("CDRNumber", "")
    site_name = fields.get("SiteName", "")
    if cdr_number and site_name:
        return f"{cdr_number} - {site_name}"
    return cdr_number or site_name or ""

def get_open_jobs_for_engineer_today(jobs_data, engineer_lookup_id):
    today = datetime.now(UK_TZ).date()
    open_jobs = []

    for job in jobs_data:
        fields = job.get("fields", {})
        job_date = sharepoint_date_to_uk_date(fields.get("Date", ""))
        assigned_ids = get_assigned_engineer_ids(fields)

        if str(engineer_lookup_id) in assigned_ids and job_date == today:
            status = str(fields.get("Status", ""))
            outcome = str(fields.get("JobOutcome", ""))

            # A previous No Access/Revisit outcome must not stop a re-dispatched job
            # being treated as open once the office has assigned an engineer again.
            if status != COMPLETED_STATUS and outcome != "Completed":
                open_jobs.append(job)

    return open_jobs


def format_open_jobs_for_end_day(open_jobs):
    lines = []

    for job in open_jobs[:8]:
        fields = job.get("fields", {})
        cdr = fields.get("CDRNumber", "")
        site = fields.get("SiteName", "")
        status = fields.get("Status", "")
        lines.append(f"- {cdr} | {site} | {status}")

    if len(open_jobs) > 8:
        lines.append(f"...and {len(open_jobs) - 8} more")

    return "\n".join(lines)


def get_assigned_engineer_ids(fields):
    assigned = []
    engineer_values = fields.get("Engineer", [])

    if isinstance(engineer_values, list):
        for engineer in engineer_values:
            assigned.append(str(engineer.get("LookupId")))

    return assigned


def is_notified(fields):
    value = fields.get("TelegramNotified", False)
    return value in [True, "true", "True", "Yes", "yes", "1", 1]




def get_yes_no_keyboard(prefix):
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Yes", callback_data=f"{prefix}|yes"),
            InlineKeyboardButton("❌ No", callback_data=f"{prefix}|no"),
        ]
    ])


def get_signed_skip_keyboard():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Signed", callback_data="signature_waiting|signed"),
            InlineKeyboardButton("⏭️ Skip", callback_data="signature_waiting|skip"),
        ]
    ])


def get_review_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Submit worksheet", callback_data="review|submit")],
        [InlineKeyboardButton("🔄 Restart worksheet", callback_data="review|restart")],
        [InlineKeyboardButton("❌ Cancel", callback_data="review|cancel")],
    ])


NO_ACCESS_REASONS = {
    "no_answer": "No answer",
    "no_keys": "No keys / access available",
    "contact_unreachable": "Contact unreachable",
    "site_closed": "Site closed",
    "access_refused": "Access refused",
    "parking_access_issue": "Parking / access issue",
    "other": "Other / see notes",
}

ABORT_REASONS = {
    "assigned_elsewhere": "Assigned to another job",
    "cannot_attend": "Cannot attend today",
    "ran_out_of_time": "Ran out of time",
    "wrong_engineer": "Wrong engineer / skillset",
    "vehicle_issue": "Vehicle issue",
    "other": "Other reason",
}


def get_no_access_reason_keyboard(item_id):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚪 No answer", callback_data=f"noaccess_reason|{item_id}|no_answer")],
        [InlineKeyboardButton("🔑 No keys / access", callback_data=f"noaccess_reason|{item_id}|no_keys")],
        [InlineKeyboardButton("📵 Contact unreachable", callback_data=f"noaccess_reason|{item_id}|contact_unreachable")],
        [InlineKeyboardButton("🏢 Site closed", callback_data=f"noaccess_reason|{item_id}|site_closed")],
        [InlineKeyboardButton("🚫 Access refused", callback_data=f"noaccess_reason|{item_id}|access_refused")],
        [InlineKeyboardButton("🚗 Parking / access issue", callback_data=f"noaccess_reason|{item_id}|parking_access_issue")],
        [InlineKeyboardButton("📝 Other / see notes", callback_data=f"noaccess_reason|{item_id}|other")],
        [InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_outcome|{item_id}")],
    ])


def get_abort_reason_keyboard(item_id):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📌 Assigned elsewhere", callback_data=f"abort_reason|{item_id}|assigned_elsewhere")],
        [InlineKeyboardButton("📅 Cannot attend today", callback_data=f"abort_reason|{item_id}|cannot_attend")],
        [InlineKeyboardButton("⏱ Ran out of time", callback_data=f"abort_reason|{item_id}|ran_out_of_time")],
        [InlineKeyboardButton("🛠 Wrong engineer / skillset", callback_data=f"abort_reason|{item_id}|wrong_engineer")],
        [InlineKeyboardButton("🚐 Vehicle issue", callback_data=f"abort_reason|{item_id}|vehicle_issue")],
        [InlineKeyboardButton("📝 Other reason", callback_data=f"abort_reason|{item_id}|other")],
        [InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_outcome|{item_id}")],
    ])


def get_job_buttons(item_id, address=None):
    rows = [
        [
            InlineKeyboardButton("🚗 Start Travelling", callback_data=f"status|{item_id}|Travelling"),
        ],
        [
            InlineKeyboardButton("📍 Arrived On Site", callback_data=f"status|{item_id}|On Site"),
        ],
        [
            InlineKeyboardButton("✅ Complete Job", callback_data=f"start_worksheet|{item_id}|Completed"),
        ],
        [
            InlineKeyboardButton("🔁 Revisit Required", callback_data=f"confirm_outcome|{item_id}|Revisit Required"),
        ],
        [
            InlineKeyboardButton("🚫 No Access", callback_data=f"noaccess|{item_id}"),
        ],
        [
            InlineKeyboardButton("⏹ Abort Attendance", callback_data=f"abort_job|{item_id}"),
        ],
    ]

    if address:
        maps_url = (
            "https://www.google.com/maps/search/?api=1&query="
            + quote_plus(str(address))
        )
        rows.append([
            InlineKeyboardButton("🗺 Open Maps", url=maps_url)
        ])

    return InlineKeyboardMarkup(rows)


def format_job(fields, engineer_name=None):
    return (
        f"CDR Number: {fields.get('CDRNumber', '')}\n"
        f"Date: {format_sharepoint_date(fields.get('Date', ''))}\n"
        f"Time: {fields.get('StartTime', '')}\n"
        f"Engineer: {engineer_name or ''}\n"
        f"Site: {fields.get('SiteName', '')}\n"
        f"Address: {fields.get('Address', '')}\n"
        f"Task: {fields.get('Task', '')}\n"
        f"Notes: {fields.get('Notes', '')}\n"
        f"Contact: {fields.get('ContactName', '')}"
    )


def is_closed_job(fields):
    status = str(fields.get("Status", "") or "").strip()
    outcome = str(fields.get("JobOutcome", "") or "").strip()

    open_statuses = {
        "",
        AWAITING_DEPLOYMENT_STATUS,
        LEGACY_AWAITING_DEPLOYMENT_STATUS,
        ASSIGNED_STATUS,
        "Travelling",
        "On Site",
    }

    # Completed is always final, even if SharePoint still shows an old live status.
    if status == COMPLETED_STATUS or outcome == "Completed":
        return True

    # No Access/Revisit are only final while the job is sitting back with the office.
    # If the office reassigns the job, the status becomes an open/dispatch status again
    # and the previous outcome remains as audit history, not a blocker.
    if status in open_statuses:
        return False

    return status in ["No Access", "Revisit Required"] or outcome in ["No Access", "Revisit Required"]


def has_engineer_action(fields, engineer_name, action):
    return engineer_has_logged(fields, engineer_name, action)


def validate_job_action(fields, engineer_name, action):
    """
    Hard gate for all job button actions.
    Prevents duplicate clicks, old-button actions, and out-of-order progress.
    """
    if is_closed_job(fields):
        return False, "This job has already been closed or returned to the office. No further action is required."

    has_travelled = has_engineer_action(fields, engineer_name, "Travelling")
    has_on_site = has_engineer_action(fields, engineer_name, "On Site")

    if action == "Travelling":
        if has_travelled:
            return False, "Travelling has already been logged for this job."
        return True, ""

    if action == "On Site":
        if not has_travelled:
            return False, "You need to click Travelling before clicking On Site."
        if has_on_site:
            return False, "On Site has already been logged for this job."
        return True, ""

    if action in ["No Access", "Revisit Required", "Completed"]:
        if not has_on_site:
            return False, "You need to click On Site before selecting this option."
        return True, ""

    return True, ""


def should_auto_send_job(fields):
    """
    Auto-send today's jobs only when:
    - an engineer is assigned
    - TelegramNotified is not already true
    - the job is dated today
    - the job is not finally closed

    Supports both Awaiting Dispatch and Awaiting Deployment while SharePoint is being tidied.
    """
    if is_notified(fields):
        return False

    assigned_ids = get_assigned_engineer_ids(fields)
    if not assigned_ids:
        return False

    job_date = sharepoint_date_to_uk_date(fields.get("Date", ""))
    today = datetime.now(UK_TZ).date()

    if job_date != today:
        return False

    status = str(fields.get("Status", "") or "").strip()
    outcome = str(fields.get("JobOutcome", "") or "").strip()

    allowed_statuses = {
        "",
        AWAITING_DEPLOYMENT_STATUS,
        LEGACY_AWAITING_DEPLOYMENT_STATUS,
        ASSIGNED_STATUS,
    }

    if status not in allowed_statuses:
        return False

    # Only Completed blocks automatic dispatch permanently.
    # No Access/Revisit may be previous outcomes and must allow re-dispatch
    # once the office assigns an engineer again and TelegramNotified is False.
    if outcome == "Completed":
        return False

    return True



def normalise_cdr(value):
    """
    Makes CDR matching forgiving:
    - ignores case
    - trims spaces
    - removes common prefixes like CDR:
    - ignores spaces/hyphens/underscores
    """
    value = str(value or "").strip().lower()

    for prefix in ["cdr:", "cdr number:", "cdrnumber:"]:
        if value.startswith(prefix):
            value = value[len(prefix):].strip()

    value = value.replace(" ", "").replace("-", "").replace("_", "")
    return value


def find_job_by_cdr(jobs_data, cdr_number):
    target = normalise_cdr(cdr_number)

    for item in jobs_data:
        fields = item.get("fields", {})
        possible_values = [
            fields.get("CDRNumber", ""),
            fields.get("Title", ""),
            fields.get("JobTitle", ""),
        ]

        for value in possible_values:
            if normalise_cdr(value) == target:
                return item

    return None


def find_job_by_item_id(jobs_data, item_id):
    for job in jobs_data:
        if str(job.get("id")) == str(item_id):
            return job
    return None


def get_drive_id(site_id, drive_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    response = requests.get(url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(f"Could not get SharePoint drives: {response.text}")

    for drive in response.json()["value"]:
        if drive["name"] == drive_name:
            return drive["id"]

    raise Exception(f"Document library not found: {drive_name}")


def ensure_folder(drive_id, folder_path):
    parts = folder_path.split("/")
    current_path = ""

    for part in parts:
        current_path = f"{current_path}/{part}" if current_path else part

        check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{current_path}"
        check_response = requests.get(check_url, headers=get_headers())

        if check_response.status_code == 200:
            continue

        if check_response.status_code != 404:
            raise Exception(f"Could not check folder {current_path}: {check_response.text}")

        parent_path = "/".join(current_path.split("/")[:-1])

        if parent_path:
            create_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}:/children"
        else:
            create_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"

        body = {
            "name": part,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace",
        }

        create_response = requests.post(create_url, headers=get_headers(), json=body)

        if create_response.status_code not in [200, 201]:
            raise Exception(f"Could not create folder {current_path}: {create_response.text}")


def upload_file_to_sharepoint(site_id, base_folder, cdr_number, file_name, file_bytes):
    drive_id = get_drive_id(site_id, PHOTO_LIBRARY)
    folder_path = f"{base_folder}/{cdr_number}"
    ensure_folder(drive_id, folder_path)

    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:"
        f"/{folder_path}/{file_name}:/content"
    )

    response = requests.put(
        url,
        headers=get_headers(content_type=False),
        data=file_bytes,
    )

    if response.status_code not in [200, 201]:
        raise Exception(f"Could not upload file: {response.text}")

    return response.json().get("webUrl", "")


def upload_photo_to_sharepoint(site_id, cdr_number, file_name, file_bytes):
    return upload_file_to_sharepoint(
        site_id,
        PHOTO_BASE_FOLDER,
        cdr_number,
        file_name,
        file_bytes,
    )


def upload_signature_to_sharepoint(site_id, cdr_number, signature_data_url):
    if "," not in signature_data_url:
        raise Exception("Invalid signature data")

    image_base64 = signature_data_url.split(",", 1)[1]
    image_bytes = base64.b64decode(image_base64)
    file_name = f"{cdr_number}_client_signature_{datetime.now(UK_TZ).strftime('%Y%m%d_%H%M%S')}.png"

    return upload_file_to_sharepoint(
        site_id,
        SIGNATURE_BASE_FOLDER,
        cdr_number,
        file_name,
        image_bytes,
    )


def get_job_by_cdr_and_token(cdr_number, token):
    site_id = get_site_id()
    jobs_list_id = get_list_id(site_id, JOBS_LIST)
    jobs_data = get_list_items(site_id, jobs_list_id)

    for job in jobs_data:
        fields = job["fields"]
        if (
            normalise_cdr(fields.get("CDRNumber", "")) == normalise_cdr(cdr_number)
            and str(fields.get("SignatureToken", "")) == str(token)
        ):
            return site_id, jobs_list_id, job

    return None, None, None


def bool_field(value):
    return value in [True, "true", "True", "Yes", "yes", "1", 1]


def create_signature_token_for_job(site_id, jobs_list_id, item_id):
    token = secrets.token_urlsafe(24)
    update_list_item_fields(
        site_id,
        jobs_list_id,
        item_id,
        {
            "ClientSignatureRequired": True,
            "ClientSignatureReceived": False,
            "SignatureToken": token,
        },
    )
    return token


def build_signature_url(cdr_number, token):
    if not SIGNATURE_BASE_URL:
        raise Exception("SIGNATURE_BASE_URL Railway variable is missing")

    return f"{SIGNATURE_BASE_URL.rstrip('/')}/sign/{cdr_number}?token={token}"


def build_review_text(worksheet):
    signature_required = "Yes" if worksheet.get("ClientSignatureRequired") else "No"
    outcome = worksheet.get("JobOutcome", "Completed")
    no_access_reason = worksheet.get("NoAccessReason", "")
    signature_received = "Yes" if worksheet.get("ClientSignatureReceived") else "No"
    no_access_line = f"No Access reason: {no_access_reason}\n\n" if outcome == "No Access" and no_access_reason else ""

    return (
        f"Please review worksheet for {worksheet['cdr_number']}:\n\n"
        f"Outcome: {outcome}\n\n"
        f"{no_access_line}"
        f"Work completed:\n{worksheet.get('WorkCompleted', '')}\n\n"
        f"Materials used:\n{worksheet.get('MaterialsUsed', '')}\n\n"
        f"Follow-on required:\n{'Yes' if worksheet.get('FollowOnRequired') else 'No'}\n\n"
        f"Follow-on notes:\n{worksheet.get('FollowOnNotes', '') or 'None'}\n\n"
        f"Engineer notes:\n{worksheet.get('EngineerCompletionNotes', '')}\n\n"
        f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n"
        f"Client signature required: {signature_required}\n"
        f"Client signature received: {signature_received}\n\n"
        f"Tap Submit worksheet to complete the job.\n"
        f"Tap Restart worksheet to redo the worksheet.\n"
        f"Tap Cancel to abandon it."
    )


async def notify_helpdesk(context, text):
    if HELPDESK_CHAT_ID:
        try:
            await context.bot.send_message(chat_id=HELPDESK_CHAT_ID, text=text)
        except Exception as e:
            print(f"Could not notify helpdesk: {e}")


@web_app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse("CDR Engineer Bot signature portal is online.")


@web_app.get("/logo.png")
def logo():
    logo_path = "cdr-logo.png"

    if os.path.exists(logo_path):
        return FileResponse(logo_path)

    fallback_svg = """
    <svg xmlns="http://www.w3.org/2000/svg" width="640" height="180" viewBox="0 0 640 180">
        <rect width="640" height="180" fill="white"/>
        <text x="320" y="82" text-anchor="middle" font-family="Arial, sans-serif" font-size="58" font-weight="700" fill="#f58220">CDR</text>
        <text x="320" y="125" text-anchor="middle" font-family="Arial, sans-serif" font-size="24" font-weight="600" fill="#333333">M&amp;E Services Ltd</text>
    </svg>
    """

    return Response(content=fallback_svg.strip(), media_type="image/svg+xml")


@web_app.get("/sign/{cdr_number}", response_class=HTMLResponse)
def signature_page(cdr_number: str, token: str):
    site_id, jobs_list_id, job = get_job_by_cdr_and_token(cdr_number, token)

    if not job:
        return HTMLResponse("Invalid or expired signature link.", status_code=404)

    fields = job["fields"]

    if bool_field(fields.get("ClientSignatureReceived")):
        return HTMLResponse("This job has already been signed.", status_code=200)

    site = fields.get("SiteName", "")
    address = fields.get("Address", "")
    task = fields.get("Task", "")

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Client Signature - CDR M&E Services Ltd</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <script src="https://cdn.jsdelivr.net/npm/signature_pad@4.1.6/dist/signature_pad.umd.min.js"></script>
        <style>
            html, body {{ overscroll-behavior: none; }}
            body {{ font-family: Arial, sans-serif; background: #f4f4f4; padding: 20px; margin: 0; }}
            .container {{ max-width: 650px; margin: auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.12); }}
            h1 {{ color: #f58220; margin-bottom: 5px; }}
            .job-box {{ background: #f7f7f7; padding: 15px; border-radius: 8px; margin: 20px 0; }}
            label {{ font-weight: bold; display: block; margin-top: 15px; }}
            input[type="text"] {{ width: 100%; padding: 12px; font-size: 16px; box-sizing: border-box; }}
            canvas {{ width: 100%; height: 260px; border: 2px solid #333; border-radius: 8px; background: white; margin-top: 10px; touch-action: none; -ms-touch-action: none; user-select: none; -webkit-user-select: none; -webkit-touch-callout: none; display: block; }}
            button {{ width: 100%; padding: 14px; margin-top: 15px; font-size: 16px; border: none; border-radius: 8px; cursor: pointer; }}
            .submit {{ background: #f58220; color: white; font-weight: bold; }}
            .clear {{ background: #555; color: white; }}
            .small {{ font-size: 13px; color: #555; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <img src="/logo.png" alt="CDR M&E Services Ltd" style="display:block; max-width:320px; width:80%; margin:0 auto 20px auto;">
            <h2 style="text-align:center;">Client Signature</h2>
            <div class="job-box">
                <p><strong>CDR Number:</strong> {cdr_number}</p>
                <p><strong>Site:</strong> {site}</p>
                <p><strong>Address:</strong> {address}</p>
                <p><strong>Task:</strong> {task}</p>
            </div>
            <form method="post" action="/submit-signature" onsubmit="return submitForm()">
                <input type="hidden" name="cdr_number" value="{cdr_number}">
                <input type="hidden" name="token" value="{token}">
                <input type="hidden" name="signature_data" id="signature_data">

                <label>Client name</label>
                <input type="text" name="client_name" required placeholder="Enter client name">

                <label>
                    <input type="checkbox" required>
                    I confirm the works/attendance have been completed as described.
                </label>

                <label>Signature</label>
                <canvas id="signature-pad"></canvas>

                <button type="button" class="clear" onclick="clearSignature()">Clear Signature</button>
                <button type="submit" class="submit">Submit Signature</button>
            </form>
            <p class="small">This digital signature will be stored against the job record for audit and completion evidence.</p>
        </div>
        <script>
            const canvas = document.getElementById("signature-pad");
            const signaturePad = new SignaturePad(canvas, {{
                minWidth: 1,
                maxWidth: 2.5,
                throttle: 0,
                velocityFilterWeight: 0.7
            }});

            // Stop mobile browsers treating signature movement as page scroll/swipe.
            // The passive:false option is important on iPhone/Android.
            ["touchstart", "touchmove", "touchend", "pointerdown", "pointermove", "pointerup"].forEach(function(eventName) {{
                canvas.addEventListener(eventName, function(event) {{
                    event.preventDefault();
                    event.stopPropagation();
                }}, {{ passive: false }});
            }});

            let savedSignature = null;

            function resizeCanvas() {{
                const ratio = Math.max(window.devicePixelRatio || 1, 1);
                const rect = canvas.getBoundingClientRect();
                if (!signaturePad.isEmpty()) {{
                    savedSignature = signaturePad.toDataURL("image/png");
                }}

                canvas.width = rect.width * ratio;
                canvas.height = rect.height * ratio;
                canvas.getContext("2d").scale(ratio, ratio);
                signaturePad.clear();

                if (savedSignature) {{
                    signaturePad.fromDataURL(savedSignature);
                    savedSignature = null;
                }}
            }}

            window.addEventListener("orientationchange", function() {{
                setTimeout(resizeCanvas, 250);
            }});
            window.addEventListener("resize", resizeCanvas);
            resizeCanvas();

            function clearSignature() {{
                signaturePad.clear();
            }}

            function submitForm() {{
                if (signaturePad.isEmpty()) {{
                    alert("Please provide a signature.");
                    return false;
                }}
                document.getElementById("signature_data").value = signaturePad.toDataURL("image/png");
                return true;
            }}
        </script>
    </body>
    </html>
    """

    return HTMLResponse(html)


@web_app.post("/submit-signature", response_class=HTMLResponse)
def submit_signature(
    cdr_number: str = Form(...),
    token: str = Form(...),
    client_name: str = Form(...),
    signature_data: str = Form(...),
):
    site_id, jobs_list_id, job = get_job_by_cdr_and_token(cdr_number, token)

    if not job:
        return HTMLResponse("Invalid or expired signature link.", status_code=404)

    if bool_field(job["fields"].get("ClientSignatureReceived")):
        return HTMLResponse("This job has already been signed.", status_code=200)

    signature_link = upload_signature_to_sharepoint(site_id, cdr_number, signature_data)

    update_list_item_fields(
        site_id,
        jobs_list_id,
        job["id"],
        {
            "ClientSignatureReceived": True,
            "ClientSignatureName": client_name,
            "ClientSignatureDateTime": graph_datetime_now(),
            "ClientSignatureLink": signature_link,
        },
    )

    return HTMLResponse("""
<!DOCTYPE html>
<html>
<head>
    <title>Signature Saved</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">

    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f4f6f8;
            margin: 0;
            padding: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            text-align: center;
        }

        .container {
            background: white;
            width: 100%;
            max-width: 520px;
            border-radius: 24px;
            padding: 40px 25px;
            box-shadow: 0 6px 25px rgba(0,0,0,0.12);
        }

        .logo {
            width: 180px;
            max-width: 80%;
            margin-bottom: 30px;
        }

        .tick {
            font-size: 90px;
            color: #22c55e;
            margin-bottom: 20px;
            font-weight: bold;
        }

        h1 {
            font-size: 40px;
            margin: 0 0 25px 0;
            color: #111827;
        }

        p {
            font-size: 24px;
            line-height: 1.6;
            color: #374151;
            margin: 0;
        }
    </style>
</head>
<body>

<div class="container">

    <img src="/logo.png" class="logo">

    <div class="tick">✓</div>

    <h1>Signature Saved</h1>

    <p>
        Thank you.<br><br>
        The worksheet has been successfully submitted.
        <br><br>
        You may now close this page.
    </p>

</div>

</body>
</html>
""")


def run_signature_web_server():
    uvicorn.run(web_app, host="0.0.0.0", port=PORT, log_level="info")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if role.lower() == "inactive":
        await update.message.reply_text(
            "You are not currently authorised to use the CDR Engineer Bot. Please ask the office to check your Engineers list record, Telegram ID, Role and Active status.",
            reply_markup=get_main_menu(role),
        )
        return

    if user_can_use_helpdesk(role):
        message = f"CDR Engineer Bot is online. Your access level is {role}."
    else:
        message = "CDR Engineer Bot is online. Use the menu below."

    await update.message.reply_text(
        message,
        reply_markup=get_main_menu(role),
    )


async def startday_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, _, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        day_logs_list_id = get_list_id(site_id, DAY_LOGS_LIST)
        day_logs = get_list_items(site_id, day_logs_list_id)
        active_day = find_active_day_log(day_logs, user_id)

        if active_day:
            await update.message.reply_text(
                "Your day is already active. You can now use 📋 My Jobs.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        context.user_data["start_day"] = {
            "site_id": site_id,
            "day_logs_list_id": day_logs_list_id,
            "engineer_name": current_engineer["name"],
            "engineer_lookup_id": current_engineer["lookup_id"],
            "engineer_telegram_id": user_id,
            "work_date": get_today_iso(),
            "van_check_answers": [],
            "van_photo_links": [],
            "day_logs": day_logs,
        }

        await update.message.reply_text(
            "Are you sure you want to start your day?",
            reply_markup=get_yes_no_keyboard("startday_confirm"),
        )
        return START_DAY_CONFIRM

    except Exception as e:
        print(f"ERROR starting day: {e}")
        await update.message.reply_text(
            "There was an error starting your day. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def startday_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, START_DAY_CONFIRM)
    if menu_result is not None:
        return menu_result

    answer = update.message.text.strip().lower()

    if answer not in ["yes", "no", "y", "n"]:
        await update.message.reply_text("Please tap Yes or No.")
        return START_DAY_CONFIRM

    if answer in ["no", "n"]:
        context.user_data.pop("start_day", None)
        await update.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    await update.message.reply_text("Starting your day. Please enter the van registration.")
    return START_DAY_VAN_REG


async def startday_confirm_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    answer = query.data.split("|", 1)[1]

    if answer == "no":
        context.user_data.pop("start_day", None)
        await query.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    await query.message.reply_text("Starting your day. Please enter the van registration.")
    return START_DAY_VAN_REG


async def startday_van_reg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, START_DAY_VAN_REG)
    if menu_result is not None:
        return menu_result

    start_day = context.user_data.get("start_day")

    if not start_day:
        await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    van_reg = update.message.text.strip().upper()

    if not van_reg:
        await update.message.reply_text("Please enter the van registration.")
        return START_DAY_VAN_REG

    start_day["van_reg"] = van_reg

    await update.message.reply_text(
        "Please enter your start mileage as a number.\n\n"
        "Example: 15234 or 0."
    )

    return START_DAY_START_MILEAGE



def parse_sharepoint_date_to_date(value):
    if not value:
        return None

    try:
        if isinstance(value, datetime):
            return value.astimezone(UK_TZ).date() if value.tzinfo else value.date()

        value = str(value).strip()

        if not value:
            return None

        if value.endswith("Z"):
            value = value.replace("Z", "+00:00")

        return datetime.fromisoformat(value).astimezone(UK_TZ).date()
    except Exception:
        try:
            return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
        except Exception:
            return None


def engineer_needs_van_check(day_logs, engineer_telegram_id, today_date=None):
    today_date = today_date or datetime.now(UK_TZ).date()
    engineer_telegram_id = str(engineer_telegram_id)

    last_check_date = None

    for item in day_logs:
        fields = item.get("fields", {})

        log_telegram_id = str(
            fields.get("EngineerTelegramID")
            or fields.get("Engineer Telegram ID")
            or ""
        )

        if log_telegram_id != engineer_telegram_id:
            continue

        van_completed = bool_field(
            fields.get("VanCheckCompleted")
            or fields.get("Van Check Completed")
        )

        if not van_completed:
            continue

        log_date = parse_sharepoint_date_to_date(
            fields.get("WorkDate")
            or fields.get("Work Date")
            or fields.get("Created")
        )

        if log_date and (last_check_date is None or log_date > last_check_date):
            last_check_date = log_date

    if last_check_date is None:
        return True, None, None

    days_since = (today_date - last_check_date).days
    return days_since >= VAN_CHECK_INTERVAL_DAYS, last_check_date, days_since


def build_start_day_log_fields(start_day, van_check_completed=False):
    return build_field_payload_for_list(
        start_day["site_id"],
        start_day["day_logs_list_id"],
        {
            "Title": f"{start_day['engineer_name']} - {start_day['work_date']}",
            "Engineer": start_day["engineer_name"],
            "Engineer Name": start_day["engineer_name"],
            "EngineerName": start_day["engineer_name"],
            "Engineer Telegram ID": start_day["engineer_telegram_id"],
            "EngineerTelegramID": start_day["engineer_telegram_id"],
            "Engineer Lookup ID": start_day["engineer_lookup_id"],
            "EngineerLookupID": start_day["engineer_lookup_id"],
            "Work Date": start_day["work_date"],
            "WorkDate": start_day["work_date"],
            "Start Time": graph_datetime_now(),
            "StartTime": graph_datetime_now(),
            "Start Mileage": start_day.get("start_mileage", "0"),
            "StartMileage": start_day.get("start_mileage", "0"),
            "Van Registration": start_day.get("van_reg", ""),
            "VanRegistration": start_day.get("van_reg", ""),
            "Van Check Completed": bool(van_check_completed),
            "VanCheckCompleted": bool(van_check_completed),
            "Van Check Answers": "\n\n".join(start_day.get("van_check_answers", [])),
            "VanCheckAnswers": "\n\n".join(start_day.get("van_check_answers", [])),
            "Van Photo Links": "\n".join(start_day.get("van_photo_links", [])),
            "VanPhotoLinks": "\n".join(start_day.get("van_photo_links", [])),
            "Status": DAY_ACTIVE_STATUS,
        },
    )


def create_start_day_log(start_day, van_check_completed=False):
    day_log_fields = build_start_day_log_fields(start_day, van_check_completed)
    create_list_item_fields(
        start_day["site_id"],
        start_day["day_logs_list_id"],
        day_log_fields,
    )



async def startday_start_mileage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, START_DAY_START_MILEAGE)
    if menu_result is not None:
        return menu_result

    start_day = context.user_data.get("start_day")

    if not start_day:
        await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    mileage = normalise_mileage(update.message.text)

    if mileage is None:
        await update.message.reply_text(
            "Please enter start mileage as numbers only. Example: 15234 or 0."
        )
        return START_DAY_START_MILEAGE

    start_day["start_mileage"] = mileage

    needs_check, last_check_date, days_since = engineer_needs_van_check(
        start_day.get("day_logs", []),
        start_day["engineer_telegram_id"],
    )

    if not needs_check:
        create_start_day_log(start_day, van_check_completed=False)
        context.user_data.pop("start_day", None)

        await update.message.reply_text(
            f"Start mileage recorded: {mileage}\n\n"
            f"Van check not due today. Last completed van check was {days_since} day(s) ago "
            f"on {last_check_date}. Your jobs are now unlocked.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    start_day["question_index"] = 0

    await update.message.reply_text(
        f"Start mileage recorded: {mileage}\n\n"
        f"Van check is due. It is required every {VAN_CHECK_INTERVAL_DAYS} days.\n\n"
        f"Van check 1 of {len(VAN_CHECK_QUESTIONS)}:\n"
        f"{VAN_CHECK_QUESTIONS[0]}"
    )

    return START_DAY_VAN_CHECK


async def startday_van_check(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, START_DAY_VAN_CHECK)
    if menu_result is not None:
        return menu_result

    start_day = context.user_data.get("start_day")

    if not start_day:
        await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    answer = update.message.text.strip()
    question_index = start_day.get("question_index", 0)
    question = VAN_CHECK_QUESTIONS[question_index]

    if not answer:
        await update.message.reply_text("Please enter an answer.")
        return START_DAY_VAN_CHECK

    start_day["van_check_answers"].append(f"{question}\nAnswer: {answer}")
    question_index += 1
    start_day["question_index"] = question_index

    if question_index < len(VAN_CHECK_QUESTIONS):
        await update.message.reply_text(
            f"Van check {question_index + 1} of {len(VAN_CHECK_QUESTIONS)}:\n"
            f"{VAN_CHECK_QUESTIONS[question_index]}"
        )
        return START_DAY_VAN_CHECK

    await update.message.reply_text(
        "Van check questions complete. Please upload these 3 van photos:\n\n"
        "1. Van cab\n"
        "2. Van rear load area\n"
        "3. Dashboard / mileage\n\n"
        "Send the 3 photos now, then type DONE when finished."
    )
    return START_DAY_VAN_PHOTOS


async def startday_van_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, START_DAY_VAN_PHOTOS)
    if menu_result is not None:
        return menu_result

    try:
        start_day = context.user_data.get("start_day")

        if not start_day:
            await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu(await get_role_for_update(update)))
            return ConversationHandler.END

        if update.message.text and update.message.text.strip().upper() == "DONE":
            photo_count = len(start_day.get("van_photo_links", []))

            if photo_count < 3:
                await update.message.reply_text(
                    f"I have {photo_count} van photo(s). Please upload all 3 required photos before typing DONE:\n\n"
                    "1. Van cab\n"
                    "2. Van rear load area\n"
                    "3. Dashboard / mileage"
                )
                return START_DAY_VAN_PHOTOS
            create_start_day_log(start_day, van_check_completed=True)

            context.user_data.pop("start_day", None)

            await update.message.reply_text(
                "Van check completed and day started. Your jobs are now unlocked. Tap 📋 My Jobs to view today's work.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        if update.message.photo:
            photo = update.message.photo[-1]
            telegram_file = await context.bot.get_file(photo.file_id)
            file_bytes = await telegram_file.download_as_bytearray()

            timestamp = datetime.now(UK_TZ).strftime("%Y%m%d_%H%M%S")
            file_name = f"{safe_folder_name(start_day.get('van_reg', 'VAN'))}_{timestamp}_{photo.file_unique_id}.jpg"

            photo_link = upload_van_check_photo_to_sharepoint(
                start_day["site_id"],
                start_day["work_date"],
                start_day.get("van_reg", "VAN"),
                file_name,
                bytes(file_bytes),
            )

            start_day["van_photo_links"].append(photo_link)
            return START_DAY_VAN_PHOTOS

        await update.message.reply_text("Please send the required van photos, or type DONE once all 3 have been uploaded.")
        return START_DAY_VAN_PHOTOS

    except Exception as e:
        print(f"ERROR saving van check/start day: {e}")
        await update.message.reply_text(
            "There was an error saving the van check. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def startday_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("start_day", None)
    await update.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def endday_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, _, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        day_logs_list_id, active_day = get_active_day_for_engineer(site_id, user_id)

        if not active_day:
            await update.message.reply_text(
                "You do not have an active day to end. Tap 🟢 Start Day when you begin work.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        _, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        open_jobs = get_open_jobs_for_engineer_today(jobs_data, current_engineer["lookup_id"])

        if open_jobs:
            await update.message.reply_text(
                "You cannot end your day while you still have job(s) assigned for today. "
                "Complete them, mark No Access, or mark Revisit Required first.\n\n"
                f"Open job(s):\n{format_open_jobs_for_end_day(open_jobs)}",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        context.user_data["end_day"] = {
            "site_id": site_id,
            "day_logs_list_id": day_logs_list_id,
            "day_log_item_id": active_day["id"],
            "day_log_fields": active_day.get("fields", {}),
            "engineer_name": current_engineer["name"],
            "engineer_lookup_id": current_engineer["lookup_id"],
        }

        await update.message.reply_text(
            "Are you sure you want to end your day?",
            reply_markup=get_yes_no_keyboard("endday_confirm"),
        )
        return END_DAY_CONFIRM

    except Exception as e:
        print(f"ERROR ending day: {e}")
        await update.message.reply_text(
            "There was an error ending your day. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def endday_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, END_DAY_CONFIRM)
    if menu_result is not None:
        return menu_result

    answer = update.message.text.strip().lower()

    if answer not in ["yes", "no", "y", "n"]:
        await update.message.reply_text("Please tap Yes or No.")
        return END_DAY_CONFIRM

    if answer in ["no", "n"]:
        context.user_data.pop("end_day", None)
        await update.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    await update.message.reply_text(
        "Please enter your end mileage as a number.\n\n"
        "If you do not need to record mileage, type 0."
    )
    return END_DAY_MILEAGE


async def endday_confirm_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    answer = query.data.split("|", 1)[1]

    if answer == "no":
        context.user_data.pop("end_day", None)
        await query.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    await query.message.reply_text(
        "Please enter your end mileage as a number.\n\n"
        "If you do not need to record mileage, type 0."
    )
    return END_DAY_MILEAGE


async def endday_mileage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, END_DAY_MILEAGE)
    if menu_result is not None:
        return menu_result

    try:
        mileage = normalise_mileage(update.message.text)

        if mileage is None:
            await update.message.reply_text(
                "Please enter mileage as numbers only. Example: 15234 or 0."
            )
            return END_DAY_MILEAGE

        end_day = context.user_data.get("end_day")

        if not end_day:
            await update.message.reply_text("Please try /endday again.", reply_markup=get_main_menu(await get_role_for_update(update)))
            return ConversationHandler.END

        # Re-check assigned jobs before closing the day in case one was added while the engineer was in the end-day flow.
        site_id, _, _, engineers, jobs_data = get_sharepoint_data()
        open_jobs = get_open_jobs_for_engineer_today(jobs_data, end_day["engineer_lookup_id"])

        if open_jobs:
            context.user_data.pop("end_day", None)
            await update.message.reply_text(
                "Your day has not been ended because you still have job(s) assigned for today. "
                "Complete them, mark No Access, or mark Revisit Required first.\n\n"
                f"Open job(s):\n{format_open_jobs_for_end_day(open_jobs)}",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        end_time = datetime.now(UK_TZ)
        start_time_value = get_field_value(
            end_day.get("day_log_fields", {}),
            "StartTime",
            "Start Time",
        )
        start_time = parse_sharepoint_datetime(start_time_value)
        hours = calculate_day_pay_hours(start_time, end_time)
        pay_summary = build_pay_summary(start_time, end_time, hours)

        start_mileage_value = get_field_value(
            end_day.get("day_log_fields", {}),
            "StartMileage",
            "Start Mileage",
        )

        total_mileage = None

        try:
            total_mileage = round(float(mileage) - float(start_mileage_value or 0), 2)
        except Exception:
            total_mileage = None

        update_payload = {
            "End Time": end_time.isoformat(),
            "EndTime": end_time.isoformat(),
            "End Mileage": mileage,
            "EndMileage": mileage,
            "Total Mileage": total_mileage if total_mileage is not None else "",
            "TotalMileage": total_mileage if total_mileage is not None else "",
            "Status": DAY_CLOSED_STATUS,
            "Pay Summary": pay_summary,
            "PaySummary": pay_summary,
        }

        if hours:
            update_payload.update({
                "Total Hours": hours["total_hours"],
                "TotalHours": hours["total_hours"],
                "Normal Hours": hours["normal_hours"],
                "NormalHours": hours["normal_hours"],
                "OOH Hours": hours["ooh_hours"],
                "OOHHours": hours["ooh_hours"],
                "Break Deducted": hours["break_deducted"],
                "BreakDeducted": hours["break_deducted"],
            })

        end_day_fields = build_field_payload_for_list(
            end_day["site_id"],
            end_day["day_logs_list_id"],
            update_payload,
        )

        update_list_item_fields(
            end_day["site_id"],
            end_day["day_logs_list_id"],
            end_day["day_log_item_id"],
            end_day_fields,
        )

        context.user_data.pop("end_day", None)

        await update.message.reply_text(
            "Day ended. Your job buttons are now locked until you start your next day.\n\n"
            f"{pay_summary}",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR saving end day: {e}")
        await update.message.reply_text(
            "There was an error saving your end day record. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def endday_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("end_day", None)
    await update.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def mystatus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, _, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return

        _, active_day = get_active_day_for_engineer(site_id, user_id)

        if active_day:
            fields = active_day["fields"]
            await update.message.reply_text(
                f"Status: Day active\n"
                f"Engineer: {current_engineer['name']}\n"
                f"Start time: {format_sharepoint_date(fields.get('StartTime', ''))} {str(fields.get('StartTime', ''))[11:16] if fields.get('StartTime') else ''}",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
        else:
            await update.message.reply_text(
                "Status: No active day. Tap 🟢 Start Day before using job buttons.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )

    except Exception as e:
        print(f"ERROR getting status: {e}")
        await update.message.reply_text("There was an error checking your status.", reply_markup=get_main_menu(await get_role_for_update(update)))


async def menu_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == MENU_START_DAY:
        return await startday_start(update, context)

    if text == MENU_MY_JOBS:
        await jobs(update, context)
        return

    if text == MENU_END_DAY:
        return await endday_start(update, context)

    if text == MENU_BUG_IDEA:
        return await bugidea_start(update, context)

    if text == MENU_UPLOAD_RECEIPTS:
        return await receipt_start(update, context)

    if text in [MENU_HELPDESK, MENU_LOG_JOB, MENU_REASSIGN_JOB, MENU_OPEN_JOBS, MENU_FIND_JOB, MENU_CANCEL_JOB, MENU_DELETE_JOB, MENU_ENGINEER_MENU]:
        return await helpdesk_menu_button(update, context)


async def helpdesk_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to use the helpdesk menu.",
            reply_markup=get_main_menu(role),
        )
        return

    await update.message.reply_text(
        "Helpdesk menu opened. Choose an option below.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
    )


async def helpdesk_menu_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    text = update.message.text

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to use this function.",
            reply_markup=get_main_menu(role),
        )
        return

    if text == MENU_ENGINEER_MENU:
        if role.lower() != "admin":
            await update.message.reply_text(
                "Only Admin users can switch to the Engineer menu.",
                reply_markup=get_helpdesk_menu(include_engineer_menu=False),
            )
            return

        await update.message.reply_text(
            "Engineer menu opened.",
            reply_markup=get_engineer_menu(include_helpdesk_menu=True),
        )
        return

    if text == MENU_LOG_JOB:
        return await logjob_start(update, context)

    if text == MENU_REASSIGN_JOB:
        return await reassign_start(update, context)

    if text == MENU_FIND_JOB:
        return await findjob_start(update, context)

    if text == MENU_OPEN_JOBS:
        return await openjobs_start(update, context)

    if text == MENU_CANCEL_JOB:
        return await canceljob_start(update, context)

    if text == MENU_DELETE_JOB:
        return await deletejob_start(update, context)

    if text == MENU_UPLOAD_RECEIPTS:
        return await receipt_start(update, context)

    coming_next = {
        MENU_LOG_JOB: "Log Job",
        MENU_REASSIGN_JOB: "Reassign Job",
        MENU_OPEN_JOBS: "Open Jobs",
        MENU_FIND_JOB: "Find Job",
        MENU_CANCEL_JOB: "Cancel Job",
        MENU_DELETE_JOB: "Delete Job",
        MENU_HELPDESK: "Helpdesk",
    }

    if text == MENU_HELPDESK:
        await update.message.reply_text(
            "Helpdesk menu opened. Choose an option below.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return

    await update.message.reply_text(
        f"{coming_next.get(text, 'This helpdesk option')} is permission-protected and ready for the next build. "
        "Next step is wiring this button into the SharePoint job create/reassign flow.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
    )




async def logjob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to log jobs.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    try:
        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)
        assignable_engineers = get_active_assignable_engineers(engineers)

        if not assignable_engineers:
            await update.message.reply_text(
                "No active assignable engineers were found. Check the Engineers list has TelegramID, Role and Active set correctly.",
                reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
            )
            return ConversationHandler.END

        context.user_data["log_job"] = {
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "assignable_engineers": assignable_engineers,
            "role": role,
        }

        await update.message.reply_text(
            "Log new job.\n\nEnter the CDR/job number.\n\nExample: CDR012896",
            reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
        )
        return LOGJOB_CDR_NUMBER

    except Exception as e:
        print(f"ERROR starting log job flow: {e}")
        await update.message.reply_text(
            "There was an error opening Log Job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def logjob_cdr_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    if not job:
        await update.message.reply_text("Please start again using ➕ Log Job.")
        return ConversationHandler.END

    value = update.message.text.strip().upper()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter a CDR/job number. Example: CDR012896")
        return LOGJOB_CDR_NUMBER

    job["cdr_number"] = value
    await update.message.reply_text("Customer name?\n\nExample: Cleveland Fire Brigade")
    return LOGJOB_CUSTOMER_NAME


async def logjob_customer_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the customer name. This appears on the worksheet.")
        return LOGJOB_CUSTOMER_NAME
    job["customer_name"] = value
    await update.message.reply_text("Customer address?\n\nThis is the billing/client address for the worksheet. You can use multiple lines.")
    return LOGJOB_CUSTOMER_ADDRESS


async def logjob_customer_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the customer address. This appears on the worksheet.")
        return LOGJOB_CUSTOMER_ADDRESS
    job["customer_address"] = value
    await update.message.reply_text("Site name?\n\nExample: Headquarters")
    return LOGJOB_SITE_NAME


async def logjob_site_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the site name.")
        return LOGJOB_SITE_NAME

    job["site_name"] = value
    job.pop("site_candidate", None)
    job.pop("site_notes", None)

    candidate = find_best_site_candidate(job["site_id"], job["jobs_list_id"], value)
    if candidate:
        job["site_candidate"] = candidate
        await update.message.reply_text(format_site_candidate(candidate))
        return LOGJOB_SITE_CONFIRM

    await update.message.reply_text("Site address?\n\nThis is the address sent to the engineer and used for Maps.")
    return LOGJOB_SITE_ADDRESS


async def logjob_site_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    if not job:
        await update.message.reply_text("Please start again using ➕ Log Job.")
        return ConversationHandler.END

    answer = update.message.text.strip().lower()
    candidate = job.get("site_candidate") or {}

    if answer in ["yes", "y", "use", "use it", "correct"]:
        job["site_name"] = candidate.get("site_name") or job.get("site_name", "")
        job["site_address"] = candidate.get("address", "")
        if candidate.get("notes"):
            job["site_notes"] = candidate.get("notes", "")
        await update.message.reply_text(prompt_for_site_notes(candidate.get("notes", "")))
        return LOGJOB_SITE_NOTES

    if answer in ["edit", "change", "amend", "wrong address"]:
        await update.message.reply_text(
            "Enter the correct site address.\n\n"
            "This will be used for the job and saved back to the Sites list."
        )
        return LOGJOB_SITE_ADDRESS

    if answer in ["no", "n", "manual", "enter manually"]:
        job.pop("site_candidate", None)
        job.pop("site_notes", None)
        await update.message.reply_text("Site address?\n\nThis is the address sent to the engineer and used for Maps.")
        return LOGJOB_SITE_ADDRESS

    await update.message.reply_text("Reply YES to use it, EDIT to change the address, or NO to enter manually.")
    return LOGJOB_SITE_CONFIRM


async def logjob_site_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the site address.")
        return LOGJOB_SITE_ADDRESS
    job["site_address"] = value

    existing_notes = (job.get("site_candidate") or {}).get("notes", "")
    await update.message.reply_text(prompt_for_site_notes(existing_notes))
    return LOGJOB_SITE_NOTES


async def logjob_site_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()

    if not is_blank_or_skip(value):
        job["site_notes"] = value
    elif not job.get("site_notes") and (job.get("site_candidate") or {}).get("notes"):
        job["site_notes"] = (job.get("site_candidate") or {}).get("notes", "")

    await update.message.reply_text("Contact name/number?\n\nType N/A if there is no contact.")
    return LOGJOB_CONTACT


async def logjob_contact(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    job["contact"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text("Task / job description?\n\nExample: HQ leaking tea boiler in ELT kitchen")
    return LOGJOB_TASK


async def logjob_task(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the task/job description.")
        return LOGJOB_TASK
    job["task"] = value
    await update.message.reply_text("Any extra engineer notes?\n\nType N/A if none.")
    return LOGJOB_NOTES


async def logjob_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    job["notes"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text("Job date?\n\nUse today, tomorrow, or DD/MM/YYYY.")
    return LOGJOB_DATE


async def logjob_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    parsed = parse_helpdesk_job_date(update.message.text)
    if not parsed:
        await update.message.reply_text("Please enter a valid date, for example today, tomorrow, or 13/05/2026.")
        return LOGJOB_DATE
    job["date"] = parsed
    try:
        job["date_display"] = datetime.fromisoformat(parsed).strftime("%d/%m/%Y")
    except Exception:
        job["date_display"] = parsed
    await update.message.reply_text("Start/time required?\n\nUse HH:MM, 0800, 13:30, now, or asap.")
    return LOGJOB_TIME


async def logjob_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    parsed = normalise_helpdesk_time(update.message.text)
    if not parsed:
        await update.message.reply_text("Please enter a valid time, for example 08:00, 0800, 13:30, now, or asap.")
        return LOGJOB_TIME
    job["time"] = parsed
    await update.message.reply_text(
        "Job category? Reply with a number:\n\n" +
        "\n".join(f"{i}. {choice}" for i, choice in enumerate(JOB_CATEGORY_CHOICES, start=1))
    )
    return LOGJOB_CATEGORY


async def logjob_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    text = update.message.text.strip()

    selected = None
    if text.isdigit():
        index = int(text)
        if 1 <= index <= len(JOB_CATEGORY_CHOICES):
            selected = JOB_CATEGORY_CHOICES[index - 1]
    else:
        for choice in JOB_CATEGORY_CHOICES:
            if text.lower() == choice.lower():
                selected = choice
                break

    if not selected:
        await update.message.reply_text(
            "Please choose a valid category number:\n\n" +
            "\n".join(f"{i}. {choice}" for i, choice in enumerate(JOB_CATEGORY_CHOICES, start=1))
        )
        return LOGJOB_CATEGORY

    job["category"] = selected
    await update.message.reply_text("Customer order number?\n\nType N/A if not available.")
    return LOGJOB_ORDER_NUMBER


async def logjob_order_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    job["order_number"] = "" if is_blank_or_skip(value) else value
    engineers = job.get("assignable_engineers", [])
    await update.message.reply_text(
        "Assign engineer(s). Reply with the number, or multiple numbers separated by commas.\n\n" +
        format_engineer_selection_list(engineers)
    )
    return LOGJOB_ASSIGN_ENGINEERS


async def logjob_assign_engineers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    engineers = job.get("assignable_engineers", [])
    selected, error = parse_engineer_selection(update.message.text, engineers)
    if error:
        await update.message.reply_text(error + "\n\n" + format_engineer_selection_list(engineers))
        return LOGJOB_ASSIGN_ENGINEERS

    job["assigned_engineers"] = selected
    await update.message.reply_text(build_log_job_review(job))
    return LOGJOB_REVIEW


async def logjob_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    job = context.user_data.get("log_job")
    role = job.get("role", "Helpdesk") if job else "Helpdesk"
    answer = update.message.text.strip().lower()

    if answer in ["no", "n", "cancel"]:
        context.user_data.pop("log_job", None)
        await update.message.reply_text(
            "Job logging cancelled. Nothing has been created.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END

    if answer in ["restart", "redo"]:
        context.user_data.pop("log_job", None)
        return await logjob_start(update, context)

    if answer not in ["yes", "y"]:
        await update.message.reply_text("Reply YES to create and send, NO to cancel, or RESTART to start again.")
        return LOGJOB_REVIEW

    try:
        site_id = job["site_id"]
        jobs_list_id = job["jobs_list_id"]

        initial_fields = build_helpdesk_job_fields(site_id, jobs_list_id, job, telegram_notified=False)
        created_item = create_list_item_fields(site_id, jobs_list_id, initial_fields)
        item_id = created_item.get("id")
        created_fields = created_item.get("fields", initial_fields)

        # Make sure the CDR folder structure exists early, before engineers upload photos or signatures.
        try:
            drive_id = get_drive_id(site_id, PHOTO_LIBRARY)
            cdr_folder = safe_folder_name(job["cdr_number"])
            ensure_folder(drive_id, f"{PHOTO_BASE_FOLDER}/{cdr_folder}")
            ensure_folder(drive_id, f"{WORKSHEET_BASE_FOLDER}/{cdr_folder}")
            ensure_folder(drive_id, f"{SIGNATURE_BASE_FOLDER}/{cdr_folder}")
        except Exception as folder_error:
            print(f"WARNING: Could not pre-create job folders for {job['cdr_number']}: {folder_error}")

        sent_to_any, failed = await send_created_job_to_engineers(
            context.bot,
            item_id,
            created_fields,
            job.get("assigned_engineers", []),
        )

        final_update = build_field_payload_for_list(
            site_id,
            jobs_list_id,
            {
                "TelegramNotified": bool(sent_to_any),
                "Telegram Notified": bool(sent_to_any),
                "Status": ASSIGNED_STATUS if sent_to_any else AWAITING_DEPLOYMENT_STATUS,
                "EngineerVisitLog": f"{now_log_time()} - Helpdesk - Job logged via Telegram and {'sent to engineer(s)' if sent_to_any else 'created but not sent'}",
                "Engineer Visit Log": f"{now_log_time()} - Helpdesk - Job logged via Telegram and {'sent to engineer(s)' if sent_to_any else 'created but not sent'}",
            },
        )
        update_list_item_fields(site_id, jobs_list_id, item_id, final_update)

        upsert_site_record_from_job(job)

        context.user_data.pop("log_job", None)

        message = (
            f"Job created in SharePoint: {job['cdr_number']}\n"
            f"Assigned to: {', '.join(e['name'] for e in job.get('assigned_engineers', []))}\n"
            f"Telegram sent: {'Yes' if sent_to_any else 'No'}"
        )
        if failed:
            message += "\n\nSend issues:\n" + "\n".join(failed[:5])

        await update.message.reply_text(
            message,
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR creating logged job: {e}")
        await update.message.reply_text(
            "There was an error creating the job. Nothing further has been sent. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def logjob_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    context.user_data.pop("log_job", None)
    await update.message.reply_text(
        "Log Job cancelled. Nothing has been created.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END


async def reassign_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        role = await get_role_for_update(update)
        if not user_can_use_helpdesk(role):
            await update.message.reply_text("You do not have permission to reassign jobs.", reply_markup=get_main_menu(role))
            return ConversationHandler.END

        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)
        assignable_engineers = get_active_assignable_engineers(engineers)

        if not assignable_engineers:
            await update.message.reply_text(
                "No assignable engineers found. Check the Engineers list has TelegramID, Role = Engineer/Admin, and Active = Yes.",
                reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
            )
            return ConversationHandler.END

        context.user_data["reassign_job"] = {
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "engineers": engineers,
            "assignable_engineers": assignable_engineers,
            "role": role,
        }

        await update.message.reply_text("Enter the CDR number to reassign.\n\nExample: CDR01012")
        return REASSIGN_CDR_NUMBER

    except Exception as e:
        print(f"ERROR opening Reassign Job: {e}")
        role = await get_role_for_update(update)
        await update.message.reply_text(
            "There was an error opening Reassign Job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
        )
        return ConversationHandler.END


async def reassign_cdr_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("reassign_job")
    if not data:
        await update.message.reply_text("Please start again using 🔁 Reassign Job.")
        return ConversationHandler.END

    cdr_number = update.message.text.strip()
    if is_blank_or_skip(cdr_number):
        await update.message.reply_text("Please enter the CDR number.")
        return REASSIGN_CDR_NUMBER

    try:
        jobs_data = get_list_items(data["site_id"], data["jobs_list_id"])
        job = find_job_by_cdr(jobs_data, cdr_number)
        if not job:
            await update.message.reply_text("I could not find that CDR number. Please check it and try again.")
            return REASSIGN_CDR_NUMBER

        fields = job.get("fields", {})
        current_engineers = get_current_assigned_engineers_from_job(fields, data.get("engineers", []))

        data.update({
            "cdr_number": fields.get("CDRNumber", cdr_number),
            "item_id": job.get("id"),
            "job_fields": fields,
            "current_engineers": current_engineers,
        })

        if current_engineers:
            await update.message.reply_text(
                "Current assigned engineer(s):\n\n" +
                format_engineer_selection_list(current_engineers) +
                "\n\nWho do you want to remove? Reply with number(s), ALL, or NONE."
            )
        else:
            data["remove_engineers"] = []
            await update.message.reply_text(
                "There are currently no engineers assigned.\n\nAssign engineer(s). Reply with number(s):\n\n" +
                format_engineer_selection_list(data.get("assignable_engineers", []))
            )
            return REASSIGN_ASSIGN_ENGINEERS

        return REASSIGN_REMOVE_ENGINEERS

    except Exception as e:
        print(f"ERROR finding job for reassignment: {e}")
        await update.message.reply_text("There was an error finding the job. Please check Railway logs.")
        return ConversationHandler.END


async def reassign_remove_engineers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("reassign_job")
    if not data:
        await update.message.reply_text("Please start again using 🔁 Reassign Job.")
        return ConversationHandler.END

    selected, error = parse_remove_engineer_selection(update.message.text, data.get("current_engineers", []))
    if error:
        await update.message.reply_text(
            error + "\n\nCurrent assigned engineer(s):\n\n" + format_engineer_selection_list(data.get("current_engineers", []))
        )
        return REASSIGN_REMOVE_ENGINEERS

    data["remove_engineers"] = selected
    await update.message.reply_text(
        "Assign new engineer(s). Reply with number(s), or type NONE if you only want to remove engineers and leave the job awaiting dispatch.\n\n" +
        format_engineer_selection_list(data.get("assignable_engineers", []))
    )
    return REASSIGN_ASSIGN_ENGINEERS


async def reassign_assign_engineers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("reassign_job")
    if not data:
        await update.message.reply_text("Please start again using 🔁 Reassign Job.")
        return ConversationHandler.END

    text = update.message.text.strip().lower()
    if text in ["none", "no", "skip", "0"]:
        selected = []
    else:
        selected, error = parse_engineer_selection(update.message.text, data.get("assignable_engineers", []))
        if error:
            await update.message.reply_text(error + "\n\n" + format_engineer_selection_list(data.get("assignable_engineers", [])))
            return REASSIGN_ASSIGN_ENGINEERS

    data["assign_engineers"] = selected

    remove_ids = {e["lookup_id"] for e in data.get("remove_engineers", [])}
    final = []
    seen = set()

    for engineer in data.get("current_engineers", []):
        if engineer["lookup_id"] not in remove_ids and engineer["lookup_id"] not in seen:
            final.append(engineer)
            seen.add(engineer["lookup_id"])

    for engineer in selected:
        if engineer["lookup_id"] not in seen:
            final.append(engineer)
            seen.add(engineer["lookup_id"])

    data["final_engineers"] = final

    await update.message.reply_text("Reason for reassignment?\n\nExample: Engineer delayed on another job / change of plan / emergency priority.")
    return REASSIGN_REASON


async def reassign_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("reassign_job")
    if not data:
        await update.message.reply_text("Please start again using 🔁 Reassign Job.")
        return ConversationHandler.END

    value = update.message.text.strip()
    data["reason"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text(build_reassign_review(data))
    return REASSIGN_REVIEW


async def reassign_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("reassign_job")
    role = data.get("role", "Helpdesk") if data else "Helpdesk"
    answer = update.message.text.strip().lower()

    if answer in ["no", "n", "cancel"]:
        context.user_data.pop("reassign_job", None)
        await update.message.reply_text(
            "Reassignment cancelled. Nothing has been changed.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END

    if answer in ["restart", "redo"]:
        context.user_data.pop("reassign_job", None)
        return await reassign_start(update, context)

    if answer not in ["yes", "y"]:
        await update.message.reply_text("Reply YES to reassign and send, NO to cancel, or RESTART to start again.")
        return REASSIGN_REVIEW

    try:
        site_id = data["site_id"]
        jobs_list_id = data["jobs_list_id"]
        item_id = data["item_id"]
        fields = data.get("job_fields", {})
        cdr_number = fields.get("CDRNumber", data.get("cdr_number", ""))
        final_engineers = data.get("final_engineers", [])
        new_engineers = data.get("assign_engineers", [])

        reassignment_note = (
            f"Removed: {', '.join(e['name'] for e in data.get('remove_engineers', [])) or 'None'}; "
            f"Assigned: {', '.join(e['name'] for e in new_engineers) or 'None'}; "
            f"Final: {', '.join(e['name'] for e in final_engineers) or 'None'}"
        )
        if data.get("reason"):
            reassignment_note += f"; Reason: {data['reason']}"

        updated_log = append_engineer_log(fields, "Helpdesk", "Reassigned", reassignment_note)

        # First place the job back into dispatch while the assignment is changed.
        interim_payload = build_field_payload_for_list(
            site_id,
            jobs_list_id,
            {
                "Status": AWAITING_DEPLOYMENT_STATUS,
                "TelegramNotified": False,
                "Telegram Notified": False,
                "EngineerVisitLog": updated_log,
                "Engineer Visit Log": updated_log,
            },
        )
        interim_payload["EngineerLookupId@odata.type"] = "Collection(Edm.Int32)"
        interim_payload["EngineerLookupId"] = [int(e["lookup_id"]) for e in final_engineers]
        update_list_item_fields(site_id, jobs_list_id, item_id, interim_payload)

        # Send the job only to newly selected engineer(s), not engineers who were already left assigned.
        send_fields = dict(fields)
        send_fields.update({
            "Status": ASSIGNED_STATUS if final_engineers else AWAITING_DEPLOYMENT_STATUS,
            "EngineerVisitLog": updated_log,
        })
        sent_to_any, failed = await send_created_job_to_engineers(
            context.bot,
            item_id,
            send_fields,
            new_engineers,
        )

        final_payload = build_field_payload_for_list(
            site_id,
            jobs_list_id,
            {
                "Status": ASSIGNED_STATUS if final_engineers else AWAITING_DEPLOYMENT_STATUS,
                "TelegramNotified": bool(sent_to_any or (final_engineers and not new_engineers)),
                "Telegram Notified": bool(sent_to_any or (final_engineers and not new_engineers)),
            },
        )
        update_list_item_fields(site_id, jobs_list_id, item_id, final_payload)

        context.user_data.pop("reassign_job", None)

        message = (
            f"Job reassigned: {cdr_number}\n"
            f"Final assigned engineer(s): {', '.join(e['name'] for e in final_engineers) or 'None'}\n"
            f"Telegram sent to new engineer(s): {'Yes' if sent_to_any else 'No'}"
        )
        if failed:
            message += "\n\nSend issues:\n" + "\n".join(failed[:5])

        await update.message.reply_text(
            message,
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR reassigning job: {e}")
        await update.message.reply_text(
            "There was an error reassigning the job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def reassign_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    context.user_data.pop("reassign_job", None)
    await update.message.reply_text(
        "Reassign Job cancelled. Nothing has been changed.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END


def job_matches_search(fields, search_text):
    needle = normalise_cdr(search_text)
    loose_needle = str(search_text or "").strip().lower()

    if not needle and not loose_needle:
        return False

    searchable_values = [
        fields.get("CDRNumber", ""),
        fields.get("Title", ""),
        fields.get("CustomerName", ""),
        fields.get("Customer Name", ""),
        fields.get("SiteName", ""),
        fields.get("Site Name", ""),
        fields.get("Address", ""),
        fields.get("Task", ""),
        fields.get("Notes", ""),
        fields.get("ContactName", ""),
        fields.get("Contact Name", ""),
        fields.get("CustomerOrderNumber", ""),
        fields.get("Customer Order Number", ""),
        fields.get("JobCategory", ""),
        fields.get("Job Category", ""),
        fields.get("Status", ""),
        fields.get("JobOutcome", ""),
    ]

    for value in searchable_values:
        text = str(value or "").lower()
        if loose_needle and loose_needle in text:
            return True
        if needle and needle in normalise_cdr(value):
            return True

    return False


def search_jobs_for_helpdesk(jobs_data, search_text, limit=10):
    matches = []

    # Exact CDR/Title matches first.
    exact = find_job_by_cdr(jobs_data, search_text)
    if exact:
        matches.append(exact)

    seen_ids = {str(item.get("id")) for item in matches}

    for item in jobs_data:
        if str(item.get("id")) in seen_ids:
            continue
        fields = item.get("fields", {})
        if job_matches_search(fields, search_text):
            matches.append(item)
            seen_ids.add(str(item.get("id")))
        if len(matches) >= limit:
            break

    return matches


def format_job_search_results(matches):
    lines = [f"Found {len(matches)} matching job(s):"]

    for index, item in enumerate(matches, start=1):
        fields = item.get("fields", {})
        cdr = get_field_value(fields, "CDRNumber", "CDR Number", "Title") or "No CDR"
        site = get_field_value(fields, "SiteName", "Site Name") or "No site"
        date = format_sharepoint_date(get_field_value(fields, "Date") or "") or "No date"
        status = get_field_value(fields, "Status") or "No status"
        outcome = get_field_value(fields, "JobOutcome", "Job Outcome") or ""
        outcome_text = f" | {outcome}" if outcome else ""
        lines.append(f"{index}. {cdr} | {site} | {date} | {status}{outcome_text}")

    lines.append("\nReply with the number to view full details, or type SEARCH to search again.")
    return "\n".join(lines)


def format_helpdesk_job_detail(job, engineers):
    fields = job.get("fields", {})
    assigned = get_current_assigned_engineers_from_job(fields, engineers)
    assigned_names = ", ".join(e["name"] for e in assigned) or "None"
    visit_log = str(get_field_value(fields, "EngineerVisitLog", "Engineer Visit Log") or "").strip()
    visit_lines = [line for line in visit_log.splitlines() if line.strip()]
    if visit_lines:
        visit_text = "\n".join(visit_lines[-8:])
    else:
        visit_text = "No visit log yet."

    worksheet_link = get_field_value(fields, "WorksheetLink", "Worksheet Link", "WorksheetPDF", "Worksheet PDF") or ""
    signature_received = "Yes" if bool_field(get_field_value(fields, "ClientSignatureReceived", "Client Signature Received")) else "No"

    return (
        "Job details:\n\n"
        f"CDR Number: {get_field_value(fields, 'CDRNumber', 'CDR Number', 'Title') or ''}\n"
        f"Customer: {get_field_value(fields, 'CustomerName', 'Customer Name') or ''}\n"
        f"Customer Address: {get_field_value(fields, 'CustomerAddress', 'Customer Address') or ''}\n"
        f"Site: {get_field_value(fields, 'SiteName', 'Site Name') or ''}\n"
        f"Address: {get_field_value(fields, 'Address') or ''}\n"
        f"Date: {format_sharepoint_date(get_field_value(fields, 'Date') or '')}\n"
        f"Time: {get_field_value(fields, 'StartTime', 'Start Time') or ''}\n"
        f"Status: {get_field_value(fields, 'Status') or ''}\n"
        f"Outcome: {get_field_value(fields, 'JobOutcome', 'Job Outcome') or ''}\n"
        f"Assigned engineer(s): {assigned_names}\n"
        f"Contact: {get_field_value(fields, 'ContactName', 'Contact Name') or ''}\n"
        f"Task: {get_field_value(fields, 'Task') or ''}\n"
        f"Notes: {get_field_value(fields, 'Notes') or ''}\n"
        f"Order Number: {get_field_value(fields, 'CustomerOrderNumber', 'Customer Order Number') or ''}\n"
        f"Category: {get_field_value(fields, 'JobCategory', 'Job Category') or ''}\n"
        f"Telegram Notified: {'Yes' if is_notified(fields) else 'No'}\n"
        f"Client Signature Received: {signature_received}\n"
        f"Worksheet Link: {worksheet_link or 'None'}\n\n"
        "Recent visit log:\n"
        f"{visit_text}\n\n"
        "Next actions:\n"
        "- Use 🔁 Reassign Job to change/send engineers.\n"
        "- Use 🔎 Find Job again to search another job."
    )


def open_job_bucket(fields):
    status = str(get_field_value(fields, "Status") or "").strip()
    outcome = str(get_field_value(fields, "JobOutcome", "Job Outcome") or "").strip()
    job_date = sharepoint_date_to_uk_date(get_field_value(fields, "Date") or "")
    today = datetime.now(UK_TZ).date()

    if status == COMPLETED_STATUS or outcome == "Completed":
        return "completed_today" if job_date == today else "completed_old"

    if status in ["No Access", "Revisit Required"] or outcome in ["No Access", "Revisit Required"]:
        if status in [AWAITING_DEPLOYMENT_STATUS, LEGACY_AWAITING_DEPLOYMENT_STATUS, ""]:
            return "returned"
        if status not in [ASSIGNED_STATUS, TRAVELLING_STATUS, ON_SITE_STATUS]:
            return "returned"

    if status in [AWAITING_DEPLOYMENT_STATUS, LEGACY_AWAITING_DEPLOYMENT_STATUS, ""]:
        return "awaiting_dispatch"

    if status == ASSIGNED_STATUS:
        return "assigned_not_started"

    if status in [TRAVELLING_STATUS, ON_SITE_STATUS]:
        return "live"

    return "other_open"


def is_helpdesk_open_job(fields):
    return open_job_bucket(fields) not in ["completed_old"]


def filter_helpdesk_open_jobs(jobs_data, filter_key="all_open"):
    jobs = []
    for job in jobs_data:
        fields = job.get("fields", {})
        bucket = open_job_bucket(fields)

        if filter_key == "all_open":
            if bucket not in ["completed_today", "completed_old"]:
                jobs.append(job)
        elif filter_key == "today":
            job_date = sharepoint_date_to_uk_date(get_field_value(fields, "Date") or "")
            if job_date == datetime.now(UK_TZ).date() and bucket != "completed_old":
                jobs.append(job)
        elif filter_key == bucket:
            jobs.append(job)

    def sort_key(job):
        fields = job.get("fields", {})
        dt = sharepoint_date_to_uk_date(get_field_value(fields, "Date") or "")
        cdr = str(get_field_value(fields, "CDRNumber", "CDR Number", "Title") or "")
        # undated jobs first so helpdesk notices them
        return (dt or datetime.min.date(), cdr)

    jobs.sort(key=sort_key)
    return jobs


def openjobs_filter_keyboard_text():
    return (
        "Open Jobs dashboard.\n\n"
        "Reply with one option:\n"
        "1. Awaiting Dispatch\n"
        "2. Assigned - not started\n"
        "3. Live - Travelling / On Site\n"
        "4. Returned - No Access / Revisit Required\n"
        "5. Completed today\n"
        "6. All open jobs\n"
        "7. Today's jobs\n\n"
        "You can also type a status name, e.g. Assigned, On Site, No Access."
    )


def parse_openjobs_filter(text):
    value = str(text or "").strip().lower()
    mapping = {
        "1": "awaiting_dispatch",
        "awaiting": "awaiting_dispatch",
        "awaiting dispatch": "awaiting_dispatch",
        "awaiting deployment": "awaiting_dispatch",
        "2": "assigned_not_started",
        "assigned": "assigned_not_started",
        "not started": "assigned_not_started",
        "3": "live",
        "live": "live",
        "travelling": "live",
        "on site": "live",
        "onsite": "live",
        "4": "returned",
        "returned": "returned",
        "no access": "returned",
        "revisit": "returned",
        "revisit required": "returned",
        "5": "completed_today",
        "completed": "completed_today",
        "completed today": "completed_today",
        "6": "all_open",
        "all": "all_open",
        "all open": "all_open",
        "open": "all_open",
        "7": "today",
        "today": "today",
        "todays": "today",
        "today's": "today",
    }
    return mapping.get(value)


def openjobs_filter_label(filter_key):
    return {
        "awaiting_dispatch": "Awaiting Dispatch",
        "assigned_not_started": "Assigned - not started",
        "live": "Live - Travelling / On Site",
        "returned": "Returned - No Access / Revisit Required",
        "completed_today": "Completed today",
        "all_open": "All open jobs",
        "today": "Today's jobs",
    }.get(filter_key, "Open jobs")


def format_openjobs_results(jobs, filter_key, engineers, max_rows=20):
    label = openjobs_filter_label(filter_key)
    lines = [f"{label}: {len(jobs)} job(s)"]

    if not jobs:
        lines.append("\nNo jobs found in this view.")
        lines.append("\nType FILTER to choose another view, or /cancel to exit.")
        return "\n".join(lines)

    for index, job in enumerate(jobs[:max_rows], start=1):
        fields = job.get("fields", {})
        cdr = get_field_value(fields, "CDRNumber", "CDR Number", "Title") or "No CDR"
        site = get_field_value(fields, "SiteName", "Site Name") or "No site"
        date = format_sharepoint_date(get_field_value(fields, "Date") or "") or "No date"
        status = get_field_value(fields, "Status") or "No status"
        outcome = get_field_value(fields, "JobOutcome", "Job Outcome") or ""
        assigned = get_current_assigned_engineers_from_job(fields, engineers)
        assigned_names = ", ".join(e["name"] for e in assigned) or "Unassigned"
        outcome_text = f" / {outcome}" if outcome and outcome != status else ""
        lines.append(f"{index}. {cdr} | {site} | {date} | {status}{outcome_text} | {assigned_names}")

    if len(jobs) > max_rows:
        lines.append(f"...and {len(jobs) - max_rows} more. Narrow the view if needed.")

    lines.append("\nReply with a number to view full details, FILTER to change view, or /cancel to exit.")
    return "\n".join(lines)


async def openjobs_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to view Open Jobs.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    try:
        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)

        context.user_data["open_jobs"] = {
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "engineers": engineers,
            "role": role,
        }

        await update.message.reply_text(
            openjobs_filter_keyboard_text(),
            reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
        )
        return OPENJOBS_FILTER

    except Exception as e:
        print(f"ERROR opening Open Jobs: {e}")
        await update.message.reply_text(
            "There was an error opening Open Jobs. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def openjobs_filter(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("open_jobs")
    if not data:
        await update.message.reply_text("Please start again using 📋 Open Jobs.")
        return ConversationHandler.END

    filter_key = parse_openjobs_filter(update.message.text)
    if not filter_key:
        await update.message.reply_text("I did not recognise that view.\n\n" + openjobs_filter_keyboard_text())
        return OPENJOBS_FILTER

    try:
        jobs_data = get_list_items(data["site_id"], data["jobs_list_id"])
        jobs = filter_helpdesk_open_jobs(jobs_data, filter_key)
        data["filter_key"] = filter_key
        data["jobs"] = jobs

        await update.message.reply_text(format_openjobs_results(jobs, filter_key, data.get("engineers", [])))
        return OPENJOBS_SELECT

    except Exception as e:
        print(f"ERROR filtering Open Jobs: {e}")
        await update.message.reply_text("There was an error loading Open Jobs. Please check Railway logs.")
        return ConversationHandler.END


async def openjobs_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("open_jobs")
    if not data:
        await update.message.reply_text("Please start again using 📋 Open Jobs.")
        return ConversationHandler.END

    text = update.message.text.strip().lower()
    if text in ["filter", "filters", "back", "change", "restart"]:
        await update.message.reply_text(openjobs_filter_keyboard_text())
        return OPENJOBS_FILTER

    if not text.isdigit():
        await update.message.reply_text("Reply with a job number, FILTER to change view, or /cancel to exit.")
        return OPENJOBS_SELECT

    jobs = data.get("jobs", [])
    index = int(text)
    if index < 1 or index > len(jobs):
        await update.message.reply_text("That number is not in the list. Reply with one of the job numbers shown.")
        return OPENJOBS_SELECT

    selected_job = jobs[index - 1]
    role = data.get("role", "Helpdesk")
    await update.message.reply_text(
        format_helpdesk_job_detail(selected_job, data.get("engineers", [])),
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
    )
    context.user_data.pop("open_jobs", None)
    return ConversationHandler.END


async def openjobs_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    context.user_data.pop("open_jobs", None)
    await update.message.reply_text(
        "Open Jobs cancelled.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END





async def canceljob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to cancel jobs.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    context.user_data["cancel_job"] = {"role": role}
    await update.message.reply_text(
        "Enter the CDR/job number you want to cancel.\n\n"
        "This will keep the SharePoint record but mark it as Cancelled.",
        reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
    )
    return CANCELJOB_CDR_NUMBER


async def canceljob_cdr_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("cancel_job") or {}
    cdr_number = update.message.text.strip()

    if is_blank_or_skip(cdr_number):
        await update.message.reply_text("Please enter a CDR/job number to cancel.")
        return CANCELJOB_CDR_NUMBER

    try:
        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)
        jobs_data = get_list_items(site_id, jobs_list_id)
        job = find_job_by_cdr(jobs_data, cdr_number)

        if not job:
            await update.message.reply_text("I could not find that job. Check the CDR number and try again.")
            return CANCELJOB_CDR_NUMBER

        data.update({
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "engineers": engineers,
            "job": job,
            "cdr_number": get_field_value(job.get("fields", {}), "CDRNumber", "CDR Number", "Title") or cdr_number,
        })
        context.user_data["cancel_job"] = data

        await update.message.reply_text(
            format_helpdesk_job_detail(job, engineers)
            + "\n\nCancel this job?\n\nReply YES to cancel, or NO to stop."
        )
        return CANCELJOB_CONFIRM

    except Exception as e:
        print(f"ERROR starting cancel job: {e}")
        await update.message.reply_text(
            "There was an error finding that job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(str(data.get("role", "")).lower() == "admin")),
        )
        return ConversationHandler.END


async def canceljob_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("cancel_job") or {}
    role = data.get("role") or await get_role_for_update(update)
    answer = update.message.text.strip().lower()

    if answer in ["no", "n", "cancel", "stop"]:
        context.user_data.pop("cancel_job", None)
        await update.message.reply_text(
            "Cancel Job stopped. Nothing has been changed.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(str(role).lower() == "admin")),
        )
        return ConversationHandler.END

    if answer not in ["yes", "y"]:
        await update.message.reply_text("Reply YES to cancel this job, or NO to stop.")
        return CANCELJOB_CONFIRM

    try:
        site_id = data["site_id"]
        jobs_list_id = data["jobs_list_id"]
        job = data["job"]
        fields = job.get("fields", {})
        cdr_number = data.get("cdr_number") or get_field_value(fields, "CDRNumber", "CDR Number", "Title") or ""
        updated_log = append_engineer_log(fields, "Helpdesk", "Cancelled", f"Cancelled by {update.effective_user.id}")

        payload = build_field_payload_for_list(
            site_id,
            jobs_list_id,
            {
                "Status": "Cancelled",
                "JobOutcome": "Cancelled by Helpdesk",
                "Job Outcome": "Cancelled by Helpdesk",
                "TelegramNotified": False,
                "Telegram Notified": False,
                "EngineerVisitLog": updated_log,
                "Engineer Visit Log": updated_log,
            },
        )
        payload.update(clear_engineer_assignment_payload())
        update_list_item_fields(site_id, jobs_list_id, job["id"], payload)

        context.user_data.pop("cancel_job", None)
        await update.message.reply_text(
            f"Job cancelled: {cdr_number}\n\nThe SharePoint record has been kept for audit history.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(str(role).lower() == "admin")),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR cancelling job: {e}")
        await update.message.reply_text(
            "There was an error cancelling the job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(str(role).lower() == "admin")),
        )
        return ConversationHandler.END


async def canceljob_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    context.user_data.pop("cancel_job", None)
    await update.message.reply_text(
        "Cancel Job stopped. Nothing has been changed.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(str(role).lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END


async def deletejob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if str(role).lower() != "admin":
        await update.message.reply_text(
            "Hard delete is admin-only. Use ❌ Cancel Job if you need to remove a job from the active workflow.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    context.user_data["delete_job"] = {"role": role}
    await update.message.reply_text(
        "ADMIN HARD DELETE.\n\n"
        "Enter the CDR/job number you want to permanently delete from the SharePoint list.\n\n"
        "This does not delete any separate files/photos already uploaded to document libraries.",
        reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
    )
    return DELETEJOB_CDR_NUMBER


async def deletejob_cdr_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("delete_job") or {}
    cdr_number = update.message.text.strip()

    if is_blank_or_skip(cdr_number):
        await update.message.reply_text("Please enter a CDR/job number to hard delete.")
        return DELETEJOB_CDR_NUMBER

    try:
        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)
        jobs_data = get_list_items(site_id, jobs_list_id)
        job = find_job_by_cdr(jobs_data, cdr_number)

        if not job:
            await update.message.reply_text("I could not find that job. Check the CDR number and try again.")
            return DELETEJOB_CDR_NUMBER

        fields = job.get("fields", {})
        actual_cdr = get_field_value(fields, "CDRNumber", "CDR Number", "Title") or cdr_number
        data.update({
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "engineers": engineers,
            "job": job,
            "cdr_number": actual_cdr,
        })
        context.user_data["delete_job"] = data

        await update.message.reply_text(
            format_helpdesk_job_detail(job, engineers)
            + "\n\n⚠️ ADMIN HARD DELETE WARNING ⚠️\n"
            + "This permanently deletes the SharePoint list item.\n\n"
            + f"To confirm, type exactly:\nDELETE {actual_cdr}\n\n"
            + "Type NO to stop."
        )
        return DELETEJOB_CONFIRM

    except Exception as e:
        print(f"ERROR starting hard delete: {e}")
        await update.message.reply_text(
            "There was an error finding that job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=True),
        )
        return ConversationHandler.END


async def deletejob_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("delete_job") or {}
    answer = update.message.text.strip()
    cdr_number = str(data.get("cdr_number", "")).strip()

    if answer.lower() in ["no", "n", "cancel", "stop"]:
        context.user_data.pop("delete_job", None)
        await update.message.reply_text(
            "Hard Delete stopped. Nothing has been changed.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=True),
        )
        return ConversationHandler.END

    expected = f"DELETE {cdr_number}"
    if answer != expected:
        await update.message.reply_text(
            f"Confirmation did not match. To permanently delete this job, type exactly:\n{expected}\n\nOr type NO to stop."
        )
        return DELETEJOB_CONFIRM

    try:
        delete_list_item(data["site_id"], data["jobs_list_id"], data["job"]["id"])
        context.user_data.pop("delete_job", None)
        await update.message.reply_text(
            f"Hard deleted SharePoint job item: {cdr_number}\n\nAny files/photos already uploaded to document libraries have not been deleted.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=True),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR hard deleting job: {e}")
        await update.message.reply_text(
            "There was an error hard deleting the job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=True),
        )
        return ConversationHandler.END


async def deletejob_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("delete_job", None)
    await update.message.reply_text(
        "Hard Delete stopped. Nothing has been changed.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=True),
    )
    return ConversationHandler.END


async def findjob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to find jobs.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    try:
        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)

        context.user_data["find_job"] = {
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "engineers": engineers,
            "role": role,
        }

        await update.message.reply_text(
            "Find job.\n\nEnter a CDR number, site name, customer, address, order number, status, or keyword.\n\nExample: CDR012896 or Headquarters",
            reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
        )
        return FINDJOB_SEARCH

    except Exception as e:
        print(f"ERROR opening Find Job: {e}")
        await update.message.reply_text(
            "There was an error opening Find Job. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def findjob_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("find_job")
    if not data:
        await update.message.reply_text("Please start again using 🔎 Find Job.")
        return ConversationHandler.END

    text = update.message.text.strip()
    if is_blank_or_skip(text):
        await update.message.reply_text("Please enter a CDR number, site/customer name, address or keyword.")
        return FINDJOB_SEARCH

    try:
        jobs_data = get_list_items(data["site_id"], data["jobs_list_id"])
        matches = search_jobs_for_helpdesk(jobs_data, text, limit=10)

        if not matches:
            await update.message.reply_text(
                "No jobs found for that search. Try a CDR number, site name, customer name, address, order number or keyword."
            )
            return FINDJOB_SEARCH

        data["matches"] = matches
        data["last_search"] = text

        if len(matches) == 1:
            await update.message.reply_text(
                format_helpdesk_job_detail(matches[0], data.get("engineers", [])),
                reply_markup=get_helpdesk_menu(include_engineer_menu=(data.get("role", "").lower() == "admin")),
            )
            context.user_data.pop("find_job", None)
            return ConversationHandler.END

        await update.message.reply_text(format_job_search_results(matches))
        return FINDJOB_SELECT

    except Exception as e:
        print(f"ERROR searching jobs: {e}")
        await update.message.reply_text("There was an error searching jobs. Please check Railway logs.")
        return ConversationHandler.END


async def findjob_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("find_job")
    if not data:
        await update.message.reply_text("Please start again using 🔎 Find Job.")
        return ConversationHandler.END

    text = update.message.text.strip().lower()
    if text in ["search", "new", "again", "restart"]:
        await update.message.reply_text("Enter a new CDR number, site/customer name, address, order number or keyword.")
        return FINDJOB_SEARCH

    if not text.isdigit():
        await update.message.reply_text("Reply with a job number from the list, or type SEARCH to search again.")
        return FINDJOB_SELECT

    index = int(text)
    matches = data.get("matches", [])
    if index < 1 or index > len(matches):
        await update.message.reply_text("That number is not in the list. Reply with one of the job numbers shown.")
        return FINDJOB_SELECT

    selected_job = matches[index - 1]
    role = data.get("role", "Helpdesk")
    await update.message.reply_text(
        format_helpdesk_job_detail(selected_job, data.get("engineers", [])),
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
    )
    context.user_data.pop("find_job", None)
    return ConversationHandler.END


async def findjob_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    context.user_data.pop("find_job", None)
    await update.message.reply_text(
        "Find Job cancelled.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END

HELPDESK_MENU_TEXTS = {
    MENU_HELPDESK,
    MENU_LOG_JOB,
    MENU_REASSIGN_JOB,
    MENU_OPEN_JOBS,
    MENU_FIND_JOB,
    MENU_UPLOAD_RECEIPTS,
    MENU_ENGINEER_MENU,
}

ALL_MENU_TEXTS = {
    MENU_START_DAY,
    MENU_MY_JOBS,
    MENU_END_DAY,
    MENU_BUG_IDEA,
    MENU_UPLOAD_RECEIPTS,
    *HELPDESK_MENU_TEXTS,
}


def is_bot_menu_text(value):
    return str(value or "").strip() in ALL_MENU_TEXTS


async def handle_menu_during_conversation(update: Update, context: ContextTypes.DEFAULT_TYPE, current_state):
    """
    Prevent reply-keyboard menu buttons being saved as answers inside active flows.
    This is especially important during worksheets, where helpdesk/admin buttons
    must never end up in Work Completed, Materials, Notes or PDF comments.
    """
    text = update.message.text.strip() if update.message and update.message.text else ""

    if text in ALL_MENU_TEXTS:
        await update.message.reply_text(
            "You are part-way through another task. Finish it or type /cancel first. "
            "This menu button has not been added to the worksheet or job record.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return current_state

    return None


async def bugidea_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, engineers, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        context.user_data["bug_idea"] = {
            "site_id": site_id,
            "engineer_name": current_engineer["name"],
            "engineer_telegram_id": user_id,
        }

        await update.message.reply_text(
            "Please type the bug, issue, improvement idea or request you want to log.\n\n"
            "Example: The photo upload message is unclear on job completion.\n\n"
            "Type /cancel to cancel."
        )

        return BUG_IDEA_TEXT

    except Exception as e:
        print(f"ERROR starting bug/idea log: {e}")
        await update.message.reply_text(
            "There was an error opening the bug/idea log. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def bugidea_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, BUG_IDEA_TEXT)
    if menu_result is not None:
        return menu_result

    try:
        bug_idea = context.user_data.get("bug_idea")

        if not bug_idea:
            await update.message.reply_text("Please try again from the menu.", reply_markup=get_main_menu(await get_role_for_update(update)))
            return ConversationHandler.END

        text_value = update.message.text.strip()

        if not text_value:
            await update.message.reply_text("Please type the bug or idea you want to log.")
            return BUG_IDEA_TEXT

        bug_ideas_list_id = get_list_id(bug_idea["site_id"], BUG_IDEAS_LIST)

        title = f"{bug_idea['engineer_name']} - {datetime.now(UK_TZ).strftime('%d/%m/%Y %H:%M')}"

        fields_to_create = build_field_payload_for_list(
            bug_idea["site_id"],
            bug_ideas_list_id,
            {
                "Title": title,
                "EngineerName": bug_idea["engineer_name"],
                "Engineer Name": bug_idea["engineer_name"],
                "EngineerTelegramID": bug_idea["engineer_telegram_id"],
                "Engineer Telegram ID": bug_idea["engineer_telegram_id"],
                "BugIdeaText": text_value,
                "Bug Idea Text": text_value,
                "DateSubmitted": graph_datetime_now(),
                "Date Submitted": graph_datetime_now(),
                "Status": "New",
            },
        )

        create_list_item_fields(
            bug_idea["site_id"],
            bug_ideas_list_id,
            fields_to_create,
        )

        context.user_data.pop("bug_idea", None)

        await update.message.reply_text(
            "Logged. Thanks — this has been sent to the office as a bug/idea.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )

        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR saving bug/idea: {e}")
        await update.message.reply_text(
            "There was an error saving the bug/idea. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def bugidea_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("bug_idea", None)
    await update.message.reply_text("Bug/idea cancelled.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def receipt_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        site_id = get_site_id()
        user_id = str(update.effective_user.id)
        engineer_name = ""

        try:
            engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
            engineers = get_list_items(site_id, engineers_list_id)
            engineers_by_telegram, _ = build_engineer_maps(engineers)
            current_engineer = engineers_by_telegram.get(user_id)
            if current_engineer:
                engineer_name = current_engineer.get("name", "")
        except Exception as e:
            print(f"WARNING: Could not auto-detect receipt engineer: {e}")

        if not engineer_name:
            await update.message.reply_text(
                "I could not match your Telegram account to an engineer record, so I cannot upload receipts under your name. Please ask the office to check your Engineers list record has your Telegram ID and EngineerName.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        context.user_data["receipt_upload"] = {
            "site_id": site_id,
            "engineer_name": engineer_name,
            "receipt_date": "",
            "receipt_links": [],
            "receipt_count": 0,
        }

        await update.message.reply_text(
            f"""Upload receipts started.

Engineer: {engineer_name}

Enter the receipt date.

Use DD/MM/YYYY, YYYY-MM-DD, or TODAY.

Type /cancel to cancel."""
        )
        return RECEIPT_DATE

    except Exception as e:
        print(f"ERROR starting receipt upload: {e}")
        await update.message.reply_text(
            "There was an error starting receipt upload. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def receipt_engineer_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Retired: receipt uploads are now locked to the logged-in Telegram user's
    # EngineerName from SharePoint so nobody can submit receipts under another name.
    return await receipt_date(update, context)


async def receipt_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, RECEIPT_DATE)
    if menu_result is not None:
        return menu_result

    data = context.user_data.get("receipt_upload")
    if not data:
        await update.message.reply_text("Please start again from 🧾 Upload Receipts.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    parsed_date = parse_helpdesk_job_date(update.message.text)
    if not parsed_date:
        await update.message.reply_text("Please enter a valid date, for example 12/05/2026 or TODAY.")
        return RECEIPT_DATE

    data["receipt_date"] = parsed_date

    await update.message.reply_text(
        f"Receipt upload ready.\n\nEngineer: {data['engineer_name']}\nDate: {parsed_date}\n\n"
        "Now send all receipt photos/PDFs/files.\n\nType DONE when finished."
    )
    return RECEIPT_UPLOADS


async def receipt_uploads(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, RECEIPT_UPLOADS)
    if menu_result is not None:
        return menu_result

    data = context.user_data.get("receipt_upload")
    if not data:
        await update.message.reply_text("Please start again from 🧾 Upload Receipts.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    if update.message.text and update.message.text.strip().upper() == "DONE":
        count = len(data.get("receipt_links", []))
        if count == 0:
            await update.message.reply_text("No receipts uploaded yet. Send at least one receipt, or type /cancel to cancel.")
            return RECEIPT_UPLOADS

        context.user_data.pop("receipt_upload", None)
        await update.message.reply_text(
            f"Receipts uploaded for finance.\n\nEngineer: {data['engineer_name']}\nDate: {data['receipt_date']}\nFiles uploaded: {count}",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    try:
        file_bytes = None
        extension = "jpg"
        unique_id = "receipt"

        if update.message.photo:
            photo = update.message.photo[-1]
            telegram_file = await context.bot.get_file(photo.file_id)
            file_bytes = await telegram_file.download_as_bytearray()
            unique_id = photo.file_unique_id
            extension = "jpg"

        elif update.message.document:
            document = update.message.document
            telegram_file = await context.bot.get_file(document.file_id)
            file_bytes = await telegram_file.download_as_bytearray()
            unique_id = document.file_unique_id
            original_name = clean_receipt_file_name(document.file_name or "receipt")
            if "." in original_name:
                extension = original_name.rsplit(".", 1)[1].lower()
            else:
                extension = "bin"

        else:
            await update.message.reply_text("Please send a receipt photo/file, or type DONE when finished.")
            return RECEIPT_UPLOADS

        data["receipt_count"] = int(data.get("receipt_count", 0)) + 1
        timestamp = datetime.now(UK_TZ).strftime("%Y%m%d_%H%M%S")
        engineer_part = safe_folder_name(data.get("engineer_name", "ENGINEER"))
        file_name = f"{data['receipt_date']}_{engineer_part}_{timestamp}_{data['receipt_count']}_{unique_id}.{extension}"

        receipt_link = upload_receipt_to_sharepoint(
            data["site_id"],
            data["receipt_date"],
            data["engineer_name"],
            file_name,
            bytes(file_bytes),
        )
        data["receipt_links"].append(receipt_link)

        # Intentionally no per-file confirmation. Engineers can upload all receipts
        # and type DONE once finished to avoid Telegram message spam.
        return RECEIPT_UPLOADS

    except Exception as e:
        print(f"ERROR uploading receipt: {e}")
        await update.message.reply_text(
            "There was an error uploading that receipt. Please try again or ask the office to check Railway logs."
        )
        return RECEIPT_UPLOADS


async def receipt_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("receipt_upload", None)
    await update.message.reply_text("Receipt upload cancelled.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Your Telegram ID is: {update.effective_user.id}")


async def jobs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        today = datetime.now(UK_TZ).date()

        site_id, _, _, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)

        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return

        if not engineer_has_active_day(site_id, user_id):
            await update.message.reply_text(
                "Please start your day first using 🟢 Start Day or /startday. Your jobs are locked until your day has started.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return

        found_any = False

        for job in jobs_data:
            fields = job["fields"]
            item_id = job["id"]

            job_date = sharepoint_date_to_uk_date(fields.get("Date", ""))
            assigned_ids = get_assigned_engineer_ids(fields)

            if current_engineer["lookup_id"] in assigned_ids and job_date == today and not is_closed_job(fields):
                found_any = True
                await update.message.reply_text(
                    "Today's job:\n\n" + format_job(fields, current_engineer["name"]),
                    reply_markup=get_job_buttons(item_id, fields.get("Address", "")),
                )

        if not found_any:
            await update.message.reply_text("No jobs assigned today.")

    except Exception as e:
        print(f"ERROR in /jobs: {e}")
        await update.message.reply_text(
            "There was an error getting your jobs. Please ask the office to check Railway logs."
        )


async def get_engineer_job_for_callback(query, require_active_day=True):
    site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
    engineers_by_telegram, _ = build_engineer_maps(engineers)

    user_id = str(query.from_user.id)
    current_engineer = engineers_by_telegram.get(user_id)

    if not current_engineer:
        await query.message.reply_text("You are not set up as an engineer.")
        return None

    if require_active_day and not engineer_has_active_day(site_id, user_id):
        await query.message.reply_text(
            "Please start your day first using 🟢 Start Day or /startday. Job buttons are locked until your day has started."
        )
        return None

    return site_id, jobs_list_id, jobs_data, current_engineer, user_id


async def abort_job_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    try:
        parts = query.data.split("|")
        item_id = parts[1]

        lookup = await get_engineer_job_for_callback(query, require_active_day=True)
        if not lookup:
            return ConversationHandler.END

        site_id, jobs_list_id, jobs_data, current_engineer, user_id = lookup
        job = find_job_by_item_id(jobs_data, item_id)

        if not job:
            await query.message.reply_text("Could not find this job.")
            return ConversationHandler.END

        fields = job.get("fields", {})

        if is_closed_job(fields):
            await query.message.reply_text(
                "This job has already been closed or returned to the office. No further action is required."
            )
            return ConversationHandler.END

        if current_engineer["lookup_id"] not in get_assigned_engineer_ids(fields):
            await query.message.reply_text("You are not assigned to this job.")
            return ConversationHandler.END

        context.user_data["abort_job"] = {
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "item_id": item_id,
            "fields": fields,
            "engineer_name": current_engineer["name"],
            "engineer_lookup_id": current_engineer["lookup_id"],
            "user_id": user_id,
        }

        await query.message.reply_text(
            "Why are you aborting attendance on this job?\n\n"
            "This is for cases where you did not attend / cannot attend, and the job needs sending back for reassignment.",
            reply_markup=get_abort_reason_keyboard(item_id),
        )
        return ABORTJOB_REASON

    except Exception as e:
        print(f"ERROR starting abort job: {e}")
        await query.message.reply_text("There was an error starting the abort job flow. Please ask the office to check Railway logs.")
        return ConversationHandler.END


async def abort_job_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    data = query.data.split("|")
    reason_key = data[2] if len(data) > 2 else "other"
    reason = ABORT_REASONS.get(reason_key, "Other reason")

    abort_job = context.user_data.get("abort_job")
    if not abort_job:
        await query.message.reply_text("Abort job session expired. Please try again from the job button.")
        return ConversationHandler.END

    abort_job["reason"] = reason

    await query.message.reply_text(
        f"Abort reason selected: {reason}\n\n"
        "Add a few lines explaining why, or type SKIP to submit with this reason only.\n\n"
        "Type /cancel to cancel without changing the job."
    )
    return ABORTJOB_NOTES


async def abort_job_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    abort_job = context.user_data.get("abort_job")

    if not abort_job:
        await update.message.reply_text("Abort job session expired. Please try again from the job button.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    notes = update.message.text.strip()
    if notes.lower() in ["skip", "none", "n/a", "na"]:
        notes = ""

    try:
        site_id = abort_job["site_id"]
        jobs_list_id = abort_job["jobs_list_id"]
        item_id = abort_job["item_id"]
        fields = abort_job["fields"]
        engineer_name = abort_job["engineer_name"]
        engineer_lookup_id = abort_job["engineer_lookup_id"]
        reason = abort_job.get("reason", "Other reason")

        extra = reason
        if notes:
            extra += f" | Notes: {notes}"

        updated_log = append_engineer_log(fields, engineer_name, "Aborted Attendance", extra)
        assigned_ids = get_assigned_engineer_ids(fields)
        is_final_engineer = len(assigned_ids) <= 1

        update_fields = {
            "EngineerVisitLog": updated_log,
            "TelegramNotified": False,
            "WorksheetSubmitted": False,
        }

        if is_final_engineer:
            update_fields["Status"] = AWAITING_DEPLOYMENT_STATUS
            update_fields["JobOutcome"] = "Aborted by Engineer"
            update_fields.update(clear_engineer_assignment_payload())
        else:
            update_fields.update(remove_current_engineer_assignment_payload(fields, engineer_lookup_id))

        update_list_item_fields(site_id, jobs_list_id, item_id, update_fields)
        update_active_day_live_status(site_id, abort_job["user_id"], "Aborted Attendance", get_job_reference(fields))

        await update.message.reply_text(
            f"Attendance aborted and sent back for reassignment.\n\n"
            f"CDR Number: {fields.get('CDRNumber', '')}\n"
            f"Reason: {reason}",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )

        await notify_helpdesk(
            context,
            (
                f"Job aborted by engineer\n\n"
                f"CDR Number: {fields.get('CDRNumber', '')}\n"
                f"Engineer: {engineer_name}\n"
                f"Site: {fields.get('SiteName', '')}\n"
                f"Reason: {reason}\n"
                f"Notes: {notes or 'None'}\n\n"
                "Job has been returned for reassignment."
            ),
        )

    except Exception as e:
        print(f"ERROR aborting job: {e}")
        await update.message.reply_text("There was an error aborting this job. Please ask the office to check Railway logs.", reply_markup=get_main_menu(await get_role_for_update(update)))
    finally:
        context.user_data.pop("abort_job", None)

    return ConversationHandler.END


async def abort_job_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("abort_job", None)
    await update.message.reply_text("Abort job cancelled. No changes made.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def status_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    try:
        data = query.data.split("|")
        action = data[0]
        item_id = data[1]

        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)

        user_id = str(query.from_user.id)
        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await query.message.reply_text("You are not set up as an engineer.")
            return

        if not engineer_has_active_day(site_id, user_id):
            await query.message.reply_text(
                "Please start your day first using 🟢 Start Day or /startday. Job buttons are locked until your day has started."
            )
            return

        job = find_job_by_item_id(jobs_data, item_id)

        if not job:
            await query.message.reply_text("Could not find this job.")
            return

        fields = job["fields"]

        if is_closed_job(fields):
            await query.message.reply_text(
                "This job has already been closed or returned to the office. No further action is required."
            )
            return

        assigned_ids = get_assigned_engineer_ids(fields)

        if current_engineer["lookup_id"] not in assigned_ids:
            await query.message.reply_text("You are not assigned to this job.")
            return

        if action == "abort_job":
            await query.message.reply_text("Opening abort job flow. If nothing happens, tap the Abort Attendance button again.")
            return

        if action == "complete_help":
            cdr_number = fields.get("CDRNumber", "")

            allowed, reason = can_click_action(
                fields,
                current_engineer["name"],
                "Completed",
            )

            if not allowed:
                await query.message.reply_text(reason)
                return

            await query.message.reply_text(
                f"To complete this job and submit the worksheet, type:\n\n/complete {cdr_number}"
            )
            return

        if action == "noaccess":
            allowed, reason = can_click_action(
                fields,
                current_engineer["name"],
                "No Access",
            )

            if not allowed:
                await query.message.reply_text(reason)
                return

            await query.message.reply_text(
                "Why was there no access?",
                reply_markup=get_no_access_reason_keyboard(item_id),
            )
            return

        if action == "noaccess_reason":
            reason_key = data[2] if len(data) > 2 else "other"
            no_access_reason = NO_ACCESS_REASONS.get(reason_key, "Other / see notes")

            allowed, reason = can_click_action(
                fields,
                current_engineer["name"],
                "No Access",
            )

            if not allowed:
                await query.message.reply_text(reason)
                return

            return await begin_worksheet_for_job(
                update,
                context,
                item_id=item_id,
                outcome="No Access",
                no_access_reason=no_access_reason,
            )

        if action == "confirm_outcome":
            selected_outcome = data[2]

            allowed, reason = can_click_action(
                fields,
                current_engineer["name"],
                selected_outcome,
            )

            if not allowed:
                await query.message.reply_text(reason)
                return

            confirm_buttons = InlineKeyboardMarkup([
                [
                    InlineKeyboardButton(
                        "Yes, confirm",
                        callback_data=f"start_worksheet|{item_id}|{selected_outcome}",
                    ),
                ],
                [
                    InlineKeyboardButton(
                        "Cancel",
                        callback_data=f"cancel_outcome|{item_id}",
                    ),
                ],
            ])

            await query.message.reply_text(
                f"Are you sure you want to mark this job as:\n\n{selected_outcome}?",
                reply_markup=confirm_buttons,
            )
            return

        if action == "cancel_outcome":
            await query.message.reply_text("Cancelled. No changes made.")
            return

        if action == "status":
            selected_status = data[2]

            if selected_status not in ["Travelling", "On Site"]:
                await query.message.reply_text("Unknown status selected.")
                return

            allowed, reason = can_click_action(
                fields,
                current_engineer["name"],
                selected_status,
            )

            if not allowed:
                await query.message.reply_text(reason)
                return

            updated_log = append_engineer_log(
                fields,
                current_engineer["name"],
                selected_status,
            )

            update_list_item_fields(
                site_id,
                jobs_list_id,
                item_id,
                {
                    "Status": selected_status,
                    "EngineerVisitLog": updated_log,
                    "WorksheetSubmitted": False,
                },
            )

            update_active_day_live_status(
                site_id,
                user_id,
                selected_status,
                get_job_reference(fields),
            )

            await query.message.reply_text(
                f"Status updated:\n\n{fields.get('CDRNumber', '')} → {selected_status}"
            )

            await notify_helpdesk(
                context,
                (
                    f"Job update\n\n"
                    f"CDR Number: {fields.get('CDRNumber', '')}\n"
                    f"Engineer: {current_engineer['name']}\n"
                    f"Update: {selected_status}\n"
                    f"Site: {fields.get('SiteName', '')}"
                ),
            )

            return

        if action == "outcome":
            selected_outcome = data[2]

            if selected_outcome not in ["No Access", "Revisit Required"]:
                await query.message.reply_text("Unknown outcome selected.")
                return

            allowed, reason = can_click_action(
                fields,
                current_engineer["name"],
                selected_outcome,
            )

            if not allowed:
                await query.message.reply_text(reason)
                return

            updated_log = append_engineer_log(
                fields,
                current_engineer["name"],
                selected_outcome,
            )

            assigned_ids = get_assigned_engineer_ids(fields)
            is_final_engineer = len(assigned_ids) <= 1

            update_fields = {
                "JobOutcome": selected_outcome,
                "EngineerVisitLog": updated_log,
            }

            if is_final_engineer:
                update_fields["Status"] = AWAITING_DEPLOYMENT_STATUS
                update_fields["TelegramNotified"] = False
                update_fields["WorksheetSubmitted"] = False
                update_fields.update(clear_engineer_assignment_payload())
            else:
                update_fields.update(
                    remove_current_engineer_assignment_payload(
                        fields,
                        current_engineer["lookup_id"],
                    )
                )

            update_list_item_fields(site_id, jobs_list_id, item_id, update_fields)

            update_active_day_live_status(
                site_id,
                user_id,
                selected_outcome,
                get_job_reference(fields),
            )

            await query.message.reply_text(
                f"Updated:\n\n"
                f"{fields.get('CDRNumber', '')} → {selected_outcome}\n"
                f"You have been removed from this job."
            )

            await notify_helpdesk(
                context,
                (
                    f"Job outcome selected\n\n"
                    f"CDR Number: {fields.get('CDRNumber', '')}\n"
                    f"Engineer: {current_engineer['name']}\n"
                    f"Outcome: {selected_outcome}\n"
                    f"Final engineer: {'Yes' if is_final_engineer else 'No'}\n"
                    f"Site: {fields.get('SiteName', '')}"
                ),
            )

            return

    except Exception as e:
        print(f"ERROR updating status/outcome: {e}")
        await query.message.reply_text("There was an error updating the job status.")



def can_start_completion(fields, engineer_name):
    status = str(fields.get("Status", "") or "").strip()
    outcome = str(fields.get("JobOutcome", "") or "").strip()

    # Only Completed blocks automatic dispatch permanently.
    # No Access/Revisit may be previous outcomes and must allow re-dispatch
    # once the office assigns an engineer again and TelegramNotified is False.
    if outcome == "Completed":
        return False, "This job has already been closed or returned to the office. No further action is required."

    if status in ["Completed", "No Access", "Revisit Required"]:
        return False, "This job has already been closed or returned to the office. No further action is required."

    # Completion is allowed from On Site. Also allow Travelling as a fallback,
    # but the normal path should still be Travelling > On Site > Complete.
    if status not in ["On Site", "Travelling", ASSIGNED_STATUS, AWAITING_DEPLOYMENT_STATUS, LEGACY_AWAITING_DEPLOYMENT_STATUS, ""]:
        return False, f"This job is currently marked as {status}. It cannot be completed from this status."

    if not engineer_has_logged(fields, engineer_name, "On Site"):
        return False, "You need to click On Site before completing this job."

    return True, ""


async def begin_worksheet_for_job(update: Update, context: ContextTypes.DEFAULT_TYPE, item_id=None, cdr_number=None, outcome="Completed", no_access_reason=""):
    """Start a worksheet from either the Complete/Revisit/No Access button or legacy /complete."""
    try:
        is_callback = update.callback_query is not None
        sender = update.callback_query.message if is_callback else update.message
        user = update.callback_query.from_user if is_callback else update.effective_user
        user_id = str(user.id)

        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)
        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await sender.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        if not engineer_has_active_day(site_id, user_id):
            await sender.reply_text(
                "Please start your day first using 🟢 Start Day or /startday before updating jobs.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        job = find_job_by_item_id(jobs_data, item_id) if item_id else find_job_by_cdr(jobs_data, cdr_number)

        if not job:
            await sender.reply_text("Could not find this job. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
            return ConversationHandler.END

        fields = job["fields"]

        if is_closed_job(fields):
            await sender.reply_text(
                "This job has already been closed or returned to the office. No further action is required.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        assigned_ids = get_assigned_engineer_ids(fields)
        if current_engineer["lookup_id"] not in assigned_ids:
            await sender.reply_text("You are not assigned to this job.", reply_markup=get_main_menu(await get_role_for_update(update)))
            return ConversationHandler.END

        allowed, reason = can_click_action(fields, current_engineer["name"], outcome)
        if not allowed:
            await sender.reply_text(reason, reply_markup=get_main_menu(await get_role_for_update(update)))
            return ConversationHandler.END

        cdr = fields.get("CDRNumber", "")
        context.user_data["worksheet"] = {
            "cdr_number": cdr,
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "item_id": job["id"],
            "engineer_name": current_engineer["name"],
            "engineer_lookup_id": current_engineer["lookup_id"],
            "fields": fields,
            "photo_links": [],
            "ClientSignatureRequired": False,
            "ClientSignatureReceived": False,
            "JobOutcome": outcome,
            "NoAccessReason": no_access_reason,
        }

        if outcome == "Completed":
            first_question = "What work was completed?"
        elif outcome == "Revisit Required":
            first_question = "What was done today, and why is a revisit required?"
        else:
            first_question = "Add any extra no access notes. If there is nothing else to add, type None."

        no_access_line = f"\nNo Access reason: {no_access_reason}\n" if outcome == "No Access" and no_access_reason else ""

        await sender.reply_text(
            f"Starting worksheet for {cdr}.\n\n"
            f"Outcome: {outcome}"
            f"{no_access_line}\n"
            f"You can type /cancel at any point before submitting.\n\n"
            f"{first_question}"
        )
        return WORK_COMPLETED

    except Exception as e:
        print(f"ERROR starting worksheet: {e}")
        target = update.callback_query.message if update.callback_query else update.message
        await target.reply_text("There was an error starting the worksheet.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END


async def complete_button_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    try:
        _, item_id, outcome = query.data.split("|", 2)
    except Exception:
        await query.message.reply_text("Could not start this worksheet. Tap 📋 My Jobs and try again.")
        return ConversationHandler.END

    return await begin_worksheet_for_job(update, context, item_id=item_id, outcome=outcome)


async def noaccess_reason_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Conversation entry point after an engineer chooses a No Access reason.

    This must be an entry point on the worksheet ConversationHandler; otherwise
    the next typed note is received by the bot but no worksheet state is active.
    """
    query = update.callback_query
    await query.answer()

    try:
        _, item_id, reason_key = query.data.split("|", 2)
    except Exception:
        await query.message.reply_text("Could not start the No Access worksheet. Tap 📋 My Jobs and try again.")
        return ConversationHandler.END

    no_access_reason = NO_ACCESS_REASONS.get(reason_key, "Other / see notes")

    return await begin_worksheet_for_job(
        update,
        context,
        item_id=item_id,
        outcome="No Access",
        no_access_reason=no_access_reason,
    )


async def complete_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text(
            "Use 📋 My Jobs, then tap Complete, Revisit or No Access on the job card.\n\n"
            "The old /complete CDR number method still works if needed: /complete CDR00001",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    return await begin_worksheet_for_job(
        update,
        context,
        cdr_number=context.args[0].strip(),
        outcome="Completed",
    )

async def worksheet_work_completed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, WORK_COMPLETED)
    if menu_result is not None:
        return menu_result

    if is_bot_menu_text(update.message.text):
        await update.message.reply_text("That is a menu button, so I have not added it to the worksheet. Please type the work completed, or type /cancel.")
        return WORK_COMPLETED

    context.user_data["worksheet"]["WorkCompleted"] = update.message.text
    await update.message.reply_text("What materials were used? Type None if none.")
    return MATERIALS_USED


async def worksheet_materials_used(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, MATERIALS_USED)
    if menu_result is not None:
        return menu_result

    if is_bot_menu_text(update.message.text):
        await update.message.reply_text("That is a menu button, so I have not added it to the worksheet. Please type materials used, or type None.")
        return MATERIALS_USED

    context.user_data["worksheet"]["MaterialsUsed"] = update.message.text
    await update.message.reply_text(
        "Is a follow-on required?",
        reply_markup=get_yes_no_keyboard("follow_on_required"),
    )
    return FOLLOW_ON_REQUIRED


async def worksheet_follow_on_required(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, FOLLOW_ON_REQUIRED)
    if menu_result is not None:
        return menu_result

    answer = update.message.text.strip().lower()

    if answer not in ["yes", "no", "y", "n"]:
        await update.message.reply_text("Please tap Yes or No.")
        return FOLLOW_ON_REQUIRED

    follow_on_required = answer in ["yes", "y"]
    context.user_data["worksheet"]["FollowOnRequired"] = follow_on_required

    if follow_on_required:
        await update.message.reply_text("What follow-on is required?")
        return FOLLOW_ON_NOTES

    context.user_data["worksheet"]["FollowOnNotes"] = ""
    await update.message.reply_text("Any engineer notes? Type None if none.")
    return ENGINEER_NOTES


async def worksheet_follow_on_required_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    answer = query.data.split("|", 1)[1]
    follow_on_required = answer == "yes"
    worksheet["FollowOnRequired"] = follow_on_required

    if follow_on_required:
        await query.message.reply_text("What follow-on is required?")
        return FOLLOW_ON_NOTES

    worksheet["FollowOnNotes"] = ""
    await query.message.reply_text("Any engineer notes? Type None if none.")
    return ENGINEER_NOTES


async def worksheet_follow_on_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, FOLLOW_ON_NOTES)
    if menu_result is not None:
        return menu_result

    if is_bot_menu_text(update.message.text):
        await update.message.reply_text("That is a menu button, so I have not added it to the worksheet. Please type the follow-on required, or type /cancel.")
        return FOLLOW_ON_NOTES

    context.user_data["worksheet"]["FollowOnNotes"] = update.message.text
    await update.message.reply_text("Any engineer notes? Type None if none.")
    return ENGINEER_NOTES


async def worksheet_engineer_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, ENGINEER_NOTES)
    if menu_result is not None:
        return menu_result

    if is_bot_menu_text(update.message.text):
        await update.message.reply_text("That is a menu button, so I have not added it to the worksheet. Please type engineer notes, or type None.")
        return ENGINEER_NOTES

    context.user_data["worksheet"]["EngineerCompletionNotes"] = update.message.text

    await update.message.reply_text(
        "Upload job photos now.\n\n"
        "Send one or more photos, then type DONE when finished.\n"
        "If no photos are needed, type DONE."
    )

    return PHOTOS


async def worksheet_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, PHOTOS)
    if menu_result is not None:
        return menu_result

    worksheet = context.user_data["worksheet"]

    if update.message.text and update.message.text.strip().upper() == "DONE":
        await update.message.reply_text(
            "Is a client signature required?",
            reply_markup=get_yes_no_keyboard("signature_required"),
        )
        return SIGNATURE_REQUIRED

    if update.message.photo:
        site_id = worksheet["site_id"]
        cdr_number = worksheet["cdr_number"]

        photo = update.message.photo[-1]
        telegram_file = await context.bot.get_file(photo.file_id)
        file_bytes = await telegram_file.download_as_bytearray()

        timestamp = datetime.now(UK_TZ).strftime("%Y%m%d_%H%M%S")
        file_name = f"{cdr_number}_{timestamp}_{photo.file_unique_id}.jpg"

        photo_link = upload_photo_to_sharepoint(
            site_id,
            cdr_number,
            file_name,
            bytes(file_bytes),
        )

        worksheet["photo_links"].append(photo_link)
        return PHOTOS

    await update.message.reply_text("Please send the required van photos, or type DONE once all 3 have been uploaded.")
    return PHOTOS


async def worksheet_signature_required(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, SIGNATURE_REQUIRED)
    if menu_result is not None:
        return menu_result

    answer = update.message.text.strip().lower()
    worksheet = context.user_data["worksheet"]

    if answer not in ["yes", "no", "y", "n"]:
        await update.message.reply_text("Please tap Yes or No.")
        return SIGNATURE_REQUIRED

    signature_required = answer in ["yes", "y"]
    worksheet["ClientSignatureRequired"] = signature_required

    if not signature_required:
        await update.message.reply_text(
            build_review_text(worksheet),
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    try:
        token = create_signature_token_for_job(
            worksheet["site_id"],
            worksheet["jobs_list_id"],
            worksheet["item_id"],
        )

        signature_url = build_signature_url(worksheet["cdr_number"], token)
        worksheet["SignatureToken"] = token
        worksheet["SignatureUrl"] = signature_url

        await update.message.reply_text(
            "Client signature required.\n\n"
            "Open this link on your phone and ask the client to sign:\n\n"
            f"{signature_url}\n\n"
            "Once signed, tap Signed. If no client is available, tap Skip.",
            reply_markup=get_signed_skip_keyboard(),
        )

        return SIGNATURE_WAITING

    except Exception as e:
        print(f"ERROR creating signature link: {e}")
        await update.message.reply_text(
            "There was an error creating the signature link. Tap Skip to continue without a signature."
        )
        return SIGNATURE_WAITING


async def worksheet_signature_required_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    signature_required = query.data.split("|", 1)[1] == "yes"
    worksheet["ClientSignatureRequired"] = signature_required

    if not signature_required:
        await query.message.reply_text(
            build_review_text(worksheet),
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    try:
        token = create_signature_token_for_job(
            worksheet["site_id"],
            worksheet["jobs_list_id"],
            worksheet["item_id"],
        )

        signature_url = build_signature_url(worksheet["cdr_number"], token)
        worksheet["SignatureToken"] = token
        worksheet["SignatureUrl"] = signature_url

        await query.message.reply_text(
            "Client signature required.\n\n"
            "Open this link on your phone and ask the client to sign:\n\n"
            f"{signature_url}\n\n"
            "Once signed, tap Signed. If no client is available, tap Skip.",
            reply_markup=get_signed_skip_keyboard(),
        )

        return SIGNATURE_WAITING

    except Exception as e:
        print(f"ERROR creating signature link: {e}")
        await query.message.reply_text(
            "There was an error creating the signature link. Tap Skip to continue without a signature.",
            reply_markup=get_signed_skip_keyboard(),
        )
        return SIGNATURE_WAITING


async def worksheet_signature_waiting(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, SIGNATURE_WAITING)
    if menu_result is not None:
        return menu_result

    answer = update.message.text.strip().upper()
    worksheet = context.user_data["worksheet"]

    if answer == "SKIP":
        worksheet["ClientSignatureReceived"] = False
        await update.message.reply_text(
            build_review_text(worksheet),
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    if answer != "SIGNED":
        await update.message.reply_text(
            "Please tap Signed once the client has signed, or Skip to continue without a signature.",
            reply_markup=get_signed_skip_keyboard(),
        )
        return SIGNATURE_WAITING

    latest_jobs = get_list_items(worksheet["site_id"], worksheet["jobs_list_id"])
    job = find_job_by_item_id(latest_jobs, worksheet["item_id"])

    if not job:
        await update.message.reply_text("Could not check the signature. Please try SIGNED again or type SKIP.")
        return SIGNATURE_WAITING

    fields = job["fields"]

    if bool_field(fields.get("ClientSignatureReceived")):
        worksheet["ClientSignatureReceived"] = True
        worksheet["ClientSignatureName"] = fields.get("ClientSignatureName", "")
        worksheet["ClientSignatureLink"] = fields.get("ClientSignatureLink", "")
        await update.message.reply_text("Signature received.")
        await update.message.reply_text(
            build_review_text(worksheet),
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    await update.message.reply_text(
        "I cannot see the signature yet. Make sure the client pressed Submit Signature, then tap Signed again.",
        reply_markup=get_signed_skip_keyboard(),
    )
    return SIGNATURE_WAITING


async def worksheet_signature_waiting_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    action = query.data.split("|", 1)[1]

    if action == "skip":
        worksheet["ClientSignatureReceived"] = False
        await query.message.reply_text(
            build_review_text(worksheet),
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    latest_jobs = get_list_items(worksheet["site_id"], worksheet["jobs_list_id"])
    job = find_job_by_item_id(latest_jobs, worksheet["item_id"])

    if not job:
        await query.message.reply_text(
            "Could not check the signature. Tap Signed again or Skip.",
            reply_markup=get_signed_skip_keyboard(),
        )
        return SIGNATURE_WAITING

    fields = job["fields"]

    if bool_field(fields.get("ClientSignatureReceived")):
        worksheet["ClientSignatureReceived"] = True
        worksheet["ClientSignatureName"] = fields.get("ClientSignatureName", "")
        worksheet["ClientSignatureLink"] = fields.get("ClientSignatureLink", "")
        await query.message.reply_text("Signature received.")
        await query.message.reply_text(
            build_review_text(worksheet),
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    await query.message.reply_text(
        "I cannot see the signature yet. Make sure the client pressed Submit Signature, then tap Signed again.",
        reply_markup=get_signed_skip_keyboard(),
    )
    return SIGNATURE_WAITING




def clean_pdf_text(value):
    text = str(value or "").strip()
    return text if text and text.lower() != "none" else "N/A"


def safe_pdf_filename(value):
    cleaned = "".join(ch for ch in str(value or "").strip() if ch.isalnum() or ch in ["-", "_", " "]).strip()
    return cleaned.replace(" ", "_") or "worksheet"



def clean_engineer_log_extra(value):
    """Keep EngineerVisitLog extra text to one safe line so it can be parsed back into the PDF."""
    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def build_visit_comment_extra(worksheet):
    """Create a compact visit comment for EngineerVisitLog and the final worksheet PDF."""
    parts = []

    outcome = worksheet.get("JobOutcome", "Completed")
    if outcome:
        parts.append(f"Outcome: {outcome}")

    if outcome == "No Access" and worksheet.get("NoAccessReason"):
        parts.append(f"No Access Reason: {worksheet.get('NoAccessReason')}")

    work_completed = clean_engineer_log_extra(worksheet.get("WorkCompleted", ""))
    if work_completed and work_completed.upper() != "N/A":
        parts.append(f"Work/Comments: {work_completed}")

    materials = clean_engineer_log_extra(worksheet.get("MaterialsUsed", ""))
    if materials and materials.lower() not in ["none", "n/a", "no"]:
        parts.append(f"Materials Used: {materials}")

    if worksheet.get("FollowOnRequired"):
        follow_on = clean_engineer_log_extra(worksheet.get("FollowOnNotes", ""))
        parts.append(f"Follow-on Required: Yes{(' - ' + follow_on) if follow_on else ''}")
    else:
        parts.append("Follow-on Required: No")

    engineer_notes = clean_engineer_log_extra(worksheet.get("EngineerCompletionNotes", ""))
    if engineer_notes and engineer_notes.lower() not in ["none", "n/a", "no"]:
        parts.append(f"Engineer Notes: {engineer_notes}")

    return " | ".join(parts) or "Worksheet submitted"


def normalise_visit_note(value):
    note = clean_engineer_log_extra(value)
    if not note:
        return ""
    if note == "Worksheet submitted":
        return ""
    return note.replace(" | ", "\n")


def build_engineer_comments_for_pdf(visits, worksheet, fields):
    """Show comments for every visit, not just the final completed visit."""
    comment_blocks = []

    for visit in visits:
        note = normalise_visit_note(visit.get("notes", ""))
        if not note:
            continue

        heading_bits = [
            clean_pdf_text(visit.get("date")),
            clean_pdf_text(visit.get("engineer")),
            clean_pdf_text(visit.get("status")),
        ]
        heading = " - ".join([bit for bit in heading_bits if bit and bit != "N/A"])
        comment_blocks.append(f"{heading}\n{note}")

    # Fallback for older jobs where previous worksheet comments were not yet being written into EngineerVisitLog.
    if not comment_blocks:
        comments = worksheet.get("WorkCompleted", "")
        if worksheet.get("MaterialsUsed") and str(worksheet.get("MaterialsUsed")).strip().lower() != "none":
            comments += f"\n\nMaterials Used: {worksheet.get('MaterialsUsed')}"
        if worksheet.get("FollowOnRequired"):
            comments += f"\n\nFollow-on Required: Yes\n{worksheet.get('FollowOnNotes', '')}"
        else:
            comments += "\n\nFollow-on Required: No"
        if worksheet.get("EngineerCompletionNotes") and str(worksheet.get("EngineerCompletionNotes")).strip().lower() != "none":
            comments += f"\n\nEngineer Notes: {worksheet.get('EngineerCompletionNotes')}"
        return comments

    return "\n\n".join(comment_blocks)


def parse_engineer_visit_log(log_text):
    """
    Build one worksheet visit row per attendance from EngineerVisitLog.
    Expected log line format:
    dd/mm/yyyy HH:MM - Engineer Name - Action - optional notes
    """
    visits = []
    active_by_engineer = {}

    for raw_line in str(log_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue

        match = re.match(
            r"^(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})\s+-\s+(.+?)\s+-\s+(.+?)(?:\s+-\s+(.*))?$",
            line,
        )
        if not match:
            continue

        visit_date = match.group(1)
        visit_time = match.group(2)
        engineer = match.group(3).strip()
        action = match.group(4).strip()
        extra = (match.group(5) or "").strip()
        key = engineer.lower()

        # Office/helpdesk audit lines are useful in SharePoint, but they are not
        # engineer attendances and should not appear on the customer worksheet.
        if key in ["helpdesk", "admin", "office"] or action in ["Reassigned", "Job logged via Telegram"]:
            continue

        if action == "Travelling":
            # Start a fresh attendance for this engineer.
            active_by_engineer[key] = {
                "date": visit_date,
                "travel": visit_time,
                "on_site": "",
                "engineer": engineer,
                "status": "Travelling",
                "off_site": "",
                "notes": extra,
            }
            visits.append(active_by_engineer[key])
            continue

        if key not in active_by_engineer:
            active_by_engineer[key] = {
                "date": visit_date,
                "travel": "",
                "on_site": "",
                "engineer": engineer,
                "status": "",
                "off_site": "",
                "notes": "",
            }
            visits.append(active_by_engineer[key])

        current = active_by_engineer[key]

        if action == "On Site":
            current["on_site"] = visit_time
            current["status"] = "On Site"
        elif action in ["Completed", "No Access", "Revisit Required"]:
            current["status"] = action
            current["off_site"] = visit_time
            if extra:
                current["notes"] = extra
            # This attendance is closed. A future Travelling click by the same engineer
            # will create a new row rather than overwriting this one.
            active_by_engineer.pop(key, None)
        else:
            current["status"] = action
            if extra:
                current["notes"] = extra

    return visits


def get_signature_image_bytes(site_id, cdr_number):
    """Download the latest saved client signature image from SharePoint, if available."""
    try:
        drive_id = get_drive_id(site_id, PHOTO_LIBRARY)
        folder_path = f"{SIGNATURE_BASE_FOLDER}/{cdr_number}"
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children"
        response = requests.get(url, headers=get_headers())

        if response.status_code != 200:
            print(f"Could not list signature folder for {cdr_number}: {response.text}")
            return None

        files = [item for item in response.json().get("value", []) if "file" in item]
        if not files:
            return None

        files.sort(key=lambda item: item.get("lastModifiedDateTime", ""), reverse=True)
        item_id = files[0]["id"]
        content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
        content_response = requests.get(content_url, headers=get_headers(content_type=False))

        if content_response.status_code != 200:
            print(f"Could not download signature image for {cdr_number}: {content_response.text}")
            return None

        return content_response.content
    except Exception as e:
        print(f"ERROR getting signature image: {e}")
        return None


def get_logo_for_pdf(max_width=48 * mm, max_height=20 * mm):
    for path in ["cdr-logo.png", "CDR-logo.png", "logo.png"]:
        if os.path.exists(path):
            try:
                logo = Image(path)
                logo._restrictSize(max_width, max_height)
                return logo
            except Exception as e:
                print(f"Could not load logo {path}: {e}")
    return Paragraph("<b>CDR</b>", ParagraphStyle("LogoFallback", fontSize=20, textColor=colors.HexColor("#f58220")))


def build_worksheet_pdf_bytes(worksheet, fields, updated_log, outcome, site_id=None):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    normal = ParagraphStyle("CDRNormal", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.8, leading=11)
    small = ParagraphStyle("CDRSmall", parent=normal, fontSize=7.6, leading=9)
    section = ParagraphStyle("CDRSection", parent=normal, fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.white, leading=11)
    title = ParagraphStyle("CDRTitle", parent=normal, fontName="Helvetica-Bold", fontSize=16, textColor=colors.HexColor("#111827"), alignment=2)

    def escape_pdf(value):
        return clean_pdf_text(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def ptxt(value):
        return Paragraph(escape_pdf(value).replace("\n", "<br/>"), normal)

    def section_box(title_text, body_text, height_padding=8):
        table = Table(
            [[Paragraph(title_text, section)], [ptxt(body_text)]],
            colWidths=[182 * mm],
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f58220")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
            ("LINEBELOW", (0, 0), (-1, 0), 0.75, colors.HexColor("#333333")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), height_padding),
            ("BOTTOMPADDING", (0, 0), (-1, -1), height_padding),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ]))
        return table

    cdr_number = worksheet.get("cdr_number") or fields.get("CDRNumber", "")
    date_logged = format_sharepoint_date(fields.get("Date", ""))
    date_complete = datetime.now(UK_TZ).strftime("%d/%m/%Y")

    customer_name = get_field_value(fields, "CustomerName", "Customer Name") or get_field_value(fields, "ClientName", "Client Name") or ""
    customer_address = get_field_value(fields, "CustomerAddress", "Customer Address") or ""
    customer_details = "\n".join([v for v in [customer_name, customer_address] if str(v or "").strip()])

    site_details = "\n".join([v for v in [
        fields.get("SiteName", ""),
        fields.get("Address", ""),
    ] if str(v or "").strip()])

    order_number = get_field_value(fields, "CustomerOrderNumber", "Customer Order Number", "OrderNumber", "Order Number") or ""
    job_category = get_field_value(fields, "JobCategory", "Job Category") or ""

    story = []

    header = Table(
        [[get_logo_for_pdf(), Paragraph("JOB WORKSHEET", title)]],
        colWidths=[85 * mm, 97 * mm],
    )
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(header)

    company = Table([[
        Paragraph(
            "<b>CDR M&amp;E Services Ltd</b><br/>"
            "6 Mandale Park, Urlay Nook Road, Egglescliffe, Stockton-on-Tees, TS16 0TA<br/>"
            "Telephone: 01642 057939 &nbsp;&nbsp; Email: helpdesk@cdrme.co.uk<br/>"
            "VAT Number: 397715249 &nbsp;&nbsp; Company No.: 13744971",
            small,
        )
    ]], colWidths=[182 * mm])
    company.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f7f7f7")),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(company)
    story.append(Spacer(1, 7))

    details = Table(
        [[Paragraph("<b>Customer Details</b><br/>" + escape_pdf(customer_details).replace("\n", "<br/>"), normal),
          Paragraph("<b>Site Details</b><br/>" + escape_pdf(site_details).replace("\n", "<br/>"), normal)]],
        colWidths=[91 * mm, 91 * mm],
    )
    details.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
        ("INNERGRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 18),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(details)
    story.append(Spacer(1, 7))

    job_data = [
        ["Job Number:", clean_pdf_text(cdr_number), "Customer Order Number:", clean_pdf_text(order_number)],
        ["Date Logged:", clean_pdf_text(date_logged), "Job Category:", clean_pdf_text(job_category)],
        ["Date Complete:", clean_pdf_text(date_complete), "Status:", clean_pdf_text(outcome)],
    ]
    job_table = Table(job_data, colWidths=[38 * mm, 53 * mm, 46 * mm, 45 * mm])
    job_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(job_table)
    story.append(Spacer(1, 7))

    story.append(section_box("Description", fields.get("Task", "") or fields.get("Description", "") or fields.get("Notes", ""), 7))
    story.append(Spacer(1, 7))

    visits_data = [["Date", "Travel", "On-Site", "Engineer", "Status", "Off-Site"]]
    visits = parse_engineer_visit_log(updated_log)
    if not visits:
        visits = [{
            "date": datetime.now(UK_TZ).strftime("%d/%m/%Y"),
            "travel": "",
            "on_site": "",
            "engineer": worksheet.get("engineer_name", ""),
            "status": outcome,
            "off_site": datetime.now(UK_TZ).strftime("%H:%M"),
        }]

    for visit in visits:
        visits_data.append([
            clean_pdf_text(visit.get("date")),
            clean_pdf_text(visit.get("travel")),
            clean_pdf_text(visit.get("on_site")),
            clean_pdf_text(visit.get("engineer")),
            clean_pdf_text(visit.get("status")),
            clean_pdf_text(visit.get("off_site")),
        ])

    visits_table = Table(visits_data, colWidths=[27 * mm, 25 * mm, 25 * mm, 45 * mm, 35 * mm, 25 * mm])
    visits_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f58220")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(Paragraph("<b>Visits</b>", normal))
    story.append(visits_table)
    story.append(Spacer(1, 7))

    comments = build_engineer_comments_for_pdf(visits, worksheet, fields)

    story.append(section_box("Engineer Comment", comments, 8))
    story.append(Spacer(1, 7))

    if worksheet.get("ClientSignatureRequired"):
        signature_rows = []
        if worksheet.get("ClientSignatureReceived"):
            signature_rows.append([Paragraph(
                f"<b>Client Name:</b> {escape_pdf(worksheet.get('ClientSignatureName', ''))}<br/>"
                f"<b>Signed Digitally:</b> Yes",
                normal,
            )])
            signature_bytes = get_signature_image_bytes(site_id, cdr_number) if site_id else None
            if signature_bytes:
                try:
                    sig_img = Image(BytesIO(signature_bytes))
                    sig_img._restrictSize(80 * mm, 28 * mm)
                    signature_rows.append([sig_img])
                except Exception as e:
                    print(f"Could not embed signature image: {e}")
                    signature_rows.append([Paragraph("Signature image saved in SharePoint but could not be embedded.", normal)])
            else:
                signature_rows.append([Paragraph("Signature image saved in SharePoint but could not be embedded.", normal)])
        else:
            signature_rows.append([Paragraph("Client signature required but not received.", normal)])
    else:
        signature_rows = [[Paragraph("Client signature not required.", normal)]]

    signature = Table([[Paragraph("Client Signature", section)]] + signature_rows, colWidths=[182 * mm])
    signature.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f58220")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#333333")),
        ("LINEBELOW", (0, 0), (-1, 0), 0.75, colors.HexColor("#333333")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(signature)

    story.append(Spacer(1, 5))
    story.append(Paragraph("CDR M&amp;E Services Ltd | 01642 057939 | helpdesk@cdrme.co.uk", small))

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def generate_and_upload_worksheet_pdf(site_id, jobs_list_id, item_id, worksheet, fields, updated_log, outcome):
    try:
        pdf_bytes = build_worksheet_pdf_bytes(worksheet, fields, updated_log, outcome, site_id)
        cdr_number = worksheet.get("cdr_number") or fields.get("CDRNumber", "JOB")
        file_name = f"{safe_pdf_filename(cdr_number)}_worksheet_{datetime.now(UK_TZ).strftime('%Y%m%d_%H%M%S')}.pdf"

        worksheet_link = upload_file_to_sharepoint(
            site_id,
            WORKSHEET_BASE_FOLDER,
            safe_folder_name(cdr_number),
            file_name,
            pdf_bytes,
        )
        return worksheet_link
    except Exception as e:
        print(f"ERROR generating worksheet PDF: {e}")
        return ""


def build_worksheet_update_fields(worksheet, fields, updated_log, outcome, is_final_engineer, worksheet_pdf_link=""):
    fields_to_update = {
        "WorkCompleted": worksheet.get("WorkCompleted", ""),
        "MaterialsUsed": worksheet.get("MaterialsUsed", ""),
        "FollowOnRequired": worksheet.get("FollowOnRequired", False),
        "FollowOnNotes": worksheet.get("FollowOnNotes", ""),
        "EngineerCompletionNotes": worksheet.get("EngineerCompletionNotes", ""),
        "WorksheetSubmitted": True,
        "EngineerVisitLog": updated_log,
        "ClientSignatureRequired": worksheet.get("ClientSignatureRequired", False),
    }

    if worksheet_pdf_link:
        fields_to_update["WorksheetPDFLink"] = worksheet_pdf_link
        fields_to_update["WorksheetGenerated"] = True

    if outcome == "No Access" and worksheet.get("NoAccessReason"):
        fields_to_update["NoAccessReason"] = worksheet.get("NoAccessReason")

    if is_final_engineer:
        fields_to_update["JobOutcome"] = outcome

        if outcome == "Completed":
            fields_to_update["Status"] = COMPLETED_STATUS
            if is_notified(fields):
                fields_to_update["TelegramNotified"] = True
        else:
            fields_to_update["Status"] = AWAITING_DEPLOYMENT_STATUS
            fields_to_update["TelegramNotified"] = False

        fields_to_update.update(clear_engineer_assignment_payload())
    else:
        fields_to_update.update(
            remove_current_engineer_assignment_payload(
                fields,
                worksheet["engineer_lookup_id"],
            )
        )

    return fields_to_update

async def worksheet_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, REVIEW)
    if menu_result is not None:
        return menu_result

    answer = update.message.text.strip().upper()
    worksheet = context.user_data["worksheet"]

    if answer == "CANCEL":
        context.user_data.pop("worksheet", None)
        await update.message.reply_text("Worksheet cancelled. Nothing has been submitted.")
        return ConversationHandler.END

    if answer == "RESTART":
        worksheet["WorkCompleted"] = ""
        worksheet["MaterialsUsed"] = ""
        worksheet["FollowOnRequired"] = False
        worksheet["FollowOnNotes"] = ""
        worksheet["EngineerCompletionNotes"] = ""
        worksheet["photo_links"] = []
        worksheet["ClientSignatureRequired"] = False
        worksheet["ClientSignatureReceived"] = False

        await update.message.reply_text(
            f"Restarting worksheet for {worksheet['cdr_number']}.\n\n"
            f"What work was completed?"
        )

        return WORK_COMPLETED

    if answer == "SUBMIT":
        site_id = worksheet["site_id"]
        jobs_list_id = worksheet["jobs_list_id"]
        item_id = worksheet["item_id"]

        latest_jobs = get_list_items(site_id, jobs_list_id)
        job = find_job_by_item_id(latest_jobs, item_id)
        fields = job["fields"] if job else worksheet["fields"]

        if is_closed_job(fields):
            context.user_data.pop("worksheet", None)
            await update.message.reply_text(
                "This job has already been closed or returned to the office. Worksheet has not been submitted again.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        assigned_ids_latest = get_assigned_engineer_ids(fields)
        if worksheet["engineer_lookup_id"] not in assigned_ids_latest:
            context.user_data.pop("worksheet", None)
            await update.message.reply_text(
                "You are no longer assigned to this job. Worksheet has not been submitted.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        outcome = worksheet.get("JobOutcome", "Completed")

        updated_log = append_engineer_log(
            fields,
            worksheet["engineer_name"],
            outcome,
            build_visit_comment_extra(worksheet),
        )

        assigned_ids = get_assigned_engineer_ids(fields)
        is_final_engineer = len(assigned_ids) <= 1

        worksheet_pdf_link = ""
        if is_final_engineer and outcome == "Completed":
            worksheet_pdf_link = generate_and_upload_worksheet_pdf(
                site_id,
                jobs_list_id,
                item_id,
                worksheet,
                fields,
                updated_log,
                outcome,
            )

        fields_to_update = build_worksheet_update_fields(
            worksheet,
            fields,
            updated_log,
            outcome,
            is_final_engineer,
            worksheet_pdf_link,
        )

        update_list_item_fields(site_id, jobs_list_id, item_id, fields_to_update)

        update_active_day_live_status(
            site_id,
            str(update.effective_user.id),
            outcome,
            get_job_reference(fields),
        )

        final_pdf_text = "\n\nFinal worksheet PDF generated." if worksheet_pdf_link else "\n\nFinal PDF will generate when the last assigned engineer completes."
        await update.message.reply_text(
            f"Worksheet submitted:\n\n{worksheet['cdr_number']} → {outcome}" + final_pdf_text,
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )

        await notify_helpdesk(
            context,
            (
                f"Worksheet submitted\n\n"
                f"CDR Number: {worksheet['cdr_number']}\n"
                f"Engineer: {worksheet['engineer_name']}\n"
                f"Outcome: {outcome}\n"
                f"No Access reason: {worksheet.get('NoAccessReason', 'N/A') if outcome == 'No Access' else 'N/A'}\n"
                f"Final engineer: {'Yes' if is_final_engineer else 'No'}\n"
                f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n"
                f"Client signature required: {'Yes' if worksheet.get('ClientSignatureRequired') else 'No'}\n"
                f"Client signature received: {'Yes' if worksheet.get('ClientSignatureReceived') else 'No'}"
            ),
        )

        context.user_data.pop("worksheet", None)
        return ConversationHandler.END

    await update.message.reply_text(
        "Please tap Submit worksheet, Restart worksheet or Cancel.",
        reply_markup=get_review_keyboard(),
    )
    return REVIEW


async def worksheet_review_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    action = query.data.split("|", 1)[1]

    if action == "cancel":
        context.user_data.pop("worksheet", None)
        await query.message.reply_text("Worksheet cancelled. Nothing has been submitted.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    if action == "restart":
        worksheet["WorkCompleted"] = ""
        worksheet["MaterialsUsed"] = ""
        worksheet["FollowOnRequired"] = False
        worksheet["FollowOnNotes"] = ""
        worksheet["EngineerCompletionNotes"] = ""
        worksheet["photo_links"] = []
        worksheet["ClientSignatureRequired"] = False
        worksheet["ClientSignatureReceived"] = False

        await query.message.reply_text(
            f"Restarting worksheet for {worksheet['cdr_number']}.\n\n"
            f"What work was completed?"
        )
        return WORK_COMPLETED

    if action != "submit":
        await query.message.reply_text(
            "Please tap Submit worksheet, Restart worksheet or Cancel.",
            reply_markup=get_review_keyboard(),
        )
        return REVIEW

    site_id = worksheet["site_id"]
    jobs_list_id = worksheet["jobs_list_id"]
    item_id = worksheet["item_id"]

    latest_jobs = get_list_items(site_id, jobs_list_id)
    job = find_job_by_item_id(latest_jobs, item_id)
    fields = job["fields"] if job else worksheet["fields"]

    if is_closed_job(fields):
        context.user_data.pop("worksheet", None)
        await query.message.reply_text(
            "This job has already been closed or returned to the office. Worksheet has not been submitted again.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    assigned_ids_latest = get_assigned_engineer_ids(fields)
    if worksheet["engineer_lookup_id"] not in assigned_ids_latest:
        context.user_data.pop("worksheet", None)
        await query.message.reply_text(
            "You are no longer assigned to this job. Worksheet has not been submitted.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    outcome = worksheet.get("JobOutcome", "Completed")

    updated_log = append_engineer_log(
        fields,
        worksheet["engineer_name"],
        outcome,
        build_visit_comment_extra(worksheet),
    )

    assigned_ids = get_assigned_engineer_ids(fields)
    is_final_engineer = len(assigned_ids) <= 1

    worksheet_pdf_link = ""
    if is_final_engineer and outcome == "Completed":
        worksheet_pdf_link = generate_and_upload_worksheet_pdf(
            site_id,
            jobs_list_id,
            item_id,
            worksheet,
            fields,
            updated_log,
            outcome,
        )

    fields_to_update = build_worksheet_update_fields(
        worksheet,
        fields,
        updated_log,
        outcome,
        is_final_engineer,
        worksheet_pdf_link,
    )

    update_list_item_fields(site_id, jobs_list_id, item_id, fields_to_update)

    update_active_day_live_status(
        site_id,
        str(query.from_user.id),
        outcome,
        get_job_reference(fields),
    )

    final_pdf_text = "\n\nFinal worksheet PDF generated." if worksheet_pdf_link else "\n\nFinal PDF will generate when the last assigned engineer completes."
    await query.message.reply_text(
        f"Worksheet submitted:\n\n{worksheet['cdr_number']} → {outcome}" + final_pdf_text,
        reply_markup=get_main_menu(await get_role_for_update(update)),
    )

    await notify_helpdesk(
        context,
        (
            f"Worksheet submitted\n\n"
            f"CDR Number: {worksheet['cdr_number']}\n"
            f"Engineer: {worksheet['engineer_name']}\n"
            f"Outcome: {outcome}\n"
            f"No Access reason: {worksheet.get('NoAccessReason', 'N/A') if outcome == 'No Access' else 'N/A'}\n"
            f"Final engineer: {'Yes' if is_final_engineer else 'No'}\n"
            f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n"
            f"Client signature required: {'Yes' if worksheet.get('ClientSignatureRequired') else 'No'}\n"
            f"Client signature received: {'Yes' if worksheet.get('ClientSignatureReceived') else 'No'}"
        ),
    )

    context.user_data.pop("worksheet", None)
    return ConversationHandler.END


async def worksheet_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("worksheet", None)
    await update.message.reply_text("Worksheet cancelled. Nothing has been submitted.")
    return ConversationHandler.END



def debug_job_dispatch_decision(fields):
    """
    Debug logging is intentionally disabled for normal use.
    The dispatch scheduler runs frequently, so printing every job every cycle
    makes Railway logs noisy even when no job is being sent.
    """
    return


async def send_new_jobs(app):
    try:
        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        _, engineers_by_lookup = build_engineer_maps(engineers)

        sent_job_ids = set()

        for job in jobs_data:
            fields = job["fields"]
            item_id = job["id"]

            if not should_auto_send_job(fields):
                continue

            assigned_ids = get_assigned_engineer_ids(fields)
            sent_to_any_engineer = False

            for engineer_id in assigned_ids:
                engineer = engineers_by_lookup.get(engineer_id)

                if not engineer:
                    print(f"WARNING: No engineer record found for lookup ID {engineer_id} on job {fields.get('CDRNumber', item_id)}")
                    continue

                try:
                    await app.bot.send_message(
                        chat_id=engineer["telegram_id"],
                        text="New job assigned:\n\n" + format_job(fields, engineer["name"]),
                        reply_markup=get_job_buttons(item_id, fields.get("Address", "")),
                    )
                    sent_to_any_engineer = True
                except Exception as e:
                    print(f"WARNING: Could not send job {fields.get('CDRNumber', item_id)} to engineer {engineer_id}: {e}")

            if sent_to_any_engineer:
                sent_job_ids.add(item_id)

        for item_id in sent_job_ids:
            try:
                update_list_item_fields(
                    site_id,
                    jobs_list_id,
                    item_id,
                    {
                        "TelegramNotified": True,
                        "Status": ASSIGNED_STATUS,
                    },
                )
            except Exception as e:
                print(f"ERROR marking job {item_id} as TelegramNotified: {e}")

    except Exception as e:
        print(f"ERROR sending new jobs: {e}")


async def remind_active_engineers_to_end_day(app):
    """Daily reminder so engineers do not forget to close their day."""
    try:
        site_id = get_site_id()
        day_logs_list_id = get_list_id(site_id, DAY_LOGS_LIST)
        day_logs = get_list_items(site_id, day_logs_list_id)
        today = get_today_iso()

        for log in day_logs:
            fields = log.get("fields", {})
            status = str(fields.get("Status", ""))
            telegram_id = str(fields.get("EngineerTelegramID") or fields.get("Engineer Telegram ID") or "")
            raw_work_date = fields.get("WorkDate") or fields.get("Work Date") or ""
            parsed_work_date = sharepoint_date_to_uk_date(raw_work_date)
            log_date = parsed_work_date.isoformat() if parsed_work_date else str(raw_work_date)[:10]

            if status == DAY_ACTIVE_STATUS and telegram_id and log_date == today:
                try:
                    await app.bot.send_message(
                        chat_id=telegram_id,
                        text=(
                            "End of day reminder: if you have finished work, please tap 🏁 End Day. "
                            "This keeps timesheets, mileage and pay hours correct."
                        ),
                        reply_markup=get_main_menu(await get_role_for_update(update)),
                    )
                except Exception as e:
                    print(f"WARNING: Could not send end-day reminder to {telegram_id}: {e}")

    except Exception as e:
        print(f"ERROR sending end-day reminders: {e}")


async def post_init(app):
    try:
        await app.bot.delete_webhook(drop_pending_updates=True)
        print("Webhook removed.")
    except Exception as e:
        print(f"Could not remove webhook: {e}")

    scheduler = AsyncIOScheduler(timezone=UK_TZ)

    scheduler.add_job(
        send_new_jobs,
        trigger="interval",
        seconds=30,
        args=[app],
    )

    scheduler.add_job(
        remind_active_engineers_to_end_day,
        trigger="cron",
        hour=16,
        minute=45,
        args=[app],
    )

    scheduler.start()
    print("Scheduler started.")


telegram_app = (
    ApplicationBuilder()
    .token(BOT_TOKEN)
    .post_init(post_init)
    .build()
)

startday_handler = ConversationHandler(
    entry_points=[
        CommandHandler("startday", startday_start),
        MessageHandler(filters.Regex(f"^{MENU_START_DAY}$"), startday_start),
    ],
    states={
        START_DAY_CONFIRM: [
            CallbackQueryHandler(startday_confirm_button, pattern=r"^startday_confirm\|"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, startday_confirm),
        ],
        START_DAY_VAN_REG: [MessageHandler(filters.TEXT & ~filters.COMMAND, startday_van_reg)],
        START_DAY_START_MILEAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, startday_start_mileage)],
        START_DAY_VAN_CHECK: [MessageHandler(filters.TEXT & ~filters.COMMAND, startday_van_check)],
        START_DAY_VAN_PHOTOS: [
            MessageHandler(filters.PHOTO, startday_van_photos),
            MessageHandler(filters.TEXT & ~filters.COMMAND, startday_van_photos),
        ],
    },
    fallbacks=[CommandHandler("cancel", startday_cancel)],
)

endday_handler = ConversationHandler(
    entry_points=[
        CommandHandler("endday", endday_start),
        MessageHandler(filters.Regex(f"^{MENU_END_DAY}$"), endday_start),
    ],
    states={
        END_DAY_CONFIRM: [
            CallbackQueryHandler(endday_confirm_button, pattern=r"^endday_confirm\|"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, endday_confirm),
        ],
        END_DAY_MILEAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, endday_mileage)],
    },
    fallbacks=[CommandHandler("cancel", endday_cancel)],
)


worksheet_handler = ConversationHandler(
    entry_points=[
        CallbackQueryHandler(complete_button_start, pattern="^start_worksheet\\|"),
        CallbackQueryHandler(noaccess_reason_start, pattern="^noaccess_reason\\|"),
        CommandHandler("complete", complete_start),
    ],
    states={
        WORK_COMPLETED: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_work_completed)],
        MATERIALS_USED: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_materials_used)],
        FOLLOW_ON_REQUIRED: [
            CallbackQueryHandler(worksheet_follow_on_required_button, pattern=r"^follow_on_required\|"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_follow_on_required),
        ],
        FOLLOW_ON_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_follow_on_notes)],
        ENGINEER_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_engineer_notes)],
        PHOTOS: [
            MessageHandler(filters.PHOTO, worksheet_photos),
            MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_photos),
        ],
        SIGNATURE_REQUIRED: [
            CallbackQueryHandler(worksheet_signature_required_button, pattern=r"^signature_required\|"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_signature_required),
        ],
        SIGNATURE_WAITING: [
            CallbackQueryHandler(worksheet_signature_waiting_button, pattern=r"^signature_waiting\|"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_signature_waiting),
        ],
        REVIEW: [
            CallbackQueryHandler(worksheet_review_button, pattern=r"^review\|"),
            MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_review),
        ],
    },
    fallbacks=[CommandHandler("cancel", worksheet_cancel)],
)


bugidea_handler = ConversationHandler(
    entry_points=[
        CommandHandler("bugidea", bugidea_start),
        MessageHandler(filters.Regex(f"^{MENU_BUG_IDEA}$"), bugidea_start),
    ],
    states={
        BUG_IDEA_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, bugidea_text)],
    },
    fallbacks=[CommandHandler("cancel", bugidea_cancel)],
)



reassign_handler = ConversationHandler(
    entry_points=[
        CommandHandler("reassign", reassign_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_REASSIGN_JOB)}$"), reassign_start),
    ],
    states={
        REASSIGN_CDR_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, reassign_cdr_number)],
        REASSIGN_REMOVE_ENGINEERS: [MessageHandler(filters.TEXT & ~filters.COMMAND, reassign_remove_engineers)],
        REASSIGN_ASSIGN_ENGINEERS: [MessageHandler(filters.TEXT & ~filters.COMMAND, reassign_assign_engineers)],
        REASSIGN_REASON: [MessageHandler(filters.TEXT & ~filters.COMMAND, reassign_reason)],
        REASSIGN_REVIEW: [MessageHandler(filters.TEXT & ~filters.COMMAND, reassign_review)],
    },
    fallbacks=[CommandHandler("cancel", reassign_cancel)],
)


openjobs_handler = ConversationHandler(
    entry_points=[
        CommandHandler("openjobs", openjobs_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_OPEN_JOBS)}$"), openjobs_start),
    ],
    states={
        OPENJOBS_FILTER: [MessageHandler(filters.TEXT & ~filters.COMMAND, openjobs_filter)],
        OPENJOBS_SELECT: [MessageHandler(filters.TEXT & ~filters.COMMAND, openjobs_select)],
    },
    fallbacks=[CommandHandler("cancel", openjobs_cancel)],
)



canceljob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("canceljob", canceljob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_CANCEL_JOB)}$"), canceljob_start),
    ],
    states={
        CANCELJOB_CDR_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, canceljob_cdr_number)],
        CANCELJOB_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, canceljob_confirm)],
    },
    fallbacks=[CommandHandler("cancel", canceljob_cancel)],
)


deletejob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("deletejob", deletejob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_DELETE_JOB)}$"), deletejob_start),
    ],
    states={
        DELETEJOB_CDR_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, deletejob_cdr_number)],
        DELETEJOB_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, deletejob_confirm)],
    },
    fallbacks=[CommandHandler("cancel", deletejob_cancel)],
)


receipt_handler = ConversationHandler(
    entry_points=[
        CommandHandler("receipts", receipt_start),
        CommandHandler("uploadreceipts", receipt_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_UPLOAD_RECEIPTS)}$"), receipt_start),
    ],
    states={
        RECEIPT_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receipt_date)],
        RECEIPT_UPLOADS: [MessageHandler((filters.PHOTO | filters.Document.ALL | (filters.TEXT & ~filters.COMMAND)), receipt_uploads)],
    },
    fallbacks=[CommandHandler("cancel", receipt_cancel)],
)


findjob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("findjob", findjob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_FIND_JOB)}$"), findjob_start),
    ],
    states={
        FINDJOB_SEARCH: [MessageHandler(filters.TEXT & ~filters.COMMAND, findjob_search)],
        FINDJOB_SELECT: [MessageHandler(filters.TEXT & ~filters.COMMAND, findjob_select)],
    },
    fallbacks=[CommandHandler("cancel", findjob_cancel)],
)


logjob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("logjob", logjob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_LOG_JOB)}$"), logjob_start),
    ],
    states={
        LOGJOB_CDR_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_cdr_number)],
        LOGJOB_CUSTOMER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_customer_name)],
        LOGJOB_CUSTOMER_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_customer_address)],
        LOGJOB_SITE_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_site_name)],
        LOGJOB_SITE_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_site_confirm)],
        LOGJOB_SITE_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_site_address)],
        LOGJOB_SITE_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_site_notes)],
        LOGJOB_CONTACT: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_contact)],
        LOGJOB_TASK: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_task)],
        LOGJOB_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_notes)],
        LOGJOB_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_date)],
        LOGJOB_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_time)],
        LOGJOB_CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_category)],
        LOGJOB_ORDER_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_order_number)],
        LOGJOB_ASSIGN_ENGINEERS: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_assign_engineers)],
        LOGJOB_REVIEW: [MessageHandler(filters.TEXT & ~filters.COMMAND, logjob_review)],
    },
    fallbacks=[CommandHandler("cancel", logjob_cancel)],
)

abortjob_handler = ConversationHandler(
    entry_points=[
        CallbackQueryHandler(abort_job_start, pattern="^abort_job\\|"),
    ],
    states={
        ABORTJOB_REASON: [CallbackQueryHandler(abort_job_reason, pattern="^abort_reason\\|")],
        ABORTJOB_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, abort_job_notes)],
    },
    fallbacks=[CommandHandler("cancel", abort_job_cancel)],
)


telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(CommandHandler("helpdesk", helpdesk_start))
telegram_app.add_handler(startday_handler)
telegram_app.add_handler(endday_handler)
telegram_app.add_handler(worksheet_handler)
telegram_app.add_handler(bugidea_handler)
telegram_app.add_handler(receipt_handler)
telegram_app.add_handler(logjob_handler)
telegram_app.add_handler(reassign_handler)
telegram_app.add_handler(openjobs_handler)
telegram_app.add_handler(canceljob_handler)
telegram_app.add_handler(deletejob_handler)
telegram_app.add_handler(findjob_handler)
telegram_app.add_handler(abortjob_handler)
telegram_app.add_handler(MessageHandler(filters.Regex(f"^({MENU_MY_JOBS}|{MENU_BUG_IDEA}|{MENU_UPLOAD_RECEIPTS}|{MENU_HELPDESK}|{MENU_LOG_JOB}|{MENU_REASSIGN_JOB}|{MENU_OPEN_JOBS}|{MENU_FIND_JOB}|{MENU_CANCEL_JOB}|{MENU_DELETE_JOB}|{MENU_ENGINEER_MENU})$"), menu_button))
telegram_app.add_handler(CallbackQueryHandler(status_button))

if __name__ == "__main__":
    threading.Thread(target=run_signature_web_server, daemon=True).start()
    print(f"Signature web server running on port {PORT}")
    print(f"Bot running... PID={os.getpid()} | Build={BUILD_VERSION}")

    telegram_app.run_polling(
        drop_pending_updates=True,
        close_loop=False,
        allowed_updates=Update.ALL_TYPES,
    )
