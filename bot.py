import os
import re
import base64
import secrets
import hmac
import hashlib
import time
import threading
import zipfile
import asyncio
from io import BytesIO
from urllib.parse import quote_plus, urlparse
from xml.sax.saxutils import escape as xml_escape
import warnings
import requests
import msal
from datetime import datetime
from zoneinfo import ZoneInfo

from cdr_core import (
    now_log_time,
    graph_datetime_now,
    get_today_iso,
    AWAITING_DEPLOYMENT_STATUS,
    LEGACY_AWAITING_DEPLOYMENT_STATUS,
    ASSIGNED_STATUS,
    TRAVELLING_STATUS,
    ON_SITE_STATUS,
    COMPLETED_STATUS,
    DAY_ACTIVE_STATUS,
    DAY_CLOSED_STATUS,
    user_can_use_helpdesk,
    role_counts_for_utilisation,
    is_vehicle_check_exempt_role,
    get_assigned_engineer_ids as core_get_assigned_engineer_ids,
    is_closed_job as core_is_closed_job,
    engineer_has_logged as core_engineer_has_logged,
    validate_job_action as core_validate_job_action,
)
print("✅ CDR Core loaded in bot:", now_log_time())
print("✅ CDR job logic loaded in bot")

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, Response, RedirectResponse
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, InputMediaPhoto, ReplyKeyboardRemove
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
CDR_ELECTRICAL_CHAT_ID = os.getenv("CDR_ELECTRICAL_CHAT_ID")
CDR_MECHANICAL_CHAT_ID = os.getenv("CDR_MECHANICAL_CHAT_ID")
SIGNATURE_BASE_URL = os.getenv("SIGNATURE_BASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
PORTAL_BASE_URL = os.getenv("PORTAL_BASE_URL", "").strip().rstrip("/")
ASSET_API_KEY = os.getenv("ASSET_API_KEY", "").strip()
PORT = int(os.getenv("PORT", "8000"))
BUILD_VERSION = "phase1-recovery-v47-admin-openai-assistant"

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"
DAY_LOGS_LIST = "Engineer Day Logs"
BUG_IDEAS_LIST = "Bug Ideas"
BOT_USERS_LIST = "Bot Users"
TASK_ACTIVITIES_LIST = "Task Activities"


PHOTO_LIBRARY = "Documents"
PHOTO_BASE_FOLDER = "15 - ENGINEER JOB PHOTOS"
SIGNATURE_BASE_FOLDER = "16 - CLIENT SIGNATURES"
VAN_CHECK_PHOTO_BASE_FOLDER = "17 - VAN CHECK PHOTOS"
WORKSHEET_BASE_FOLDER = "18 - JOB WORKSHEETS"
RECEIPT_BASE_FOLDER = "19 - RECEIPTS"



MENU_START_DAY = "🟢 Start Day"
MENU_MY_JOBS = "📋 My Jobs"
MENU_END_DAY = "🏁 End Day"
MENU_BUG_IDEA = "🐞 Bug / Ideas"
MENU_UPLOAD_RECEIPTS = "🧾 Receipts / Returns"
MENU_REQUEST_JOB = "📣 Request Job"
MENU_QUOTE_REMINDER = "📋 Task / Activity"
MENU_MESSAGE_ENGINEER = "📢 Message Engineer"
MENU_CANCEL_TASK_ACTIVITY = "❌ Cancel Task / Activity"
MENU_HELPDESK = "🧰 Helpdesk"
MENU_LOG_JOB = "➕ Log Job"
MENU_REASSIGN_JOB = "🔁 Reassign Job"
MENU_REOPEN_JOB = "♻️ Reopen Job"
MENU_OPEN_JOBS = "📋 Open Jobs"
MENU_FIND_JOB = "🔎 Find Job"
MENU_CANCEL_JOB = "❌ Cancel Job"
MENU_DELETE_JOB = "🗑 Delete Job"
MENU_ASK_CHATBOT = "🤖 Ask ChatGPT"
MENU_CREATE_ASSET = "🏷 Create Asset"
MENU_ENGINEER_MENU = "👷 Engineer Menu"


UK_TZ = ZoneInfo("Europe/London")


VAN_CHECK_INTERVAL_DAYS = 14
# Apprentice role supported. Apprentice users can be excluded from van/mileage prompts in the Start/End Day flow.

WORK_COMPLETED = 0
MATERIALS_USED = 1
FOLLOW_ON_REQUIRED = 2
FOLLOW_ON_NOTES = 3
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
RECEIPT_TYPE = 59
RECEIPT_DATE = 60
RECEIPT_UPLOADS = 61
ABORTJOB_REASON = 62
ABORTJOB_NOTES = 63
QUOTE_CLIENT = 64
QUOTE_ADDRESS = 65
QUOTE_DATE = 66
QUOTE_TIME = 67
QUOTE_SCOPE = 68
QUOTE_REVIEW = 69
QUOTE_RECIPIENT = 70
MESSAGE_ENGINEER_SELECT = 71
MESSAGE_ENGINEER_TEMPLATE = 72
MESSAGE_ENGINEER_TEXT = 73
MESSAGE_ENGINEER_REVIEW = 74
CANCEL_TASK_SELECT = 75
CANCEL_TASK_REASON = 76
CANCEL_TASK_REVIEW = 77
REASSIGN_DATE = 78
REOPENJOB_CDR_NUMBER = 79
REOPENJOB_DATE = 80
REOPENJOB_ASSIGN_ENGINEERS = 81
REOPENJOB_REASON = 82
REOPENJOB_REVIEW = 83
ASK_CHATBOT_QUESTION = 84
ASSET_CUSTOMER = 85
ASSET_SITE_NAME = 86
ASSET_SITE_ADDRESS = 87
ASSET_CATEGORY = 88
ASSET_TYPE = 89
ASSET_DESCRIPTION = 90
ASSET_LOCATION = 91
ASSET_MANUFACTURER = 92
ASSET_MODEL = 93
ASSET_SERIAL = 94
ASSET_PHOTO = 95
ASSET_NOTES = 96

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



def is_group_chat(update):
    return update.effective_chat and update.effective_chat.type != "private"


def is_private_chat(update):
    return update.effective_chat and update.effective_chat.type == "private"


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

    # SharePoint Graph rejects the entire update if even one optional column name
    # is not present in the list. Remove only the rejected field and retry so
    # core fields that do exist will still be saved.
    payload = dict(fields_to_update or {})
    ignored_fields = []

    while True:
        response = requests.patch(
            url,
            headers=get_headers(),
            json=payload,
        )

        if response.status_code in [200, 204]:
            if ignored_fields:
                print(f"WARNING: SharePoint ignored missing field(s) on item {item_id}: {', '.join(ignored_fields)}")
            return

        error_text = response.text or ""
        match = re.search(r"Field '([^']+)' is not recognized", error_text)
        if match:
            missing_field = match.group(1)
            if missing_field in payload:
                payload.pop(missing_field, None)
                payload.pop(f"{missing_field}@odata.type", None)
                ignored_fields.append(missing_field)
                continue

        raise Exception(f"Could not update item {item_id}: {response.text}")


def update_list_item_fields_in_safe_chunks(site_id, list_id, item_id, fields_to_update, required_fields=None, chunk_size=6):
    """Update SharePoint using small chunks, with individual fallback for optional fields.

    This is used by Helpdesk approval/rejection because those updates can include
    many optional pending-review/audit fields. One invalid optional field should not
    block the core status/outcome update. Lookup collection @odata.type keys are
    kept with their matching LookupId field so Graph receives a valid payload.
    """
    required_fields = set(required_fields or [])
    fields_to_update = dict(fields_to_update or {})

    grouped_items = []
    consumed = set()
    for key, value in fields_to_update.items():
        if key in consumed or key.endswith("@odata.type"):
            continue

        group = {}
        odata_key = f"{key}@odata.type"
        if odata_key in fields_to_update:
            group[odata_key] = fields_to_update[odata_key]
            consumed.add(odata_key)
        group[key] = value
        consumed.add(key)
        grouped_items.append(group)

    for key, value in fields_to_update.items():
        if key not in consumed:
            grouped_items.append({key: value})
            consumed.add(key)

    for start in range(0, len(grouped_items), chunk_size):
        chunk_groups = grouped_items[start:start + chunk_size]
        chunk_payload = {}
        for group in chunk_groups:
            chunk_payload.update(group)

        try:
            update_list_item_fields(site_id, list_id, item_id, chunk_payload)
            continue
        except Exception as chunk_error:
            print(f"WARNING: SharePoint chunk update failed for item {item_id}; retrying fields individually: {chunk_error}")

        for group in chunk_groups:
            required_group = any(key in required_fields for key in group.keys())
            try:
                update_list_item_fields(site_id, list_id, item_id, group)
            except Exception as field_error:
                if required_group:
                    raise
                print(f"WARNING: SharePoint optional field skipped on item {item_id}: {', '.join(group.keys())} - {field_error}")


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


def get_remaining_assigned_engineer_ids(fields, current_lookup_id):
    """Return assigned engineer lookup IDs excluding the current engineer.

    SharePoint multi-person fields can come back in slightly different shapes
    depending on Graph/List expansion. Older code only read the display lookup
    field (Engineer), which meant one engineer submitting a worksheet could
    accidentally overwrite the whole assignment list with an empty list.
    This helper centralises the calculation so non-final submissions only remove
    the submitting engineer and preserve every other assigned engineer.
    """
    current_lookup_id = str(current_lookup_id or "").strip()
    remaining = []
    seen = set()

    for lookup_id in get_assigned_engineer_ids(fields):
        lookup_id = str(lookup_id or "").strip()
        if not lookup_id or lookup_id == current_lookup_id or lookup_id in seen:
            continue
        remaining.append(int(lookup_id))
        seen.add(lookup_id)

    return remaining


def remove_current_engineer_assignment_payload(fields, current_lookup_id):
    remaining_ids = get_remaining_assigned_engineer_ids(fields, current_lookup_id)

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

    reset_actions = ["Completed", "No Access", "Revisit Required", "Aborted Attendance", "Travelling Reverted"]
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
        [MENU_REQUEST_JOB, MENU_CREATE_ASSET],
    ]

    # Only Admin users should be able to switch from the Engineer menu
    # into the Helpdesk menu. Helpdesk users stay in the Helpdesk menu,
    # and Engineer users stay in the Engineer menu.
    if include_helpdesk_menu:
        rows.append([MENU_HELPDESK, MENU_ASK_CHATBOT])

    return ReplyKeyboardMarkup(
        rows,
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def get_helpdesk_menu(include_engineer_menu=False):
    buttons = [
        MENU_LOG_JOB,
        MENU_REASSIGN_JOB,
        MENU_REOPEN_JOB,
        MENU_OPEN_JOBS,
        MENU_FIND_JOB,
        MENU_CANCEL_JOB,
    ]

    if include_engineer_menu:
        buttons.append(MENU_DELETE_JOB)

    buttons.extend([
        MENU_QUOTE_REMINDER,
        MENU_MESSAGE_ENGINEER,
        MENU_CANCEL_TASK_ACTIVITY,
        MENU_UPLOAD_RECEIPTS,
        MENU_CREATE_ASSET,
        MENU_BUG_IDEA,
    ])

    if include_engineer_menu:
        buttons.append(MENU_ASK_CHATBOT)
        buttons.append(MENU_ENGINEER_MENU)

    rows = [buttons[i:i + 3] for i in range(0, len(buttons), 3)]

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

        if role.lower() in ["admin", "helpdesk", "engineer", "apprentice"]:
            return role.title()

        return "Engineer"

    print(f"Telegram ID {telegram_id} not found in Engineers list; no access granted.")
    return "Inactive"





JOB_CATEGORY_CHOICES = [
    "Electrical",
    "Mechanical",
    "Plumbing",
    "HVAC",
    "Fire",
    "Catering",
    "Building Fabric",
    "Other",
]


def is_blank_or_skip(value):
    return str(value or "").strip().lower() in ["", "skip", "none", "n/a", "na", "⏭️ skip", "n/a"]


def get_skip_keyboard():
    return ReplyKeyboardMarkup([["⏭️ Skip"]], resize_keyboard=True, one_time_keyboard=False)


def get_date_keyboard():
    return ReplyKeyboardMarkup(
        [["Today", "Tomorrow"], ["Custom"]],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def get_time_keyboard():
    return ReplyKeyboardMarkup(
        [["ASAP", "08:00"], ["11:00", "14:00"], ["Custom"]],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def get_category_keyboard():
    rows = []
    for i in range(0, len(JOB_CATEGORY_CHOICES), 2):
        rows.append(JOB_CATEGORY_CHOICES[i:i + 2])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True, one_time_keyboard=False)


def get_review_reply_keyboard():
    return ReplyKeyboardMarkup(
        [["✅ Yes", "❌ No"], ["🔄 Restart"]],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def get_assign_engineer_keyboard():
    return ReplyKeyboardMarkup(
        [["⏭️ Skip for now"], ["/cancel"]],
        resize_keyboard=True,
        one_time_keyboard=False,
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



def parse_task_activity_datetime(value):
    """Parse Task / Activity date input into ISO date and display text."""
    raw = str(value or "").strip()
    parsed_date = parse_helpdesk_job_date(raw)
    if parsed_date:
        return parsed_date, format_sharepoint_date(parsed_date) or parsed_date, ""

    # Accept date/time in one line: 15/05/2026 10:00 or Today 10:00
    parts = raw.split()
    if len(parts) >= 2:
        candidate_date = " ".join(parts[:-1])
        candidate_time = parts[-1]
        parsed_date = parse_helpdesk_job_date(candidate_date)
        parsed_time = normalise_helpdesk_time(candidate_time)
        if parsed_date:
            return parsed_date, format_sharepoint_date(parsed_date) or parsed_date, parsed_time or ""

    return None, "", ""


def create_task_activity_record(site_id, reminder):
    """Store Task / Activity so 08:00 reminders and cancellation survive restarts."""
    tasks_list_id = get_list_id(site_id, TASK_ACTIVITIES_LIST)
    recipients = reminder.get("recipients_selected") or [reminder.get("recipient", {})]
    telegram_ids = [str(r.get("telegram_id", "")).strip() for r in recipients if str(r.get("telegram_id", "")).strip()]
    names = [str(r.get("name", "")).strip() for r in recipients if str(r.get("name", "")).strip()]
    details = (
        f"Client/Site: {reminder.get('client', '')}\n"
        f"Address/Site: {reminder.get('address', '')}\n"
        f"Details: {reminder.get('scope', '')}\n"
        f"Sent by: {reminder.get('set_by', 'Helpdesk')}"
    )
    task_date = reminder.get("date") or get_today_iso()
    reminder_sent = task_date <= get_today_iso()
    payload = build_field_payload_for_list(
        site_id,
        tasks_list_id,
        {
            "Title": reminder.get("client") or reminder.get("scope") or "Task / Activity",
            "TaskDate": task_date,
            "Task Date": task_date,
            "TaskTime": reminder.get("time", ""),
            "Task Time": reminder.get("time", ""),
            "TaskDetails": details,
            "Task Details": details,
            "AssignedTelegramIDs": "\n".join(telegram_ids),
            "Assigned Telegram IDs": "\n".join(telegram_ids),
            "AssignedNames": ", ".join(names),
            "Assigned Names": ", ".join(names),
            "Status": "Active",
            "ReminderSent": reminder_sent,
            "Reminder Sent": reminder_sent,
            "CreatedBy": reminder.get("set_by", "Helpdesk"),
            "Created By": reminder.get("set_by", "Helpdesk"),
            "CreatedDateTime": graph_datetime_now(),
            "Created Date Time": graph_datetime_now(),
        },
    )
    return create_list_item_fields(site_id, tasks_list_id, payload)

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
        if role not in ["engineer", "admin", "apprentice"]:
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

    contact_line = f"Contact: {job.get('contact', '')}\n" if job.get("contact") else ""
    notes_line = f"Notes: {job.get('notes', '')}\n" if job.get("notes") else ""
    order_line = f"Order Number: {job.get('order_number', '')}\n" if job.get("order_number") else ""

    return (
        "Please review the new job before I create it in SharePoint:\n\n"
        f"CDR Number: {job.get('cdr_number', '')}\n\n"
        f"Customer:\n{job.get('customer_name', '')}\n\n"
        f"Site:\n{job.get('site_name', '')}\n\n"
        f"{contact_line}"
        f"Task: {job.get('task', '')}\n"
        f"{notes_line}"
        f"Date: {job.get('date_display', job.get('date', ''))}\n"
        f"Time: {job.get('time', '')}\n"
        f"Category: {job.get('category', '')}\n"
        f"{order_line}"
        f"Assigned To: {assigned_names}\n\n"
        "Reply YES to create and send, NO to cancel, or RESTART to start again."
    )

def build_helpdesk_job_fields(site_id, jobs_list_id, job, telegram_notified=False):
    job_notes = job.get("notes", "") or ""

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
            "SiteName": job["site_name"],
            "Site Name": job["site_name"],
            "ContactName": job.get("contact", ""),
            "Contact Name": job.get("contact", ""),
            "Task": job["task"],
            "Notes": job_notes,
            "CustomerOrderNumber": job.get("order_number", ""),
            "Customer Order Number": job.get("order_number", ""),
            "JobCategory": job.get("category", ""),
            "Job Category": job.get("category", ""),
            "Status": ASSIGNED_STATUS if job.get("assigned_engineers") else AWAITING_DEPLOYMENT_STATUS,
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
    date_text = data.get("new_date_display") or "Keep current date"

    return (
        "Please review this reassignment before I update SharePoint:\n\n"
        f"CDR Number: {fields.get('CDRNumber', data.get('cdr_number', ''))}\n"
        f"Site: {fields.get('SiteName', '')}\n"
        f"Current engineer(s): {current}\n"
        f"Remove: {removing}\n"
        f"Assign/send to: {assigning}\n"
        f"Final assigned engineer(s): {final}\n"
        f"Attendance date: {date_text}\n"
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
                reply_markup=get_job_buttons(item_id, fields.get("SiteName", "")),
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


def is_weekend_today():
    """Return True on Saturday/Sunday in UK time so scheduled reminders stay quiet at weekends."""
    return datetime.now(UK_TZ).weekday() >= 5


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


def upload_receipt_to_sharepoint(site_id, receipt_date, engineer_name, file_name, file_bytes, upload_type="Receipt"):
    upload_type = str(upload_type or "Receipt").strip().lower()

    if upload_type == "return":
        folder_name = f"Returns/{receipt_date}/{safe_folder_name(engineer_name)}"
    else:
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


def hours_between(start_dt, end_dt):
    if not start_dt or not end_dt or end_dt <= start_dt:
        return 0.0
    return (end_dt - start_dt).total_seconds() / 3600


def round_hours(value):
    try:
        return round(float(value or 0), 2)
    except Exception:
        return 0.0


def parse_engineer_log_datetime(date_text, time_text):
    try:
        parsed = datetime.strptime(f"{date_text} {time_text}", "%d/%m/%Y %H:%M")
        return parsed.replace(tzinfo=UK_TZ)
    except Exception:
        return None


def get_engineer_productive_intervals_for_day(jobs_data, engineer_name, work_date, start_dt=None, end_dt=None):
    """
    Build productive intervals from existing job action logs.

    This intentionally uses the job buttons engineers already press:
    Travelling -> On Site -> Completed/No Access/Revisit Required.
    No extra engineer workflow is required.
    """
    engineer_key = str(engineer_name or "").strip().lower()
    intervals = []

    if not engineer_key or not work_date:
        return intervals

    for job in jobs_data or []:
        fields = job.get("fields", {})
        log_text = get_field_value(fields, "EngineerVisitLog", "Engineer Visit Log") or ""

        try:
            visits = parse_engineer_visit_log(log_text)
        except Exception as e:
            print(f"WARNING: Could not parse productive intervals: {e}")
            visits = []

        for visit in visits:
            if str(visit.get("engineer", "")).strip().lower() != engineer_key:
                continue

            date_text = visit.get("date", "")
            try:
                visit_date = datetime.strptime(date_text, "%d/%m/%Y").date()
            except Exception:
                continue

            if visit_date != work_date:
                continue

            start_text = visit.get("travel") or visit.get("on_site")
            end_text = visit.get("off_site") or visit.get("on_site")
            interval_start = parse_engineer_log_datetime(date_text, start_text)
            interval_end = parse_engineer_log_datetime(date_text, end_text)

            if not interval_start or not interval_end or interval_end <= interval_start:
                continue

            if start_dt and interval_end < start_dt:
                continue
            if end_dt and interval_start > end_dt:
                continue

            if start_dt and interval_start < start_dt:
                interval_start = start_dt
            if end_dt and interval_end > end_dt:
                interval_end = end_dt

            intervals.append((interval_start, interval_end))

    return merge_time_intervals(intervals)


def merge_time_intervals(intervals):
    cleaned = sorted(
        [(start, end) for start, end in intervals if start and end and end > start],
        key=lambda item: item[0],
    )

    if not cleaned:
        return []

    merged = [cleaned[0]]
    for start, end in cleaned[1:]:
        last_start, last_end = merged[-1]
        if start <= last_end:
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))

    return merged


def calculate_activity_hours(start_dt, end_dt, jobs_data=None, engineer_name=""):
    """Calculate productive/inactive time from job logs for the engineer's day."""
    if not start_dt or not end_dt or end_dt <= start_dt:
        return {
            "productive_hours": 0.0,
            "inactive_hours": 0.0,
            "utilisation_percent": 0.0,
            "productive_intervals": [],
        }

    work_date = start_dt.astimezone(UK_TZ).date()
    intervals = get_engineer_productive_intervals_for_day(
        jobs_data or [],
        engineer_name,
        work_date,
        start_dt,
        end_dt,
    )

    productive_hours = sum(hours_between(start, end) for start, end in intervals)
    total_hours = hours_between(start_dt, end_dt)
    inactive_hours = max(0.0, total_hours - productive_hours)
    utilisation_percent = (productive_hours / total_hours * 100) if total_hours > 0 else 0.0

    return {
        "productive_hours": round_hours(productive_hours),
        "inactive_hours": round_hours(inactive_hours),
        "utilisation_percent": round_hours(utilisation_percent),
        "productive_intervals": intervals,
    }


def calculate_commute_deductions(start_dt, end_dt, productive_intervals):
    """
    Apply CDR's simple automatic commute rule without extra engineer inputs.

    - Morning: if the day starts before 08:00 and the first job travel/action is before 08:00,
      deduct up to 30 minutes from payable OOH.
    - Evening: if the day ends after 16:30, deduct up to 30 minutes from the gap after the
      last productive job action.
    """
    if not start_dt or not end_dt or end_dt <= start_dt:
        return {
            "morning_commute_deduction_hours": 0.0,
            "evening_commute_deduction_hours": 0.0,
            "commute_deduction_hours": 0.0,
        }

    normal_start = start_dt.replace(hour=8, minute=0, second=0, microsecond=0)
    normal_end = start_dt.replace(hour=16, minute=30, second=0, microsecond=0)

    morning_deduction = 0.0
    evening_deduction = 0.0

    if productive_intervals:
        first_productive_start = productive_intervals[0][0]
        last_productive_end = productive_intervals[-1][1]

        if start_dt < normal_start and first_productive_start < normal_start:
            pre_8_duration = hours_between(start_dt, min(end_dt, normal_start))
            morning_deduction = min(0.5, pre_8_duration)

        if end_dt > normal_end and last_productive_end < end_dt:
            commute_gap_start = max(last_productive_end, normal_end)
            post_work_gap = hours_between(commute_gap_start, end_dt)
            evening_deduction = min(0.5, post_work_gap)
    else:
        # No job actions were logged. Do not auto-deduct commute time because the day may have
        # been office/training/admin work. Office can review inactive time separately.
        morning_deduction = 0.0
        evening_deduction = 0.0

    total = morning_deduction + evening_deduction
    return {
        "morning_commute_deduction_hours": round_hours(morning_deduction),
        "evening_commute_deduction_hours": round_hours(evening_deduction),
        "commute_deduction_hours": round_hours(total),
    }


def calculate_day_pay_hours(start_dt, end_dt, jobs_data=None, engineer_name=""):
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

    activity = calculate_activity_hours(start_dt, end_dt, jobs_data, engineer_name)
    commute = calculate_commute_deductions(
        start_dt,
        end_dt,
        activity.get("productive_intervals", []),
    )

    payable_ooh_hours = max(0.0, ooh_hours - commute["commute_deduction_hours"])

    return {
        "total_hours": round_hours(total_hours),
        "normal_hours": round_hours(normal_hours),
        "ooh_hours": round_hours(ooh_hours),
        "break_deducted": round_hours(break_deducted),
        "morning_commute_deduction_hours": commute["morning_commute_deduction_hours"],
        "evening_commute_deduction_hours": commute["evening_commute_deduction_hours"],
        "commute_deduction_hours": commute["commute_deduction_hours"],
        "payable_ooh_hours": round_hours(payable_ooh_hours),
        "productive_hours": activity["productive_hours"],
        "inactive_hours": activity["inactive_hours"],
        "utilisation_percent": activity["utilisation_percent"],
    }


def build_pay_summary(start_dt, end_dt, hours):
    if not hours:
        return "Unable to calculate hours - start or end time missing."

    return (
        f"Start: {start_dt.strftime('%d/%m/%Y %H:%M')}\n"
        f"End: {end_dt.strftime('%d/%m/%Y %H:%M')}\n"
        f"Total hours: {hours['total_hours']}\n"
        f"Normal hours: {hours['normal_hours']}\n"
        f"OOH before commute deduction: {hours['ooh_hours']} at 1.5x\n"
        f"Morning commute deduction: {hours['morning_commute_deduction_hours']}\n"
        f"Evening commute deduction: {hours['evening_commute_deduction_hours']}\n"
        f"Total commute deduction: {hours['commute_deduction_hours']}\n"
        f"Payable OOH hours: {hours['payable_ooh_hours']}\n"
        f"Break deducted: {hours['break_deducted']}\n"
        f"Productive hours: {hours['productive_hours']}\n"
        f"Inactive/unproductive hours: {hours['inactive_hours']}\n"
        f"Utilisation: {hours['utilisation_percent']}%"
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
    """Return only active jobs that should block an engineer ending today.

    Returned jobs such as Revisit Required / No Access should not block End Day
    once they are back with Helpdesk/Awaiting Deployment. If Helpdesk later
    reassigns the job for a future date, it should also not block today's End Day.
    """
    today = datetime.now(UK_TZ).date()
    open_jobs = []
    active_statuses = {ASSIGNED_STATUS, TRAVELLING_STATUS, ON_SITE_STATUS}

    for job in jobs_data:
        fields = job.get("fields", {})
        job_date = sharepoint_date_to_uk_date(fields.get("Date", ""))
        assigned_ids = get_assigned_engineer_ids(fields)

        if str(engineer_lookup_id) not in assigned_ids or job_date != today:
            continue

        status = str(fields.get("Status", "") or "").strip()
        outcome = str(fields.get("JobOutcome", "") or "").strip()

        # Completed is never an open day blocker.
        if status == COMPLETED_STATUS or outcome == "Completed":
            continue

        # Only live/active engineer-owned statuses block End Day. Jobs returned
        # to Helpdesk/Awaiting Deployment are office-owned, even if SharePoint
        # still carries old same-day visit history.
        if status in active_statuses:
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
    """Return assigned engineer lookup IDs from any SharePoint field shape available."""
    assigned = []
    seen = set()

    def add_lookup_id(value):
        value = str(value or "").strip()
        if not value or value.lower() in ["none", "null"] or value in seen:
            return
        assigned.append(value)
        seen.add(value)

    engineer_values = fields.get("Engineer", [])
    if isinstance(engineer_values, list):
        for engineer in engineer_values:
            if isinstance(engineer, dict):
                add_lookup_id(engineer.get("LookupId"))
            else:
                add_lookup_id(engineer)

    # Depending on Graph expansion, the writable lookup ID field may also be present.
    for key in ["EngineerLookupId", "Engineer Lookup Id", "EngineerId", "Engineer ID"]:
        value = fields.get(key)
        if isinstance(value, list):
            for item in value:
                add_lookup_id(item)
        else:
            add_lookup_id(value)

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
        [InlineKeyboardButton("✅ Send / submit worksheet", callback_data="review|submit")],
        [InlineKeyboardButton("✏️ Change outcome", callback_data="review|edit_outcome")],
        [InlineKeyboardButton("⏱ Edit travel / on site / off site times", callback_data="review|edit_times")],
        [InlineKeyboardButton("📝 Edit work completed", callback_data="review|edit_work")],
        [InlineKeyboardButton("🧰 Edit materials", callback_data="review|edit_materials")],
        [InlineKeyboardButton("🔁 Edit follow-on", callback_data="review|edit_follow_on")],
        [InlineKeyboardButton("📷 Add more photos", callback_data="review|add_photos")],
        [InlineKeyboardButton("✍️ Change signature required", callback_data="review|edit_signature")],
        [InlineKeyboardButton("🔄 Restart worksheet", callback_data="review|restart")],
        [InlineKeyboardButton("❌ Cancel", callback_data="review|cancel")],
    ])


def get_outcome_edit_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Completed", callback_data="review_outcome|Completed")],
        [InlineKeyboardButton("🔁 Revisit Required", callback_data="review_outcome|Revisit Required")],
        [InlineKeyboardButton("🚫 No Access", callback_data="review_outcome|No Access")],
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


def get_job_buttons(item_id, maps_query=None):
    rows = [
        [
            InlineKeyboardButton("🚗 Start Travelling", callback_data=f"status|{item_id}|Travelling"),
        ],
        [
            InlineKeyboardButton("📍 Arrived On Site", callback_data=f"status|{item_id}|On Site"),
        ],
        [
            InlineKeyboardButton("✅ Complete Job", callback_data=f"confirm_outcome|{item_id}|Completed"),
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
        [
            InlineKeyboardButton("🏷 Create Site Asset", callback_data=f"asset_create_job|{item_id}"),
        ],
    ]

    if maps_query:
        maps_url = (
            "https://www.google.com/maps/dir/?api=1&destination="
            + quote_plus(str(maps_query))
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



def get_sharepoint_rest_headers():
    """Headers for SharePoint REST API calls.

    Graph calls use a Graph token. Folder colouring is a SharePoint REST
    endpoint, so this requests a SharePoint audience token instead.
    """
    site_hostname = SHAREPOINT_SITE.split("/")[2]
    token_result = msal_app.acquire_token_for_client(
        scopes=[f"https://{site_hostname}/.default"]
    )

    if "access_token" not in token_result:
        raise Exception(f"Could not get SharePoint REST token: {token_result}")

    return {
        "Authorization": f"Bearer {token_result['access_token']}",
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
    }


def get_drive_web_url(site_id, drive_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    response = requests.get(url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(f"Could not get SharePoint drives: {response.text}")

    for drive in response.json().get("value", []):
        if drive.get("name") == drive_name:
            return drive.get("webUrl", "")

    raise Exception(f"Document library not found: {drive_name}")




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


def get_current_visit_times_from_worksheet(worksheet):
    """Return travel/on-site/off-site times for the current engineer visit."""
    edited = worksheet.get("EditedVisitTimes") or {}
    times = {
        "travel": edited.get("travel", ""),
        "on_site": edited.get("on_site", ""),
        "off_site": edited.get("off_site", ""),
    }

    if all(times.values()):
        return times

    fields = worksheet.get("fields", {}) or {}
    log_text = get_field_value(fields, "EngineerVisitLog", "Engineer Visit Log") or ""
    engineer_name = str(worksheet.get("engineer_name", "") or "").strip().lower()

    try:
        visits = parse_engineer_visit_log(log_text)
        matching = [
            visit for visit in visits
            if str(visit.get("engineer", "") or "").strip().lower() == engineer_name
        ]
        if matching:
            latest = matching[-1]
            times["travel"] = times["travel"] or latest.get("travel", "")
            times["on_site"] = times["on_site"] or latest.get("on_site", "")
            times["off_site"] = times["off_site"] or latest.get("off_site", "")
    except Exception as e:
        print(f"WARNING: Could not read current visit times for final review: {e}")

    times["off_site"] = times["off_site"] or datetime.now(UK_TZ).strftime("%H:%M")
    return times


def build_review_text(worksheet):
    signature_required = "Yes" if worksheet.get("ClientSignatureRequired") else "No"
    outcome = worksheet.get("JobOutcome", "Completed")
    no_access_reason = worksheet.get("NoAccessReason", "")
    signature_received = "Yes" if worksheet.get("ClientSignatureReceived") else "No"
    no_access_line = f"No Access reason: {no_access_reason}\n" if outcome == "No Access" and no_access_reason else ""
    times = get_current_visit_times_from_worksheet(worksheet)

    return (
        f"Final check before sending worksheet for {worksheet['cdr_number']}:\n\n"
        f"Outcome: {outcome}\n"
        f"{no_access_line}"
        f"Travel time: {times.get('travel') or 'Missing'}\n"
        f"On site time: {times.get('on_site') or 'Missing'}\n"
        f"Off site time: {times.get('off_site') or 'Missing'}\n\n"
        f"Work completed:\n{worksheet.get('WorkCompleted', '')}\n\n"
        f"Materials used:\n{worksheet.get('MaterialsUsed', '')}\n\n"
        f"Follow-on required: {'Yes' if worksheet.get('FollowOnRequired') else 'No'}\n"
        f"Follow-on notes: {worksheet.get('FollowOnNotes', '') or 'None'}\n\n"
        f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n"
        f"Client signature required: {signature_required}\n"
        f"Client signature received: {signature_received}\n\n"
        "If anything is wrong, edit it below before pressing Send."
    )


def format_time_for_group(value):
    return str(value or "").strip() or "N/A"


def yes_no(value):
    return "Yes" if value else "No"


def build_trade_group_text_summary(worksheet, fields, updated_log, outcome):
    """Build a WhatsApp-style text update for the relevant trade group.

    This is deliberately separate from worksheet PDF generation. It only uses
    fields the engineer actually interacts with during the job/worksheet flow.
    """
    cdr_number = worksheet.get("cdr_number") or fields.get("CDRNumber", "")
    engineer_name = worksheet.get("engineer_name", "")
    date_to_attend = format_sharepoint_date(fields.get("Date", ""))
    start_time = fields.get("StartTime", "") or fields.get("Start Time", "")
    site_name = fields.get("SiteName", "") or ""
    task = fields.get("Task", "") or ""
    category = get_field_value(fields, "JobCategory", "Job Category") or ""

    travel_time = ""
    on_site_time = ""
    off_site_time = ""

    try:
        visits = parse_engineer_visit_log(updated_log)
        matching_visits = [
            visit for visit in visits
            if str(visit.get("engineer", "")).strip().lower() == str(engineer_name).strip().lower()
        ]
        if matching_visits:
            visit = matching_visits[-1]
            travel_time = visit.get("travel", "")
            on_site_time = visit.get("on_site", "")
            off_site_time = visit.get("off_site", "")
    except Exception as e:
        print(f"WARNING: Could not parse visit times for electrical group summary: {e}")

    date_line = date_to_attend
    if start_time:
        date_line = f"{date_line} {start_time}".strip()

    lines = [
        "🔧 JOB UPDATE",
        "",
        f"Date To Attend: {date_line}",
        f"Job Location: {site_name}",
        f"Job Number: {cdr_number}",
        f"Engineer: {engineer_name}",
        f"Description: {task}",
    ]

    if category:
        lines.append(f"Job Type: {category}")

    lines.extend([
        f"Travel: {format_time_for_group(travel_time)}",
        f"On Site: {format_time_for_group(on_site_time)}",
        f"Off Site: {format_time_for_group(off_site_time)}",
        f"Photos Taken?: {yes_no(len(worksheet.get('photo_links', [])) > 0)}",
        f"Outcome: {outcome}",
    ])

    if outcome == "No Access" and worksheet.get("NoAccessReason"):
        lines.append(f"No Access Reason: {worksheet.get('NoAccessReason')}")

    work_completed = str(worksheet.get("WorkCompleted", "") or "").strip()
    if work_completed and work_completed.lower() not in ["n/a", "na", "none", "no"]:
        lines.extend(["", f"Work Completed / Comments: {work_completed}"])

    materials_used = str(worksheet.get("MaterialsUsed", "") or "").strip()
    if materials_used and materials_used.lower() not in ["n/a", "na", "none", "no"]:
        lines.extend(["", f"Materials Used: {materials_used}"])

    if worksheet.get("FollowOnRequired"):
        follow_on_notes = str(worksheet.get("FollowOnNotes", "") or "").strip()
        lines.extend(["", "Follow-on Required?: Yes"])
        if follow_on_notes and follow_on_notes.lower() not in ["n/a", "na", "none", "no"]:
            lines.append(follow_on_notes)
    else:
        lines.extend(["", "Follow-on Required?: No"])

    return "\n".join(lines)


def get_trade_group_chat_id(fields):
    """Route job outcome summaries to the correct trade Telegram group."""
    category = str(get_field_value(fields, "JobCategory", "Job Category") or "").strip().lower()

    electrical_categories = {"electrical", "fire"}
    mechanical_categories = {"mechanical", "plumbing", "hvac", "catering"}

    if category in electrical_categories:
        return CDR_ELECTRICAL_CHAT_ID, "CDR Electrical"

    if category in mechanical_categories:
        return CDR_MECHANICAL_CHAT_ID, "CDR Mechanical"

    return None, ""


async def notify_trade_group(context, worksheet, fields, updated_log, outcome):
    """Send the trade group text summary and uploaded job photos.

    This does not generate or send worksheets/PDFs. Photos are sent as Telegram
    media albums using the actual photo bytes captured during the worksheet flow.
    """
    chat_id, group_name = get_trade_group_chat_id(fields)

    if not chat_id:
        return

    try:
        message = build_trade_group_text_summary(worksheet, fields, updated_log, outcome)
        await context.bot.send_message(
            chat_id=chat_id,
            text=message,
        )
        await send_trade_group_photos(context, chat_id, worksheet)
    except Exception as e:
        print(f"WARNING: Could not send trade group summary/photos to {group_name}: {e}")


async def send_trade_group_photos(context, chat_id, worksheet):
    """Send every engineer photo to the trade group using max-size Telegram albums.

    Telegram media groups can contain a maximum of 10 photos. This function sends
    all available photos in full 10-photo albums, then sends the final remainder.
    If a media group times out or fails, it falls back to individual photo sends
    for that batch so one bad upload does not stop the rest of the job photos.
    """
    photo_items = worksheet.get("photo_files_for_group", []) or []

    if not photo_items:
        print("Trade group photo send skipped: worksheet contains no photo bytes.")
        return

    total_photos = len(photo_items)
    sent_total = 0
    failed_total = 0
    invalid_total = 0

    async def send_single_photo(file_bytes, file_name):
        single_buffer = BytesIO(file_bytes)
        single_buffer.name = file_name
        await context.bot.send_photo(
            chat_id=chat_id,
            photo=single_buffer,
            connect_timeout=30,
            read_timeout=90,
            write_timeout=120,
            pool_timeout=30,
        )

    valid_items = []
    for index, item in enumerate(photo_items):
        if not isinstance(item, dict):
            invalid_total += 1
            continue

        file_bytes = item.get("bytes")
        file_name = item.get("file_name") or f"job_photo_{index + 1}.jpg"

        if not file_bytes:
            invalid_total += 1
            continue

        valid_items.append((file_bytes, file_name))

    if not valid_items:
        print("WARNING: Trade group photos had no valid photo bytes to send.")
        return

    try:
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"📷 Sending {len(valid_items)} job photo(s) to this group.",
        )
    except Exception as e:
        print(f"WARNING: Could not send trade group photo count notice: {e}")

    # Max out Telegram albums at 10 photos per group: 24 photos = 10 + 10 + 4.
    for batch_start in range(0, len(valid_items), 10):
        batch = valid_items[batch_start:batch_start + 10]
        batch_number = (batch_start // 10) + 1
        batch_end = batch_start + len(batch)

        media = []
        for file_bytes, file_name in batch:
            album_buffer = BytesIO(file_bytes)
            album_buffer.name = file_name
            media.append(InputMediaPhoto(media=album_buffer))

        try:
            if len(media) == 1:
                await send_single_photo(batch[0][0], batch[0][1])
            else:
                await context.bot.send_media_group(
                    chat_id=chat_id,
                    media=media,
                    connect_timeout=30,
                    read_timeout=180,
                    write_timeout=240,
                    pool_timeout=30,
                )

            sent_total += len(batch)
            print(f"Trade group photo batch {batch_number} sent photos {batch_start + 1}-{batch_end}/{len(valid_items)}.")
            await asyncio.sleep(1.0)

        except Exception as album_error:
            print(
                f"WARNING: Trade group photo album batch {batch_number} failed "
                f"for photos {batch_start + 1}-{batch_end}; sending individually: {album_error}"
            )

            for file_bytes, file_name in batch:
                try:
                    await send_single_photo(file_bytes, file_name)
                    sent_total += 1
                    await asyncio.sleep(0.7)
                except Exception as single_error:
                    failed_total += 1
                    print(f"WARNING: Could not send trade group photo {file_name}: {single_error}")

    print(
        f"Trade group photos sent {sent_total}/{len(valid_items)} valid photos "
        f"from {total_photos} uploaded. Invalid: {invalid_total}. Failed: {failed_total}."
    )

async def notify_helpdesk(context, text):
    if HELPDESK_CHAT_ID:
        try:
            await context.bot.send_message(chat_id=HELPDESK_CHAT_ID, text=text)
        except Exception as e:
            print(f"Could not notify helpdesk: {e}")


@web_app.get("/", response_class=HTMLResponse)
def bot_health_page():
    return HTMLResponse("CDR Engineer Bot signature portal is online.")


@web_app.get("/sign/{cdr_number}", response_class=HTMLResponse)
def signature_page(cdr_number: str, token: str):
    site_id, jobs_list_id, job = get_job_by_cdr_and_token(cdr_number, token)

    if not job:
        return HTMLResponse("Invalid or expired signature link.", status_code=404)

    fields = job["fields"]

    if bool_field(fields.get("ClientSignatureReceived")):
        return HTMLResponse("This job has already been signed.", status_code=200)

    site = fields.get("SiteName", "")
    address = ""
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
        

:root {{
    --bg: #070b12;
    --panel: rgba(17, 24, 39, 0.82);
    --panel-solid: #111827;
    --panel-soft: rgba(255,255,255,0.055);
    --border: rgba(255,255,255,0.10);
    --border-strong: rgba(245,130,32,0.34);
    --text: #f8fafc;
    --muted: #94a3b8;
    --muted-2: #64748b;
    --orange: #f58220;
    --orange-soft: rgba(245,130,32,0.16);
    --green: #22c55e;
    --blue: #38bdf8;
    --amber: #f59e0b;
    --red: #ef4444;
    --grey: #64748b;
    --shadow: 0 24px 70px rgba(0,0,0,0.36);
    --radius-xl: 24px;
    --radius-lg: 18px;
}}

* {{ box-sizing: border-box; }}

body {{
    margin: 0;
    min-height: 100vh;
    color: var(--text);
    font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
    background:
        radial-gradient(circle at top left, rgba(245,130,32,0.18), transparent 30%),
        radial-gradient(circle at top right, rgba(56,189,248,0.10), transparent 26%),
        linear-gradient(135deg, #05070b 0%, #0b1020 46%, #111827 100%);
}}

body::before {{
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    background-image:
        linear-gradient(rgba(255,255,255,0.032) 1px, transparent 1px),
        linear-gradient(90deg, rgba(255,255,255,0.032) 1px, transparent 1px);
    background-size: 42px 42px;
    mask-image: linear-gradient(to bottom, rgba(0,0,0,.6), transparent 70%);
}}

.container, .wrap, main {{
    width: min(1480px, calc(100% - 32px));
    margin: 0 auto;
}}

.header, .topbar {{
    position: sticky;
    top: 0;
    z-index: 20;
    backdrop-filter: blur(18px);
    background: rgba(7, 11, 18, 0.78);
    border-bottom: 1px solid var(--border);
}}

.header-inner, .topbar-inner {{
    width: min(1480px, calc(100% - 32px));
    margin: 0 auto;
    padding: 18px 0;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
}}

.brand, .brand-wrap, .title-wrap {{
    display: flex;
    align-items: center;
    gap: 14px;
}}

.brand img, .brand-logo, img[alt="CDR M&E Services Ltd"] {{
    height: 54px !important;
    width: auto !important;
    max-width: 230px !important;
    object-fit: contain !important;
    display: block !important;
    background: rgba(255,255,255,0.96);
    padding: 7px 10px;
    border-radius: 14px;
    box-shadow: 0 16px 35px rgba(0,0,0,0.28);
}}

h1 {{
    margin: 0;
    font-size: clamp(24px, 2.8vw, 38px);
    letter-spacing: -0.04em;
    line-height: 1.05;
}}

.subtitle, .small, .muted {{
    color: var(--muted);
}}

.nav, .tabs, .dashboard-nav {{
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin: 18px 0 24px;
    padding: 8px;
    border: 1px solid var(--border);
    background: rgba(255,255,255,0.055);
    border-radius: 18px;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.06);
}}

.nav a, .tabs a, .dashboard-nav a {{
    color: var(--muted);
    text-decoration: none;
    padding: 11px 15px;
    border-radius: 13px;
    border: 1px solid transparent;
    transition: all .18s ease;
    font-weight: 700;
    font-size: 14px;
}}

.nav a:hover, .tabs a:hover, .dashboard-nav a:hover {{
    color: var(--text);
    background: rgba(255,255,255,0.08);
    border-color: var(--border);
    transform: translateY(-1px);
}}

.nav a.active, .tabs a.active, .dashboard-nav a.active {{
    color: #fff;
    background: linear-gradient(135deg, rgba(245,130,32,0.95), rgba(255,168,76,0.78));
    border-color: rgba(255,255,255,0.18);
    box-shadow: 0 12px 28px rgba(245,130,32,0.22);
}}

.metrics, .stats, .summary-grid {{
    display: grid;
    grid-template-columns: repeat(6, minmax(130px, 1fr));
    gap: 14px;
    margin: 18px 0 24px;
}}

.metric, .stat, .summary-card {{
    background: linear-gradient(180deg, rgba(255,255,255,0.085), rgba(255,255,255,0.04));
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 16px;
    box-shadow: 0 14px 40px rgba(0,0,0,0.18);
}}

.metric .label, .stat .label, .summary-card .label {{
    color: var(--muted);
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .08em;
    font-weight: 800;
}}

.metric .value, .stat .value, .summary-card .value {{
    font-size: clamp(24px, 2.4vw, 36px);
    line-height: 1;
    font-weight: 900;
    margin-top: 9px;
    letter-spacing: -0.04em;
}}

.grid, .cards, .engineer-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(305px, 1fr));
    gap: 18px;
}}

.card, .engineer-card, .job-card, .panel {{
    position: relative;
    overflow: hidden;
    background: linear-gradient(180deg, rgba(17,24,39,0.92), rgba(15,23,42,0.72));
    border: 1px solid var(--border);
    border-radius: var(--radius-xl);
    padding: 20px;
    box-shadow: var(--shadow);
}}

.card::before, .engineer-card::before, .job-card::before, .panel::before {{
    content: "";
    position: absolute;
    inset: 0 0 auto 0;
    height: 3px;
    background: linear-gradient(90deg, var(--orange), rgba(56,189,248,.75), transparent);
    opacity: .85;
}}

.card:hover, .engineer-card:hover, .job-card:hover {{
    transform: translateY(-2px);
    transition: transform .18s ease, border-color .18s ease;
    border-color: rgba(255,255,255,0.18);
}}

.avatar, .initials {{
    display: inline-grid;
    place-items: center;
    width: 46px;
    height: 46px;
    border-radius: 16px;
    background: linear-gradient(135deg, var(--orange), #ffb15e);
    color: #111827;
    font-weight: 950;
    box-shadow: 0 14px 30px rgba(245,130,32,.24);
}}

.status, .pill, .badge {{
    display: inline-flex;
    align-items: center;
    gap: 7px;
    padding: 7px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 900;
    letter-spacing: .01em;
    border: 1px solid rgba(255,255,255,0.12);
    background: rgba(255,255,255,0.065);
}}

.status.on-site, .pill.green, .badge.green {{ color: #86efac; background: rgba(34,197,94,.13); border-color: rgba(34,197,94,.24); }}
.status.travelling, .pill.blue, .badge.blue {{ color: #7dd3fc; background: rgba(56,189,248,.12); border-color: rgba(56,189,248,.24); }}
.status.idle, .pill.amber, .badge.amber {{ color: #fcd34d; background: rgba(245,158,11,.13); border-color: rgba(245,158,11,.26); }}
.status.offline, .pill.grey, .badge.grey {{ color: #cbd5e1; background: rgba(100,116,139,.16); border-color: rgba(100,116,139,.25); }}
.status.alert, .pill.red, .badge.red {{ color: #fca5a5; background: rgba(239,68,68,.13); border-color: rgba(239,68,68,.27); }}

.card-title, .engineer-name, h2, h3 {{
    letter-spacing: -0.03em;
}}

.card-title, .engineer-name {{
    font-size: 20px;
    font-weight: 950;
    margin: 0;
}}

.meta-row, .row, .line {{
    display: flex;
    justify-content: space-between;
    gap: 14px;
    padding: 9px 0;
    border-bottom: 1px solid rgba(255,255,255,0.07);
}}

.meta-row:last-child, .row:last-child, .line:last-child {{
    border-bottom: none;
}}

.meta-row span:first-child, .row span:first-child, .line span:first-child {{
    color: var(--muted);
}}

.meta-row span:last-child, .row span:last-child, .line span:last-child {{
    text-align: right;
    font-weight: 800;
}}

.progress, .util-bar {{
    height: 9px;
    width: 100%;
    border-radius: 999px;
    background: rgba(255,255,255,0.085);
    overflow: hidden;
    margin-top: 10px;
}}

.progress > div, .util-fill {{
    height: 100%;
    border-radius: inherit;
    background: linear-gradient(90deg, var(--orange), #ffd19a);
}}

table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0 8px;
}}

th {{
    color: var(--muted);
    text-align: left;
    padding: 10px 12px;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .08em;
}}

td {{
    padding: 13px 12px;
    background: rgba(255,255,255,0.055);
    border-top: 1px solid var(--border);
    border-bottom: 1px solid var(--border);
}}

td:first-child {{
    border-left: 1px solid var(--border);
    border-radius: 14px 0 0 14px;
}}

td:last-child {{
    border-right: 1px solid var(--border);
    border-radius: 0 14px 14px 0;
}}

.footer, .refresh {{
    color: var(--muted-2);
    font-size: 12px;
    margin: 26px 0;
}}

@media (max-width: 960px) {{
    .metrics, .stats, .summary-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    .header-inner, .topbar-inner {{ flex-direction: column; align-items: flex-start; }}
}}

@media (max-width: 620px) {{
    .container, .wrap, main, .header-inner, .topbar-inner {{ width: min(100% - 20px, 1480px); }}
    .metrics, .stats, .summary-grid {{ grid-template-columns: 1fr; }}
    .grid, .cards, .engineer-grid {{ grid-template-columns: 1fr; }}
    .nav, .tabs, .dashboard-nav {{ overflow-x: auto; flex-wrap: nowrap; }}
    .nav a, .tabs a, .dashboard-nav a {{ white-space: nowrap; }}
}}

/* Signature page hard overrides. Dashboard CSS is shared in this file,
   so these rules must be last to stop dark dashboard text colours leaking
   into the white client signature form. */
body.signature-page {{
    background: #f4f4f4 !important;
    color: #111827 !important;
    font-family: Arial, sans-serif !important;
    padding: 20px !important;
}}
body.signature-page .signature-container {{
    max-width: 650px !important;
    width: min(650px, calc(100% - 24px)) !important;
    margin: auto !important;
    background: #ffffff !important;
    color: #111827 !important;
    padding: 25px !important;
    border-radius: 12px !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.12) !important;
}}
body.signature-page h2,
body.signature-page p,
body.signature-page label,
body.signature-page strong,
body.signature-page .small {{
    color: #111827 !important;
}}
body.signature-page .job-box {{
    background: #f3f4f6 !important;
    color: #111827 !important;
    border: 1px solid #e5e7eb !important;
}}
body.signature-page input[type="text"] {{
    background: #ffffff !important;
    color: #111827 !important;
    border: 1px solid #6b7280 !important;
}}
body.signature-page canvas {{
    background: #ffffff !important;
    border: 2px solid #111827 !important;
}}

</style>
    
<meta name="theme-color" content="#070b12">
<meta name="apple-mobile-web-app-capable" content="yes">
</head>
    <body class="signature-page">
        <div class="container signature-container">
            <img src="/logo.png" alt="CDR M&E Services Ltd" style="display:block; max-width:320px; width:80%; margin:0 auto 20px auto;">
            <h2 style="text-align:center;">Client Signature</h2>
            <div class="job-box">
                <p><strong>CDR Number:</strong> {cdr_number}</p>
                <p><strong>Site:</strong> {site}</p>
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
                velocityFilterWeight: 0.7,
                penColor: "#111827",
                backgroundColor: "#ffffff"
            }});

            // Stop mobile browsers treating signature movement as page scroll/swipe.
            // The passive:false option is important on iPhone/Android.
            ["touchstart", "touchmove", "touchend", "pointerdown", "pointermove", "pointerup"].forEach(function(eventName) {{
                canvas.addEventListener(eventName, function(event) {{
                    event.preventDefault();
                    event.stopPropagation();
                }}, {{ passive: false }});
            }});

            // Desktop reliability: make sure a mouse/pointer release is always
            // received, even if the cursor moves outside the canvas. Without
            // this, some desktop browsers can leave the pad in drawing mode.
            canvas.addEventListener("pointerdown", function(event) {{
                try {{ canvas.setPointerCapture(event.pointerId); }} catch (e) {{}}
            }}, {{ passive: false }});

            function stopSignatureDrawing(event) {{
                try {{
                    if (event && event.pointerId !== undefined) {{
                        try {{ canvas.releasePointerCapture(event.pointerId); }} catch (e) {{}}
                    }}
                    if (signaturePad && signaturePad._isDrawing) {{
                        if (typeof signaturePad._strokeEnd === "function") {{
                            signaturePad._strokeEnd(event || new Event("pointerup"));
                        }}
                        signaturePad._isDrawing = false;
                    }}
                }} catch (e) {{
                    try {{ signaturePad._isDrawing = false; }} catch (ignored) {{}}
                }}
            }}

            canvas.addEventListener("pointerup", stopSignatureDrawing, {{ passive: false }});
            canvas.addEventListener("pointercancel", stopSignatureDrawing, {{ passive: false }});
            canvas.addEventListener("mouseleave", stopSignatureDrawing, {{ passive: false }});
            window.addEventListener("pointerup", stopSignatureDrawing, {{ passive: false }});
            window.addEventListener("mouseup", stopSignatureDrawing, {{ passive: false }});

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



async def my_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return

    chat = update.effective_chat
    user = update.effective_user

    await update.message.reply_text(
        f"📌 Chat ID:\n{chat.id}\n\n"
        f"💬 Chat Type:\n{chat.type}\n\n"
        f"👤 User:\n{user.full_name}"
    )



async def group_chat_cleanup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Remove any old bot keyboard from group chats and stop group workflows."""
    if not update.effective_chat or update.effective_chat.type == "private":
        return ConversationHandler.END

    if update.message:
        try:
            await update.message.reply_text(
                "CDR bot controls only work in private chat.",
                reply_markup=ReplyKeyboardRemove(),
            )
        except Exception as e:
            print(f"WARNING: Could not remove group keyboard: {e}")

    return ConversationHandler.END


def run_signature_web_server():
    uvicorn.run(web_app, host="0.0.0.0", port=PORT, log_level="info")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return

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
    if is_group_chat(update):
        return ConversationHandler.END

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

        role = get_bot_user_role(site_id, user_id)

        context.user_data["start_day"] = {
            "role": role,
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
    if is_group_chat(update):
        return ConversationHandler.END

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

    start_day = context.user_data.get("start_day")
    if start_day and is_vehicle_check_exempt_role(start_day.get("role")):
        create_vehicle_exempt_start_day_log(start_day)
        role_label = str(start_day.get("role") or "User").title()
        context.user_data.pop("start_day", None)
        await update.message.reply_text(
            f"{role_label} day started. No van check, van registration or mileage is required. Your jobs are now unlocked.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    await update.message.reply_text("Starting your day. Please enter the van registration.")
    return START_DAY_VAN_REG


async def startday_confirm_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    answer = query.data.split("|", 1)[1]

    if answer == "no":
        context.user_data.pop("start_day", None)
        await query.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    start_day = context.user_data.get("start_day")
    if start_day and is_vehicle_check_exempt_role(start_day.get("role")):
        create_vehicle_exempt_start_day_log(start_day)
        role_label = str(start_day.get("role") or "User").title()
        context.user_data.pop("start_day", None)
        await query.message.reply_text(
            f"{role_label} day started. No van check, van registration or mileage is required. Your jobs are now unlocked.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    await query.message.reply_text("Starting your day. Please enter the van registration.")
    return START_DAY_VAN_REG


async def startday_van_reg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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


def is_apprentice_role(role):
    return str(role or "").strip().lower() == "apprentice"



def create_vehicle_exempt_start_day_log(start_day):
    """Start a day without van registration, mileage, van check or van photos."""
    start_day["van_reg"] = ""
    start_day["start_mileage"] = ""
    start_day["van_check_answers"] = []
    start_day["van_photo_links"] = []
    create_start_day_log(start_day, van_check_completed=False)


def create_apprentice_start_day_log(start_day):
    """Backward-compatible wrapper for older apprentice start-day call sites."""
    create_vehicle_exempt_start_day_log(start_day)


async def close_end_day_record(context, end_day, mileage=None, include_mileage=True):
    """Close the active day log. Apprentices can close without mileage prompts."""
    # Re-check assigned jobs before closing the day in case one was added during the end-day flow.
    site_id, _, _, engineers, jobs_data = get_sharepoint_data()
    open_jobs = get_open_jobs_for_engineer_today(jobs_data, end_day["engineer_lookup_id"])

    if open_jobs:
        return False, (
            "Your day has not been ended because you still have job(s) assigned for today. "
            "Complete them, mark No Access, or mark Revisit Required first.\n\n"
            f"Open job(s):\n{format_open_jobs_for_end_day(open_jobs)}"
        ), None

    end_time = datetime.now(UK_TZ)
    start_time_value = get_field_value(
        end_day.get("day_log_fields", {}),
        "StartTime",
        "Start Time",
    )
    start_time = parse_sharepoint_datetime(start_time_value)
    hours = calculate_day_pay_hours(
        start_time,
        end_time,
        jobs_data=jobs_data,
        engineer_name=end_day.get("engineer_name", ""),
    )
    pay_summary = build_pay_summary(start_time, end_time, hours)

    update_payload = {
        "End Time": end_time.isoformat(),
        "EndTime": end_time.isoformat(),
        "Status": DAY_CLOSED_STATUS,
        "Pay Summary": pay_summary,
        "PaySummary": pay_summary,
    }

    if include_mileage:
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

        update_payload.update({
            "End Mileage": mileage,
            "EndMileage": mileage,
            "Total Mileage": total_mileage if total_mileage is not None else "",
            "TotalMileage": total_mileage if total_mileage is not None else "",
        })

    if hours:
        update_payload.update({
            "Total Hours": hours["total_hours"],
            "TotalHours": hours["total_hours"],
            "Normal Hours": hours["normal_hours"],
            "NormalHours": hours["normal_hours"],
            "OOH Hours": hours["ooh_hours"],
            "OOHHours": hours["ooh_hours"],
            "Payable OOH Hours": hours["payable_ooh_hours"],
            "PayableOOHHours": hours["payable_ooh_hours"],
            "Commute Deduction Hours": hours["commute_deduction_hours"],
            "CommuteDeductionHours": hours["commute_deduction_hours"],
            "Morning Commute Deduction Hours": hours["morning_commute_deduction_hours"],
            "MorningCommuteDeductionHours": hours["morning_commute_deduction_hours"],
            "Evening Commute Deduction Hours": hours["evening_commute_deduction_hours"],
            "EveningCommuteDeductionHours": hours["evening_commute_deduction_hours"],
            "Productive Hours": hours["productive_hours"],
            "ProductiveHours": hours["productive_hours"],
            "Inactive Hours": hours["inactive_hours"],
            "InactiveHours": hours["inactive_hours"],
            "Unproductive Hours": hours["inactive_hours"],
            "UnproductiveHours": hours["inactive_hours"],
            "Utilisation Percent": hours["utilisation_percent"],
            "UtilisationPercent": hours["utilisation_percent"],
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
    return True, "Day ended. Your job buttons are now locked until you start your next day.", pay_summary


async def startday_start_mileage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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

        role = get_bot_user_role(site_id, user_id)

        context.user_data["end_day"] = {
            "role": role,
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
    if is_group_chat(update):
        return ConversationHandler.END

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

    end_day = context.user_data.get("end_day")
    if end_day and is_vehicle_check_exempt_role(end_day.get("role")):
        ok, message, pay_summary = await close_end_day_record(context, end_day, include_mileage=False)
        await update.message.reply_text(
            message + (f"\n\n{pay_summary}" if ok and pay_summary else ""),
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "Please enter your end mileage as a number.\n\n"
        "If you do not need to record mileage, type 0."
    )
    return END_DAY_MILEAGE


async def endday_confirm_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    answer = query.data.split("|", 1)[1]

    if answer == "no":
        context.user_data.pop("end_day", None)
        await query.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    end_day = context.user_data.get("end_day")
    if end_day and is_vehicle_check_exempt_role(end_day.get("role")):
        ok, message, pay_summary = await close_end_day_record(context, end_day, include_mileage=False)
        await query.message.reply_text(
            message + (f"\n\n{pay_summary}" if ok and pay_summary else ""),
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    await query.message.reply_text(
        "Please enter your end mileage as a number.\n\n"
        "If you do not need to record mileage, type 0."
    )
    return END_DAY_MILEAGE


async def endday_mileage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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

        ok, message, pay_summary = await close_end_day_record(
            context,
            end_day,
            mileage=mileage,
            include_mileage=True,
        )

        await update.message.reply_text(
            message + (f"\n\n{pay_summary}" if ok and pay_summary else ""),
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
    if is_group_chat(update):
        return ConversationHandler.END

    context.user_data.pop("end_day", None)
    await update.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def mystatus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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



def get_active_helpdesk_users(engineers):
    """Return active Helpdesk users from the Engineers list with Telegram IDs.

    Admin users are intentionally excluded unless their Role is set to Helpdesk.
    This keeps engineer job requests going to helpdesk phones only.
    """
    users = []
    seen_telegram_ids = set()

    for item in engineers:
        fields = item.get("fields", {})
        role = str(get_field_value(fields, "Role") or "").strip().lower()
        if role != "helpdesk":
            continue

        active_value = get_field_value(fields, "Active")
        if active_value not in [None, ""] and not bool_field(active_value):
            continue

        telegram_id = str(get_field_value(fields, "TelegramID", "Telegram ID") or "").strip()
        if not telegram_id or telegram_id in seen_telegram_ids:
            continue

        name = str(get_field_value(fields, "EngineerName", "Engineer Name", "Title") or "Helpdesk").strip()
        users.append({"name": name, "telegram_id": telegram_id})
        seen_telegram_ids.add(telegram_id)

    return users


async def request_job(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return

    """Engineer one-tap request for another job.

    This does not start a conversation and does not change any SharePoint job.
    It privately notifies every active Helpdesk user in the Engineers list.
    """
    try:
        user_id = str(update.effective_user.id)
        site_id, engineers_list_id, _, current_engineer = get_engineer_for_telegram_id(user_id)
        role = get_bot_user_role(site_id, user_id)

        if str(role).lower() not in ["engineer", "admin"]:
            await update.message.reply_text(
                "This option is for engineers to request another job.",
                reply_markup=get_main_menu(role),
            )
            return

        engineer_name = current_engineer["name"] if current_engineer else (update.effective_user.full_name or user_id)
        engineers = get_list_items(site_id, engineers_list_id)
        helpdesk_users = get_active_helpdesk_users(engineers)

        if not helpdesk_users:
            await update.message.reply_text(
                "No active Helpdesk users with Telegram IDs were found. Please ask the office to check the Engineers list Role, Active and TelegramID columns.",
                reply_markup=get_main_menu(role),
            )
            return

        message = (
            "📣 Job Request\n\n"
            f"Engineer: {engineer_name}\n"
            f"Time: {now_log_time()}\n\n"
            f"{engineer_name} is requesting another job / next instruction."
        )

        sent_count = 0
        failed = []

        for helpdesk_user in helpdesk_users:
            try:
                await context.bot.send_message(
                    chat_id=helpdesk_user["telegram_id"],
                    text=message,
                )
                sent_count += 1
            except Exception as send_error:
                failed.append(helpdesk_user["name"])
                print(f"WARNING: Could not send job request to {helpdesk_user['name']}: {send_error}")

        if sent_count == 0:
            await update.message.reply_text(
                "I found Helpdesk users, but could not send the request. Please check their Telegram IDs and that they have started the bot.",
                reply_markup=get_main_menu(role),
            )
            return

        reply = f"Request sent to {sent_count} Helpdesk user(s)."
        if failed:
            reply += f"\n\nNot sent to: {', '.join(failed)}"

        await update.message.reply_text(
            reply,
            reply_markup=get_main_menu(role),
        )

    except Exception as e:
        print(f"ERROR sending job request: {e}")
        await update.message.reply_text(
            "There was an error sending the request. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )



# =========================
# Asset Register / Portal integration
# =========================
ASSET_CATEGORY_OPTIONS = ["Mechanical", "Electrical", "HVAC", "Fire & Security", "Plumbing", "Building Fabric", "Company Asset", "Other"]
ASSET_TYPE_OPTIONS = ["Boiler", "DB Board", "AC Unit", "AHU", "Fan Coil", "Pump", "Fire Alarm Panel", "Emergency Lighting", "Water Heater", "Pressurisation Unit", "Extract Fan", "Company Asset", "Other"]


def asset_skip_keyboard():
    return ReplyKeyboardMarkup([["⏭️ Skip"], ["/cancel"]], resize_keyboard=True, one_time_keyboard=False)


def asset_category_keyboard():
    rows = [[x] for x in ASSET_CATEGORY_OPTIONS]
    rows.append(["/cancel"])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True, one_time_keyboard=False)


def asset_type_keyboard():
    rows = [ASSET_TYPE_OPTIONS[i:i + 2] for i in range(0, len(ASSET_TYPE_OPTIONS), 2)]
    rows.append(["/cancel"])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True, one_time_keyboard=False)


def asset_clean_answer(text):
    value = str(text or "").strip()
    return "" if value == "⏭️ Skip" else value


def portal_asset_headers():
    headers = {"Content-Type": "application/json"}
    if ASSET_API_KEY:
        headers["x-api-key"] = ASSET_API_KEY
    return headers


def create_asset_in_portal(payload):
    if not PORTAL_BASE_URL:
        raise RuntimeError("PORTAL_BASE_URL is not set in Railway for the bot.")
    response = requests.post(
        f"{PORTAL_BASE_URL}/api/assets",
        json=payload,
        headers=portal_asset_headers(),
        timeout=20,
    )
    try:
        data = response.json()
    except Exception:
        data = {"ok": False, "error": response.text[:500]}
    if not response.ok or not data.get("ok"):
        raise RuntimeError(data.get("error") or f"Portal returned HTTP {response.status_code}")
    return data


def asset_payload_from_context(context):
    data = context.user_data.get("asset_create", {}) or {}
    payload = {
        "created_by": data.get("created_by", "Telegram Bot"),
        "engineer": data.get("engineer", ""),
        "category": data.get("category", "Other"),
        "asset_type": data.get("asset_type", "Other"),
        "description": data.get("description", ""),
        "customer": data.get("customer", ""),
        "site_name": data.get("site_name", ""),
        "site_address": data.get("site_address", ""),
        "location": data.get("location", ""),
        "manufacturer": data.get("manufacturer", ""),
        "model": data.get("model", ""),
        "serial_number": data.get("serial_number", ""),
        "cdr_number": data.get("cdr_number", ""),
        "notes": data.get("notes", ""),
        "history_note": f"Asset created from Telegram bot" + (f" against {data.get('cdr_number')}" if data.get("cdr_number") else ""),
        "event_type": "Created",
    }
    if data.get("photo_file_id"):
        payload["photo_file_id"] = data.get("photo_file_id")
    if data.get("photo_data_uri"):
        payload["photo_data_uri"] = data.get("photo_data_uri")
    return payload


async def asset_start_manual(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END
    role = await get_role_for_update(update)
    context.user_data["asset_create"] = {"source": "manual", "created_by": f"Telegram Bot - {role}"}
    await update.message.reply_text(
        "Create Asset started. This will create a QR-linked asset in the portal.\n\nCustomer name?",
        reply_markup=asset_skip_keyboard(),
    )
    return ASSET_CUSTOMER


async def asset_start_from_job(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    try:
        item_id = query.data.split("|")[1]
        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)
        user_id = str(query.from_user.id)
        current_engineer = engineers_by_telegram.get(user_id)
        if not current_engineer:
            await query.message.reply_text("You are not set up as an engineer.")
            return ConversationHandler.END
        if not engineer_has_active_day(site_id, user_id):
            await query.message.reply_text("Please start your day first using 🟢 Start Day before creating a site asset from a job.")
            return ConversationHandler.END
        job = find_job_by_item_id(jobs_data, item_id)
        if not job:
            await query.message.reply_text("Could not find this job. Tap 📋 My Jobs and try again.")
            return ConversationHandler.END
        fields = job["fields"]
        assigned_ids = get_assigned_engineer_ids(fields)
        if current_engineer["lookup_id"] not in assigned_ids:
            await query.message.reply_text("You are not assigned to this job.")
            return ConversationHandler.END
        context.user_data["asset_create"] = {
            "source": "job",
            "job_item_id": item_id,
            "created_by": "Telegram Bot",
            "engineer": current_engineer.get("name", ""),
            "cdr_number": fields.get("CDRNumber", ""),
            "customer": fields.get("CustomerName", "") or fields.get("Customer", ""),
            "site_name": fields.get("SiteName", ""),
            "site_address": fields.get("Address", "") or fields.get("SiteAddress", ""),
            "notes": f"Created from job {fields.get('CDRNumber','')}.",
        }
        await query.message.reply_text(
            "Create Site Asset started. I have pulled the job/site details through.\n\nChoose the asset category:",
            reply_markup=asset_category_keyboard(),
        )
        return ASSET_CATEGORY
    except Exception as e:
        print(f"ERROR starting asset from job: {e}")
        await query.message.reply_text("Could not start asset creation. Please check Railway logs.")
        return ConversationHandler.END


async def asset_customer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["customer"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Site name?", reply_markup=asset_skip_keyboard())
    return ASSET_SITE_NAME


async def asset_site_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["site_name"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Site address?", reply_markup=asset_skip_keyboard())
    return ASSET_SITE_ADDRESS


async def asset_site_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["site_address"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Choose the asset category:", reply_markup=asset_category_keyboard())
    return ASSET_CATEGORY


async def asset_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["category"] = asset_clean_answer(update.message.text) or "Other"
    await update.message.reply_text("Choose the asset type:", reply_markup=asset_type_keyboard())
    return ASSET_TYPE


async def asset_type(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["asset_type"] = asset_clean_answer(update.message.text) or "Other"
    await update.message.reply_text("Asset description?\n\nExample: Main plant room boiler / Ground floor DB / Roof AC unit", reply_markup=asset_skip_keyboard())
    return ASSET_DESCRIPTION


async def asset_description(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["description"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Specific location?\n\nExample: Plant room, riser cupboard, roof, kitchen.", reply_markup=asset_skip_keyboard())
    return ASSET_LOCATION


async def asset_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["location"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Manufacturer?", reply_markup=asset_skip_keyboard())
    return ASSET_MANUFACTURER


async def asset_manufacturer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["manufacturer"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Model?", reply_markup=asset_skip_keyboard())
    return ASSET_MODEL


async def asset_model(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["model"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Serial number?", reply_markup=asset_skip_keyboard())
    return ASSET_SERIAL


async def asset_serial(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("asset_create", {})["serial_number"] = asset_clean_answer(update.message.text)
    await update.message.reply_text("Send a photo of the asset, or tap Skip.", reply_markup=asset_skip_keyboard())
    return ASSET_PHOTO


async def asset_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.setdefault("asset_create", {})
    if update.message.photo:
        data["photo_file_id"] = update.message.photo[-1].file_id
        try:
            telegram_file = await context.bot.get_file(data["photo_file_id"])
            photo_bytes = await telegram_file.download_as_bytearray()
            data["photo_data_uri"] = "data:image/jpeg;base64," + base64.b64encode(bytes(photo_bytes)).decode("ascii")
        except Exception as exc:
            print(f"WARNING: Could not download asset photo bytes: {exc}")
        await update.message.reply_text("Photo saved. Any extra notes for the asset?", reply_markup=asset_skip_keyboard())
        return ASSET_NOTES
    if update.message.text == "⏭️ Skip":
        await update.message.reply_text("Any extra notes for the asset?", reply_markup=asset_skip_keyboard())
        return ASSET_NOTES
    await update.message.reply_text("Please send a photo, or tap Skip.", reply_markup=asset_skip_keyboard())
    return ASSET_PHOTO


async def asset_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.setdefault("asset_create", {})
    extra_notes = asset_clean_answer(update.message.text)
    if extra_notes:
        data["notes"] = ((data.get("notes", "") + "\n" + extra_notes).strip())
    try:
        result = create_asset_in_portal(asset_payload_from_context(context))
        asset = result.get("asset", {})
        await update.message.reply_text(
            "✅ Asset created in the portal.\n\n"
            f"Asset ID: {asset.get('asset_id','')}\n"
            f"Type: {asset.get('asset_type','')}\n"
            f"Site: {asset.get('site_name','')}\n\n"
            f"View Asset:\n{result.get('asset_url','')}\n\n"
            f"QR Code:\n{result.get('qr_url','')}\n\n"
            f"Print Label:\n{result.get('label_url','')}",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
    except Exception as e:
        print(f"ERROR creating asset in portal: {e}")
        await update.message.reply_text(
            f"I could not create the asset in the portal.\n\nReason: {e}\n\nCheck PORTAL_BASE_URL and ASSET_API_KEY in Railway.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
    context.user_data.pop("asset_create", None)
    return ConversationHandler.END


async def asset_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("asset_create", None)
    await update.message.reply_text("Asset creation cancelled.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def menu_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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

    if text == MENU_CREATE_ASSET:
        return await asset_start_manual(update, context)

    if text == MENU_REQUEST_JOB:
        return await request_job(update, context)

    if text == MENU_CREATE_ASSET:
        return await asset_start_manual(update, context)

    if text == MENU_QUOTE_REMINDER:
        return await quote_reminder_start(update, context)

    if text == MENU_CANCEL_TASK_ACTIVITY:
        return await start_cancel_task_activity(update, context)

    if text == MENU_ASK_CHATBOT:
        return await ask_chatbot_start(update, context)

    if text in [MENU_CREATE_ASSET, MENU_HELPDESK, MENU_LOG_JOB, MENU_REASSIGN_JOB, MENU_REOPEN_JOB, MENU_OPEN_JOBS, MENU_FIND_JOB, MENU_CANCEL_JOB, MENU_DELETE_JOB, MENU_QUOTE_REMINDER, MENU_CANCEL_TASK_ACTIVITY, MENU_ASK_CHATBOT, MENU_ENGINEER_MENU]:
        return await helpdesk_menu_button(update, context)


async def helpdesk_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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

    if text == MENU_ASK_CHATBOT:
        return await ask_chatbot_start(update, context)

    if text == MENU_LOG_JOB:
        return await logjob_start(update, context)

    if text == MENU_REASSIGN_JOB:
        return await reassign_start(update, context)

    if text == MENU_REOPEN_JOB:
        return await reopenjob_start(update, context)

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

    if text == MENU_CREATE_ASSET:
        return await asset_start_manual(update, context)

    if text == MENU_REQUEST_JOB:
        return await request_job(update, context)

    if text == MENU_CREATE_ASSET:
        return await asset_start_manual(update, context)

    if text == MENU_QUOTE_REMINDER:
        return await quote_reminder_start(update, context)

    if text == MENU_CANCEL_TASK_ACTIVITY:
        return await start_cancel_task_activity(update, context)

    coming_next = {
        MENU_LOG_JOB: "Log Job",
        MENU_REASSIGN_JOB: "Reassign Job",
        MENU_REOPEN_JOB: "Reopen Job",
        MENU_OPEN_JOBS: "Open Jobs",
        MENU_FIND_JOB: "Find Job",
        MENU_CANCEL_JOB: "Cancel Job",
        MENU_DELETE_JOB: "Delete Job",
        MENU_ASK_CHATBOT: "Ask ChatGPT",
        MENU_QUOTE_REMINDER: "Task / Activity",
        MENU_CANCEL_TASK_ACTIVITY: "Cancel Task / Activity",
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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    if not job:
        await update.message.reply_text("Please start again using ➕ Log Job.")
        return ConversationHandler.END

    value = update.message.text.strip().upper()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter a CDR/job number. Example: CDR012896")
        return LOGJOB_CDR_NUMBER

    job["cdr_number"] = value
    await update.message.reply_text(
        "Customer name/address?\n\n"
        "Example:\n"
        "FM4U\n"
        "7-8 Delta Bank Road\n"
        "Gateshead\n"
        "NE11 9DJ"
    )
    return LOGJOB_CUSTOMER_NAME


async def logjob_customer_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the customer name/address. This appears on the worksheet.")
        return LOGJOB_CUSTOMER_NAME

    job["customer_name"] = value
    # Customer address is no longer asked separately. Keep the field blank so it does not duplicate on worksheets.
    job["customer_address"] = ""

    await update.message.reply_text(
        "Site name/address?\n\n"
        "Example:\n"
        "Park View\n"
        "Feetham Avenue, Forest Hall\n"
        "Newcastle\n"
        "NE12 9QN"
    )
    return LOGJOB_SITE_NAME


async def logjob_customer_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    # Legacy state retained for safety only. New Log Job flow no longer asks this separately.
    return await logjob_site_name(update, context)


async def logjob_site_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the site name/address.")
        return LOGJOB_SITE_NAME

    job["site_name"] = value
    # Site address is now included in SiteName/Site block. Do not write to a separate Address field.
    job["site_address"] = ""

    await update.message.reply_text(
        "Contact name/number?",
        reply_markup=get_skip_keyboard(),
    )
    return LOGJOB_CONTACT


async def logjob_contact(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    job["contact"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text("Task / job description?\n\nExample: HQ leaking tea boiler in ELT kitchen")
    return LOGJOB_TASK


async def logjob_task(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the task/job description.")
        return LOGJOB_TASK
    job["task"] = value
    await update.message.reply_text(
        "Any job notes/access notes?",
        reply_markup=get_skip_keyboard(),
    )
    return LOGJOB_NOTES


async def logjob_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    job["notes"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text(
        "Job date?",
        reply_markup=get_date_keyboard(),
    )
    return LOGJOB_DATE


async def logjob_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    if update.message.text.strip().lower() == "custom":
        await update.message.reply_text("Enter the date as DD/MM/YYYY.")
        return LOGJOB_DATE

    parsed = parse_helpdesk_job_date(update.message.text)
    if not parsed:
        await update.message.reply_text(
            "Please enter a valid date, for example today, tomorrow, or 13/05/2026.",
            reply_markup=get_date_keyboard(),
        )
        return LOGJOB_DATE
    job["date"] = parsed
    try:
        job["date_display"] = datetime.fromisoformat(parsed).strftime("%d/%m/%Y")
    except Exception:
        job["date_display"] = parsed
    await update.message.reply_text(
        "Start/time required?",
        reply_markup=get_time_keyboard(),
    )
    return LOGJOB_TIME


async def logjob_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    if update.message.text.strip().lower() == "custom":
        await update.message.reply_text("Enter the time as HH:MM, for example 13:30.")
        return LOGJOB_TIME

    parsed = normalise_helpdesk_time(update.message.text)
    if not parsed:
        await update.message.reply_text(
            "Please enter a valid time, for example 08:00, 0800, 13:30, now, or asap.",
            reply_markup=get_time_keyboard(),
        )
        return LOGJOB_TIME
    job["time"] = parsed
    await update.message.reply_text(
        "Job category? Tap a button or reply with a number:\n\n" +
        "\n".join(f"{i}. {choice}" for i, choice in enumerate(JOB_CATEGORY_CHOICES, start=1)),
        reply_markup=get_category_keyboard(),
    )
    return LOGJOB_CATEGORY


async def logjob_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
            "Please choose a valid category number or tap a category button:\n\n" +
            "\n".join(f"{i}. {choice}" for i, choice in enumerate(JOB_CATEGORY_CHOICES, start=1)),
            reply_markup=get_category_keyboard(),
        )
        return LOGJOB_CATEGORY

    job["category"] = selected
    await update.message.reply_text(
        "Customer order number?",
        reply_markup=get_skip_keyboard(),
    )
    return LOGJOB_ORDER_NUMBER


async def logjob_order_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    value = update.message.text.strip()
    job["order_number"] = "" if is_blank_or_skip(value) else value
    engineers = job.get("assignable_engineers", [])
    await update.message.reply_text(
        "Assign engineer(s), or tap Skip for now to create the job as Awaiting Dispatch.\n\n"
        "Reply with the number, or multiple numbers separated by commas.\n\n" +
        format_engineer_selection_list(engineers),
        reply_markup=get_assign_engineer_keyboard(),
    )
    return LOGJOB_ASSIGN_ENGINEERS


async def logjob_assign_engineers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    engineers = job.get("assignable_engineers", [])
    text = update.message.text.strip()

    if is_blank_or_skip(text) or text.lower() in ["skip for now", "⏭️ skip for now", "awaiting dispatch", "unassigned", "0"]:
        job["assigned_engineers"] = []
        await update.message.reply_text(build_log_job_review(job), reply_markup=get_review_reply_keyboard())
        return LOGJOB_REVIEW

    selected, error = parse_engineer_selection(text, engineers)
    if error:
        await update.message.reply_text(
            error + "\n\nOr tap Skip for now to create the job as Awaiting Dispatch.\n\n" + format_engineer_selection_list(engineers),
            reply_markup=get_assign_engineer_keyboard(),
        )
        return LOGJOB_ASSIGN_ENGINEERS

    job["assigned_engineers"] = selected
    await update.message.reply_text(build_log_job_review(job), reply_markup=get_review_reply_keyboard())
    return LOGJOB_REVIEW


async def logjob_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    job = context.user_data.get("log_job")
    role = job.get("role", "Helpdesk") if job else "Helpdesk"
    answer = update.message.text.strip().lower()
    answer = answer.replace("✅", "").replace("❌", "").replace("🔄", "").strip()

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

        assigned_engineers = job.get("assigned_engineers", [])
        if assigned_engineers:
            sent_to_any, failed = await send_created_job_to_engineers(
                context.bot,
                item_id,
                created_fields,
                assigned_engineers,
            )
        else:
            sent_to_any, failed = False, []

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

        context.user_data.pop("log_job", None)

        assigned_text = ', '.join(e['name'] for e in job.get('assigned_engineers', [])) or 'None - Awaiting Dispatch'
        message = (
            f"Job created in SharePoint: {job['cdr_number']}\n"
            f"Assigned to: {assigned_text}\n"
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
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)
    context.user_data.pop("log_job", None)
    await update.message.reply_text(
        "Log Job cancelled. Nothing has been created.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END



async def reopenjob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)
    if not user_can_use_helpdesk(role):
        await update.message.reply_text("You do not have permission to reopen jobs.", reply_markup=get_main_menu(role))
        return ConversationHandler.END

    try:
        site_id = get_site_id()
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)
        assignable_engineers = get_active_assignable_engineers(engineers)
        context.user_data["reopen_job"] = {
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "engineers": engineers,
            "assignable_engineers": assignable_engineers,
            "role": role,
        }
        await update.message.reply_text("Enter the CDR number to reopen.\n\nExample: CDR01012")
        return REOPENJOB_CDR_NUMBER
    except Exception as e:
        print(f"ERROR opening Reopen Job: {e}")
        await update.message.reply_text("There was an error opening Reopen Job. Please check Railway logs.")
        return ConversationHandler.END


async def reopenjob_cdr_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END
    data = context.user_data.get("reopen_job")
    if not data:
        return ConversationHandler.END
    cdr_number = update.message.text.strip()
    try:
        jobs_data = get_list_items(data["site_id"], data["jobs_list_id"])
        job = find_job_by_cdr(jobs_data, cdr_number)
        if not job:
            await update.message.reply_text("I could not find that CDR number. Please check it and try again.")
            return REOPENJOB_CDR_NUMBER
        fields = job.get("fields", {})
        data.update({"job": job, "item_id": job.get("id"), "job_fields": fields, "cdr_number": get_field_value(fields, "CDRNumber", "CDR Number", "Title") or cdr_number})
        await update.message.reply_text("What date should the reopened job be attended? Example: Today, Tomorrow, 15/05/2026")
        return REOPENJOB_DATE
    except Exception as e:
        print(f"ERROR finding job to reopen: {e}")
        await update.message.reply_text("There was an error finding the job. Please check Railway logs.")
        return ConversationHandler.END


async def reopenjob_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END
    data = context.user_data.get("reopen_job")
    if not data:
        return ConversationHandler.END
    parsed_date = parse_helpdesk_job_date(update.message.text.strip())
    if not parsed_date:
        await update.message.reply_text("Please enter a valid date. Example: Today, Tomorrow, 15/05/2026")
        return REOPENJOB_DATE
    data["new_date"] = parsed_date
    data["new_date_display"] = format_sharepoint_date(parsed_date) or parsed_date
    await update.message.reply_text("Assign engineer(s) for the reopened job. Reply with number(s):\n\n" + format_engineer_selection_list(data.get("assignable_engineers", [])))
    return REOPENJOB_ASSIGN_ENGINEERS


async def reopenjob_assign_engineers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END
    data = context.user_data.get("reopen_job")
    if not data:
        return ConversationHandler.END
    selected, error = parse_engineer_selection(update.message.text, data.get("assignable_engineers", []))
    if error:
        await update.message.reply_text(error + "\n\n" + format_engineer_selection_list(data.get("assignable_engineers", [])))
        return REOPENJOB_ASSIGN_ENGINEERS
    data["assign_engineers"] = selected
    await update.message.reply_text("Reason for reopening? Example: Completed in error / revisit actually required / worksheet correction.")
    return REOPENJOB_REASON


async def reopenjob_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END
    data = context.user_data.get("reopen_job")
    if not data:
        return ConversationHandler.END
    data["reason"] = update.message.text.strip()
    engineers_text = ", ".join(e["name"] for e in data.get("assign_engineers", []))
    await update.message.reply_text(
        "Review reopen job:\n\n"
        f"CDR: {data.get('cdr_number')}\n"
        f"Date: {data.get('new_date_display')}\n"
        f"Assign to: {engineers_text}\n"
        f"Reason: {data.get('reason')}\n\n"
        "This will keep the old worksheet for audit, reset the job for a fresh worksheet, and resend it.\n\n"
        "Reply YES to reopen, or NO to cancel.",
        reply_markup=get_review_reply_keyboard(),
    )
    return REOPENJOB_REVIEW


async def reopenjob_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END
    data = context.user_data.get("reopen_job")
    role = data.get("role", "Helpdesk") if data else "Helpdesk"
    answer = update.message.text.strip().lower().replace("✅", "").replace("❌", "").strip()
    if answer in ["no", "n", "cancel"]:
        context.user_data.pop("reopen_job", None)
        await update.message.reply_text("Reopen cancelled. Nothing has been changed.", reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")))
        return ConversationHandler.END
    if answer not in ["yes", "y"]:
        await update.message.reply_text("Reply YES to reopen, or NO to cancel.")
        return REOPENJOB_REVIEW
    try:
        site_id = data["site_id"]
        jobs_list_id = data["jobs_list_id"]
        fields = data.get("job_fields", {})
        final_engineers = data.get("assign_engineers", [])
        reason = data.get("reason", "Reopened by Helpdesk")
        updated_log = append_engineer_log(fields, "Helpdesk", "Reopened", reason)
        payload_fields = {
            "Date": data.get("new_date"),
            "Status": ASSIGNED_STATUS if final_engineers else AWAITING_DEPLOYMENT_STATUS,
            "JobOutcome": "Revisit Required",
            "Job Outcome": "Revisit Required",
            "WorksheetGenerated": False,
            "Worksheet Generated": False,
            "WorksheetSubmitted": False,
            "Worksheet Submitted": False,
            "TelegramNotified": False,
            "Telegram Notified": False,
            "ClientSignatureRequired": False,
            "ClientSignatureReceived": False,
            "SignatureToken": "",
            "EngineerVisitLog": updated_log,
            "Engineer Visit Log": updated_log,
        }
        payload = build_field_payload_for_list(site_id, jobs_list_id, payload_fields)
        payload["EngineerLookupId@odata.type"] = "Collection(Edm.Int32)"
        payload["EngineerLookupId"] = [int(e["lookup_id"]) for e in final_engineers]
        update_list_item_fields(site_id, jobs_list_id, data["item_id"], payload)

        send_fields = dict(fields)
        send_fields.update(payload_fields)
        sent_to_any, failed = await send_created_job_to_engineers(context.bot, data["item_id"], send_fields, final_engineers)

        final_payload = build_field_payload_for_list(site_id, jobs_list_id, {"TelegramNotified": bool(sent_to_any), "Telegram Notified": bool(sent_to_any)})
        update_list_item_fields(site_id, jobs_list_id, data["item_id"], final_payload)

        context.user_data.pop("reopen_job", None)
        message = f"Job reopened and resent: {data.get('cdr_number')}\nAssigned to: {', '.join(e['name'] for e in final_engineers)}"
        if failed:
            message += "\n\nSend issues:\n" + "\n".join(failed[:5])
        await update.message.reply_text(message, reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")))
        return ConversationHandler.END
    except Exception as e:
        print(f"ERROR reopening job: {e}")
        await update.message.reply_text("There was an error reopening the job. Please check Railway logs.", reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")))
        return ConversationHandler.END


async def reopenjob_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    context.user_data.pop("reopen_job", None)
    await update.message.reply_text("Reopen Job cancelled.", reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role))
    return ConversationHandler.END

async def reassign_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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

    await update.message.reply_text(
        "What date should this reassignment be attended?\n\n"
        "Reply Today, Tomorrow, DD/MM/YYYY, or KEEP to leave the current job date unchanged."
    )
    return REASSIGN_DATE



async def reassign_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    data = context.user_data.get("reassign_job")
    if not data:
        await update.message.reply_text("Please start again using 🔁 Reassign Job.")
        return ConversationHandler.END

    value = update.message.text.strip()
    if value.lower() in ["keep", "same", "skip", "no change"]:
        data["new_date"] = None
        data["new_date_display"] = "Keep current date"
    else:
        parsed_date = parse_helpdesk_job_date(value)
        if not parsed_date:
            await update.message.reply_text("Please enter a valid date, or type KEEP. Example: Today, Tomorrow, 15/05/2026")
            return REASSIGN_DATE
        data["new_date"] = parsed_date
        data["new_date_display"] = format_sharepoint_date(parsed_date) or parsed_date

    await update.message.reply_text("Reason for reassignment?\n\nExample: revisit booked / no access rebooked / change of plan / emergency priority.")
    return REASSIGN_REASON

async def reassign_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    data = context.user_data.get("reassign_job")
    if not data:
        await update.message.reply_text("Please start again using 🔁 Reassign Job.")
        return ConversationHandler.END

    value = update.message.text.strip()
    data["reason"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text(build_reassign_review(data))
    return REASSIGN_REVIEW


async def reassign_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    data = context.user_data.get("reassign_job")
    role = data.get("role", "Helpdesk") if data else "Helpdesk"
    answer = update.message.text.strip().lower()
    answer = answer.replace("✅", "").replace("❌", "").replace("🔄", "").strip()

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
        if data.get("new_date"):
            reassignment_note += f"; Attendance Date: {data.get('new_date_display') or data.get('new_date')}"
        else:
            reassignment_note += "; Attendance Date: Keep current date"

        if data.get("reason"):
            reassignment_note += f"; Reason: {data['reason']}"

        updated_log = append_engineer_log(fields, "Helpdesk", "Reassigned", reassignment_note)

        # First place the job back into dispatch while the assignment is changed.
        interim_fields = {
            "Status": AWAITING_DEPLOYMENT_STATUS,
            "TelegramNotified": False,
            "Telegram Notified": False,
            "EngineerVisitLog": updated_log,
            "Engineer Visit Log": updated_log,
        }
        if data.get("new_date"):
            interim_fields["Date"] = data.get("new_date")
        interim_payload = build_field_payload_for_list(
            site_id,
            jobs_list_id,
            interim_fields,
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
        if data.get("new_date"):
            send_fields["Date"] = data.get("new_date")
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
    if is_group_chat(update):
        return ConversationHandler.END

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

    customer_detail = get_field_value(fields, 'CustomerName', 'Customer Name') or ''
    site_detail = get_field_value(fields, 'SiteName', 'Site Name') or ''

    return (
        "Job details:\n\n"
        f"CDR Number: {get_field_value(fields, 'CDRNumber', 'CDR Number', 'Title') or ''}\n"
        f"Customer: {customer_detail}\n"
        f"Site: {site_detail}\n"
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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)
    context.user_data.pop("open_jobs", None)
    await update.message.reply_text(
        "Open Jobs cancelled.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END





async def canceljob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)
    context.user_data.pop("cancel_job", None)
    await update.message.reply_text(
        "Cancel Job stopped. Nothing has been changed.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(str(role).lower() == "admin")) if user_can_use_helpdesk(role) else get_main_menu(role),
    )
    return ConversationHandler.END


async def deletejob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

    context.user_data.pop("delete_job", None)
    await update.message.reply_text(
        "Hard Delete stopped. Nothing has been changed.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=True),
    )
    return ConversationHandler.END


async def findjob_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    MENU_REOPEN_JOB,
    MENU_OPEN_JOBS,
    MENU_FIND_JOB,
    MENU_UPLOAD_RECEIPTS,
    MENU_CANCEL_JOB,
    MENU_DELETE_JOB,
    MENU_QUOTE_REMINDER,
    MENU_MESSAGE_ENGINEER,
    MENU_CANCEL_TASK_ACTIVITY,
    MENU_ASK_CHATBOT,
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
    must never end up in Work Completed, Materials, Notes or worksheet comments.
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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

    context.user_data.pop("bug_idea", None)
    await update.message.reply_text("Bug/idea cancelled.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


def get_receipt_type_keyboard():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("🧾 Receipts", callback_data="receipt_type|receipt"),
            InlineKeyboardButton("↩️ Returns", callback_data="receipt_type|return"),
        ],
        [InlineKeyboardButton("❌ Cancel", callback_data="receipt_type|cancel")],
    ])


async def receipt_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
            print(f"WARNING: Could not auto-detect receipt/return engineer: {e}")

        if not engineer_name:
            await update.message.reply_text(
                "I could not match your Telegram account to an engineer record, so I cannot upload receipts/returns under your name. Please ask the office to check your Engineers list record has your Telegram ID and EngineerName.",
                reply_markup=get_main_menu(await get_role_for_update(update)),
            )
            return ConversationHandler.END

        context.user_data["receipt_upload"] = {
            "site_id": site_id,
            "engineer_name": engineer_name,
            "upload_type": "",
            "receipt_date": "",
            "receipt_links": [],
            "receipt_count": 0,
        }

        await update.message.reply_text(
            f"Receipts / Returns upload started.\n\nEngineer: {engineer_name}\n\nWhat are you uploading?",
            reply_markup=get_receipt_type_keyboard(),
        )
        return RECEIPT_TYPE

    except Exception as e:
        print(f"ERROR starting receipts/returns upload: {e}")
        await update.message.reply_text(
            "There was an error starting receipts/returns upload. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END


async def receipt_type_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    data = context.user_data.get("receipt_upload")
    if not data:
        await query.message.reply_text("Please start again from 🧾 Receipts / Returns.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    selected = query.data.split("|", 1)[1]

    if selected == "cancel":
        context.user_data.pop("receipt_upload", None)
        await query.message.reply_text("Receipts / Returns upload cancelled.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    if selected not in ["receipt", "return"]:
        await query.message.reply_text("Please choose Receipts or Returns.", reply_markup=get_receipt_type_keyboard())
        return RECEIPT_TYPE

    data["upload_type"] = "Return" if selected == "return" else "Receipt"
    label = "returns" if selected == "return" else "receipts"

    await query.message.reply_text(
        f"Uploading {label}.\n\nEngineer: {data['engineer_name']}\n\nEnter the date.\n\nUse DD/MM/YYYY, YYYY-MM-DD, or TODAY.\n\nType /cancel to cancel."
    )
    return RECEIPT_DATE


async def receipt_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    menu_result = await handle_menu_during_conversation(update, context, RECEIPT_DATE)
    if menu_result is not None:
        return menu_result

    data = context.user_data.get("receipt_upload")
    if not data:
        await update.message.reply_text("Please start again from 🧾 Receipts / Returns.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    parsed_date = parse_helpdesk_job_date(update.message.text)
    if not parsed_date:
        await update.message.reply_text("Please enter a valid date, for example 12/05/2026 or TODAY.")
        return RECEIPT_DATE

    data["receipt_date"] = parsed_date
    upload_label = "Returns" if str(data.get("upload_type", "")).lower() == "return" else "Receipts"

    await update.message.reply_text(
        f"{upload_label} upload ready.\n\nEngineer: {data['engineer_name']}\nDate: {parsed_date}\n\n"
        "Now send all photos/PDFs/files.\n\nType DONE when finished."
    )
    return RECEIPT_UPLOADS


async def receipt_uploads(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    menu_result = await handle_menu_during_conversation(update, context, RECEIPT_UPLOADS)
    if menu_result is not None:
        return menu_result

    data = context.user_data.get("receipt_upload")
    if not data:
        await update.message.reply_text("Please start again from 🧾 Receipts / Returns.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    upload_label = "Returns" if str(data.get("upload_type", "")).lower() == "return" else "Receipts"

    if update.message.text and update.message.text.strip().upper() == "DONE":
        count = len(data.get("receipt_links", []))
        if count == 0:
            await update.message.reply_text(f"No {upload_label.lower()} uploaded yet. Send at least one file, or type /cancel to cancel.")
            return RECEIPT_UPLOADS

        context.user_data.pop("receipt_upload", None)
        await update.message.reply_text(
            f"{upload_label} uploaded for finance.\n\nEngineer: {data['engineer_name']}\nDate: {data['receipt_date']}\nFiles uploaded: {count}",
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
            await update.message.reply_text(f"Please send a {upload_label.lower()} photo/file, or type DONE when finished.")
            return RECEIPT_UPLOADS

        data["receipt_count"] = int(data.get("receipt_count", 0)) + 1
        timestamp = datetime.now(UK_TZ).strftime("%Y%m%d_%H%M%S")
        engineer_part = safe_folder_name(data.get("engineer_name", "ENGINEER"))
        upload_type = str(data.get("upload_type") or "Receipt")
        type_part = "return" if upload_type.lower() == "return" else "receipt"
        file_name = f"{data['receipt_date']}_{engineer_part}_{timestamp}_{data['receipt_count']}_{type_part}_{unique_id}.{extension}"

        receipt_link = upload_receipt_to_sharepoint(
            data["site_id"],
            data["receipt_date"],
            data["engineer_name"],
            file_name,
            bytes(file_bytes),
            upload_type,
        )
        data["receipt_links"].append(receipt_link)

        # Intentionally no per-file confirmation. Engineers can upload all receipts/returns
        # and type DONE once finished to avoid Telegram message spam.
        return RECEIPT_UPLOADS

    except Exception as e:
        print(f"ERROR uploading receipt/return: {e}")
        await update.message.reply_text(
            "There was an error uploading that file. Please try again or ask the office to check Railway logs."
        )
        return RECEIPT_UPLOADS


async def receipt_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("receipt_upload", None)
    await update.message.reply_text("Receipts / Returns upload cancelled.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    await update.message.reply_text(f"Your Telegram ID is: {update.effective_user.id}")


def task_activity_matches_user_today(fields, telegram_id):
    status = str(get_field_value(fields, "Status") or "Active").strip().lower()
    if status != "active":
        return False

    task_date = task_date_iso_from_fields(fields)
    if task_date != get_today_iso():
        return False

    return str(telegram_id) in task_telegram_ids_from_fields(fields)


def format_task_activity_for_my_jobs(fields):
    title = str(get_field_value(fields, "Title") or "Task / Activity").strip()
    task_date = task_date_iso_from_fields(fields)
    time_text = str(get_field_value(fields, "TaskTime", "Task Time") or "").strip()
    details = str(get_field_value(fields, "TaskDetails", "Task Details") or "").strip()

    return (
        "📋 Task / Activity:\n\n"
        f"Title: {title}\n"
        f"Date: {format_sharepoint_date(task_date) or task_date or 'N/A'}\n"
        f"Time: {time_text or 'N/A'}\n"
        f"Details:\n{details or 'N/A'}"
    )


def get_today_task_activities_for_user(site_id, telegram_id):
    try:
        tasks_list_id = get_list_id(site_id, TASK_ACTIVITIES_LIST)
        tasks = get_list_items(site_id, tasks_list_id)
    except Exception as e:
        print(f"WARNING: Could not load Task Activities for My Jobs: {e}")
        return []

    return [
        task for task in tasks
        if task_activity_matches_user_today(task.get("fields", {}), telegram_id)
    ]


async def jobs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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

        role = await get_role_for_update(update)
        today_tasks = get_today_task_activities_for_user(site_id, user_id)

        if not engineer_has_active_day(site_id, user_id):
            if today_tasks:
                for task in today_tasks:
                    await update.message.reply_text(
                        format_task_activity_for_my_jobs(task.get("fields", {})),
                        reply_markup=get_main_menu(role),
                    )

            if user_can_use_helpdesk(role):
                await update.message.reply_text(
                    "You are logged in as helpdesk/admin. Engineer jobs are locked until you start your engineer day. Use the Helpdesk menu for office actions.",
                    reply_markup=get_main_menu(role),
                )
            else:
                await update.message.reply_text(
                    "Please start your day first using 🟢 Start Day or /startday. Your engineer jobs are locked until your day has started.",
                    reply_markup=get_main_menu(role),
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
                    reply_markup=get_job_buttons(item_id, fields.get("SiteName", "")),
                )

        for task in today_tasks:
            found_any = True
            await update.message.reply_text(
                format_task_activity_for_my_jobs(task.get("fields", {})),
                reply_markup=get_main_menu(role),
            )

        if not found_any:
            await update.message.reply_text("No jobs or task / activity items assigned today.")

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
        remaining_ids = get_remaining_assigned_engineer_ids(fields, engineer_lookup_id)
        is_final_engineer = len(remaining_ids) == 0

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
    if is_group_chat(update):
        return ConversationHandler.END

    context.user_data.pop("abort_job", None)
    await update.message.reply_text("Abort job cancelled. No changes made.", reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END





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
            "photo_files_for_group": [],
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


async def status_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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

        if action == "undo_travelling":
            has_travelled = has_engineer_action(fields, current_engineer["name"], "Travelling")
            has_on_site = has_engineer_action(fields, current_engineer["name"], "On Site")
            current_status = str(fields.get("Status", "") or "").strip()

            if not has_travelled:
                await query.message.reply_text("Travelling has not been logged on this job, so there is nothing to undo.")
                return

            if has_on_site or current_status == ON_SITE_STATUS:
                await query.message.reply_text("Travelling cannot be undone after Arrived On Site has been logged.")
                return

            updated_log = append_engineer_log(
                fields,
                current_engineer["name"],
                "Travelling Reverted",
                "Accidental travelling click reversed by engineer",
            )

            update_list_item_fields(
                site_id,
                jobs_list_id,
                item_id,
                {
                    "Status": ASSIGNED_STATUS,
                    "EngineerVisitLog": updated_log,
                    "WorksheetSubmitted": False,
                },
            )

            update_active_day_live_status(
                site_id,
                user_id,
                ASSIGNED_STATUS,
                get_job_reference(fields),
            )

            await query.message.reply_text(
                f"Travelling reversed:\n\n{fields.get('CDRNumber', '')} → Assigned\n\nThe job is still assigned to you. You can press Start Travelling again when needed."
            )

            await notify_helpdesk(
                context,
                (
                    f"Job travelling reversed\n\n"
                    f"CDR Number: {fields.get('CDRNumber', '')}\n"
                    f"Engineer: {current_engineer['name']}\n"
                    f"Site: {fields.get('SiteName', '')}\n"
                    "The job has been put back to Assigned and remains with the engineer."
                ),
            )

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

            reply_markup = None
            if selected_status == "Travelling":
                reply_markup = InlineKeyboardMarkup([
                    [InlineKeyboardButton("↩️ Undo Travelling", callback_data=f"undo_travelling|{item_id}")]
                ])

            await query.message.reply_text(
                f"Status updated:\n\n{fields.get('CDRNumber', '')} → {selected_status}",
                reply_markup=reply_markup,
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

            remaining_ids = get_remaining_assigned_engineer_ids(fields, current_engineer["lookup_id"])
            is_final_engineer = len(remaining_ids) == 0

            update_fields = {
                "EngineerVisitLog": updated_log,
            }

            if is_final_engineer:
                update_fields["JobOutcome"] = selected_outcome
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
                f"Your attendance has been logged. Other assigned engineer(s) remain on the job until they submit their own worksheet."
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





async def complete_button_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    try:
        _, item_id, outcome = query.data.split("|", 2)
    except Exception:
        await query.message.reply_text("Could not start this worksheet. Tap 📋 My Jobs and try again.")
        return ConversationHandler.END

    return await begin_worksheet_for_job(update, context, item_id=item_id, outcome=outcome)


async def noaccess_reason_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    await update.message.reply_text(
        "Upload job photos now.\n\n"
        "Send one or more photos, then type DONE when finished.\n"
        "If no photos are needed, type DONE."
    )
    return PHOTOS


async def worksheet_follow_on_required_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    await query.message.reply_text(
        "Upload job photos now.\n\n"
        "Send one or more photos, then type DONE when finished.\n"
        "If no photos are needed, type DONE."
    )
    return PHOTOS


async def worksheet_follow_on_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    menu_result = await handle_menu_during_conversation(update, context, FOLLOW_ON_NOTES)
    if menu_result is not None:
        return menu_result

    if is_bot_menu_text(update.message.text):
        await update.message.reply_text("That is a menu button, so I have not added it to the worksheet. Please type the follow-on required, or type /cancel.")
        return FOLLOW_ON_NOTES

    context.user_data["worksheet"]["FollowOnNotes"] = update.message.text
    await update.message.reply_text(
        "Upload job photos now.\n\n"
        "Send one or more photos, then type DONE when finished.\n"
        "If no photos are needed, type DONE."
    )
    return PHOTOS


def get_photos_done_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Continue - photo count is correct", callback_data="photos_done|continue")],
        [InlineKeyboardButton("➕ Add more photos / wait for uploads", callback_data="photos_done|add_more")],
    ])


async def worksheet_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    menu_result = await handle_menu_during_conversation(update, context, PHOTOS)
    if menu_result is not None:
        return menu_result

    worksheet = context.user_data["worksheet"]

    if update.message.text and update.message.text.strip().upper() == "DONE":
        photo_count = len(worksheet.get("photo_links", []))
        await update.message.reply_text(
            f"I have received {photo_count} job photo(s).\n\n"
            "Check this number before continuing. If the engineer sent 22 photos but this says 18, "
            "do NOT continue yet. Wait for the missing photos to appear, or tap Add more photos / wait for uploads and resend the missing photos.\n\n"
            "Tap Continue only when the photo count is definitely correct.",
            reply_markup=get_photos_done_keyboard(),
        )
        return PHOTOS

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
        worksheet.setdefault("photo_files_for_group", []).append({
            "file_name": file_name,
            "bytes": bytes(file_bytes),
        })

        count = len(worksheet.get("photo_links", []))

        # Only acknowledge the first successful photo upload.
        # This avoids spamming the engineer and reduces Telegram API calls when
        # 10, 20+ photos are being uploaded from poor signal areas.
        if not worksheet.get("photo_upload_started_notice_sent"):
            worksheet["photo_upload_started_notice_sent"] = True
            await update.message.reply_text(
                "📷 Photo upload started. Send all required photos, then type DONE when finished."
            )

        return PHOTOS

    await update.message.reply_text("Please send job photos, or type DONE when finished.")
    return PHOTOS


async def worksheet_photos_done_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    action = query.data.split("|", 1)[1]
    photo_count = len(worksheet.get("photo_links", []))

    if action == "add_more":
        await query.message.reply_text(
            f"No problem. I currently have {photo_count} photo(s). Send the missing photos now, then type DONE again."
        )
        return PHOTOS

    if action == "continue":
        await query.message.reply_text(
            "Is a client signature required?",
            reply_markup=get_yes_no_keyboard("signature_required"),
        )
        return SIGNATURE_REQUIRED

    await query.message.reply_text("Please tap Continue or Add more photos.", reply_markup=get_photos_done_keyboard())
    return PHOTOS


async def worksheet_signature_required(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    if is_group_chat(update):
        return ConversationHandler.END

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
    """Keep EngineerVisitLog extra text to one safe line so it can be parsed back into the worksheet."""
    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def build_visit_comment_extra(worksheet):
    """Create a compact visit comment for EngineerVisitLog and the final worksheet document."""
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

        # Office/helpdesk/internal audit lines are useful in SharePoint, but they are not
        # customer-facing engineer attendances and must not appear on the worksheet.
        internal_actions = {
            "Reassigned",
            "Cancelled",
            "Hard Deleted",
            "Job logged via Telegram",
            "Aborted Attendance",
        }

        is_internal_actor = key in ["helpdesk", "admin", "office"]
        is_internal_action = action in internal_actions or action.startswith("Job logged via Telegram")

        if action == "Aborted Attendance":
            # If an engineer travelled/on-site then aborted, remove that incomplete
            # attendance from the customer worksheet entirely. It stays in SharePoint
            # for audit, but it is not a client visit/no-access/completion record.
            current = active_by_engineer.pop(key, None)
            if current and current in visits:
                visits.remove(current)
            continue

        if is_internal_actor or is_internal_action:
            continue

        if action == "Travelling Reverted":
            current = active_by_engineer.pop(key, None)
            if current and current in visits:
                visits.remove(current)
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
        elif action.startswith("Submitted for Helpdesk Review - "):
            # Compatibility for older builds. These lines were created when the engineer
            # submitted the worksheet, so they are the engineer's off-site/final action
            # and must still count for productivity and worksheet visit history.
            submitted_outcome = action.replace("Submitted for Helpdesk Review - ", "", 1).strip() or "Completed"
            current["status"] = submitted_outcome
            current["off_site"] = visit_time
            if extra:
                current["notes"] = extra
            active_by_engineer.pop(key, None)
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

    customer_details = get_field_value(fields, "CustomerName", "Customer Name") or get_field_value(fields, "ClientName", "Client Name") or ""
    site_details = fields.get("SiteName", "") or ""

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




def clean_docx_text(value):
    text = str(value or "").strip()
    return text if text and text.lower() not in ["none", "nan"] else "N/A"


def safe_docx_filename(value):
    cleaned = "".join(ch for ch in str(value or "").strip() if ch.isalnum() or ch in ["-", "_", " "]).strip()
    return cleaned.replace(" ", "_") or "worksheet"


def docx_escape(value):
    text = str(value or "")
    text = "".join(ch for ch in text if ch in "\t\n\r" or ord(ch) >= 32)
    return xml_escape(text)


DOCX_PAGE_WIDTH = 11906
DOCX_CONTENT_WIDTH = 10320
DOCX_ORANGE = "595959"  # worksheet section/header grey
DOCX_DARK = "333333"
DOCX_LIGHT = "F7F7F7"
DOCX_GREY = "D9D9D9"


def docx_run(text, bold=False, size=18, color=None):
    rpr = f"<w:sz w:val='{size}'/>"
    if bold:
        rpr += "<w:b/>"
    if color:
        rpr += f"<w:color w:val='{color}'/>"
    return f"<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space='preserve'>{docx_escape(text)}</w:t></w:r>"


def docx_paragraph(value="", bold=False, size=18, align=None, spacing_after=80, color=None, clean=False):
    ppr = f"<w:spacing w:after='{spacing_after}'/>"
    if align:
        ppr += f"<w:jc w:val='{align}'/>"

    text = clean_docx_text(value) if clean else str(value or "")
    lines = text.splitlines() or [""]
    runs = []
    for i, line in enumerate(lines):
        if i:
            runs.append("<w:r><w:br/></w:r>")
        runs.append(docx_run(line, bold=bold, size=size, color=color))

    return f"<w:p><w:pPr>{ppr}</w:pPr>{''.join(runs)}</w:p>"


def docx_cell(value="", bold=False, width=2500, shade=None, color=None, size=18, align=None):
    tcpr = f"<w:tcPr><w:tcW w:w='{width}' w:type='dxa'/>"
    tcpr += "<w:tcMar><w:top w:w='90' w:type='dxa'/><w:left w:w='120' w:type='dxa'/><w:bottom w:w='90' w:type='dxa'/><w:right w:w='120' w:type='dxa'/></w:tcMar>"
    tcpr += "<w:vAlign w:val='top'/>"
    if shade:
        tcpr += f"<w:shd w:val='clear' w:color='auto' w:fill='{shade}'/>"
    tcpr += "</w:tcPr>"
    return f"<w:tc>{tcpr}{docx_paragraph(value, bold=bold, size=size, spacing_after=0, color=color, align=align, clean=False)}</w:tc>"


def docx_table(rows, header_first=False, widths=None, table_width=DOCX_CONTENT_WIDTH, header_orange=False, label_columns=None):
    if not rows:
        return ""

    column_count = max(len(row) for row in rows)
    if not widths:
        widths = [int(table_width / column_count)] * column_count

    grid = "".join(f"<w:gridCol w:w='{widths[min(i, len(widths)-1)]}'/>" for i in range(column_count))
    tbl = [
        "<w:tbl>",
        "<w:tblPr>",
        f"<w:tblW w:w='{table_width}' w:type='dxa'/>",
        "<w:tblLayout w:type='fixed'/>",
        "<w:tblBorders>",
        f"<w:top w:val='single' w:sz='6' w:space='0' w:color='{DOCX_DARK}'/>",
        f"<w:left w:val='single' w:sz='6' w:space='0' w:color='{DOCX_DARK}'/>",
        f"<w:bottom w:val='single' w:sz='6' w:space='0' w:color='{DOCX_DARK}'/>",
        f"<w:right w:val='single' w:sz='6' w:space='0' w:color='{DOCX_DARK}'/>",
        f"<w:insideH w:val='single' w:sz='4' w:space='0' w:color='{DOCX_GREY}'/>",
        f"<w:insideV w:val='single' w:sz='4' w:space='0' w:color='{DOCX_GREY}'/>",
        "</w:tblBorders>",
        "<w:tblCellMar><w:top w:w='90' w:type='dxa'/><w:left w:w='120' w:type='dxa'/><w:bottom w:w='90' w:type='dxa'/><w:right w:w='120' w:type='dxa'/></w:tblCellMar>",
        "</w:tblPr>",
        f"<w:tblGrid>{grid}</w:tblGrid>",
    ]

    label_columns = set(label_columns or [])
    for row_index, row in enumerate(rows):
        is_header = header_first and row_index == 0
        tbl.append("<w:tr>")
        for col_index in range(column_count):
            raw_value = row[col_index] if col_index < len(row) else ""
            value = clean_docx_text(raw_value)
            is_label = col_index in label_columns and not is_header
            shade = None
            color = None
            bold = False
            if is_header:
                shade = DOCX_ORANGE if header_orange else DOCX_LIGHT
                color = "FFFFFF" if header_orange else None
                bold = True
            elif is_label:
                shade = DOCX_LIGHT
                bold = True
            tbl.append(docx_cell(
                value,
                bold=bold,
                width=widths[min(col_index, len(widths)-1)],
                shade=shade,
                color=color,
                size=17,
            ))
        tbl.append("</w:tr>")

    tbl.append("</w:tbl>")
    return "".join(tbl)


def docx_spacer(height=90):
    return docx_paragraph("", spacing_after=height)


def docx_section(title, value):
    return (
        docx_table([[title], [clean_docx_text(value)]], header_first=True, header_orange=True, widths=[DOCX_CONTENT_WIDTH])
        + docx_spacer(100)
    )

def get_local_logo_image_bytes_for_docx():
    """Return the local CDR logo bytes for the worksheet DOCX, if available."""
    for path in ["cdr-logo.png", "CDR-logo.png", "logo.png", "/app/logo.png", "/mnt/data/logo.png"]:
        try:
            if os.path.exists(path):
                with open(path, "rb") as logo_file:
                    return logo_file.read()
        except Exception as e:
            print(f"Could not load worksheet logo {path}: {e}")
    return None


def docx_image_paragraph(relationship_id, width_emu=3024000, height_emu=1058400, name="Image", align="left"):
    """Return WordprocessingML for an embedded image.

    width_emu/height_emu control the displayed image size in Word.
    """
    return f"""
<w:p>
<w:pPr><w:spacing w:after='100'/><w:jc w:val='{align}'/></w:pPr>
<w:r>
<w:drawing>
<wp:inline distT='0' distB='0' distL='0' distR='0'>
<wp:extent cx='{width_emu}' cy='{height_emu}'/>
<wp:effectExtent l='0' t='0' r='0' b='0'/>
<wp:docPr id='1' name='{docx_escape(name)}'/>
<wp:cNvGraphicFramePr/>
<a:graphic>
<a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/picture'>
<pic:pic>
<pic:nvPicPr>
<pic:cNvPr id='0' name='{docx_escape(name)}.png'/>
<pic:cNvPicPr/>
</pic:nvPicPr>
<pic:blipFill>
<a:blip r:embed='{relationship_id}'/>
<a:stretch><a:fillRect/></a:stretch>
</pic:blipFill>
<pic:spPr>
<a:xfrm><a:off x='0' y='0'/><a:ext cx='{width_emu}' cy='{height_emu}'/></a:xfrm>
<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>
</pic:spPr>
</pic:pic>
</a:graphicData>
</a:graphic>
</wp:inline>
</w:drawing>
</w:r>
</w:p>"""



def build_worksheet_docx_bytes(worksheet, fields, updated_log, outcome, site_id=None):
    # Build an editable DOCX matching the previous professional PDF worksheet layout.
    cdr_number = worksheet.get("cdr_number") or fields.get("CDRNumber", "") or fields.get("Title", "")
    date_logged = format_sharepoint_date(fields.get("Date", ""))
    date_complete = datetime.now(UK_TZ).strftime("%d/%m/%Y")

    customer_details = (
        get_field_value(fields, "CustomerName", "Customer Name")
        or get_field_value(fields, "ClientName", "Client Name")
        or ""
    )
    site_details = fields.get("SiteName", "") or ""
    order_number = get_field_value(fields, "CustomerOrderNumber", "Customer Order Number", "OrderNumber", "Order Number") or ""
    job_category = get_field_value(fields, "JobCategory", "Job Category") or ""
    task = fields.get("Task", "") or fields.get("Description", "") or fields.get("Notes", "")

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

    visit_rows = [["Date", "Travel", "On-Site", "Engineer", "Status", "Off-Site"]]
    for visit in visits:
        visit_rows.append([
            visit.get("date", ""),
            visit.get("travel", ""),
            visit.get("on_site", ""),
            visit.get("engineer", ""),
            visit.get("status", ""),
            visit.get("off_site", ""),
        ])

    comments = build_engineer_comments_for_pdf(visits, worksheet, fields)

    signature_image_bytes = None
    if worksheet.get("ClientSignatureRequired"):
        if worksheet.get("ClientSignatureReceived"):
            signature_image_bytes = get_signature_image_bytes(site_id, cdr_number) if site_id else None
            signature_text = (
                f"Client Name: {worksheet.get('ClientSignatureName', '')}\n"
                "Signed Digitally: Yes"
                + ("\nSignature embedded below." if signature_image_bytes else "\nSignature image saved in SharePoint but could not be embedded.")
            )
        else:
            signature_text = "Client signature required but not received."
    else:
        signature_text = "Client signature not required."

    logo_image_bytes = get_local_logo_image_bytes_for_docx()

    body = []
    if logo_image_bytes:
        body.append(docx_image_paragraph("rIdLogo1", width_emu=2250000, height_emu=900000, name="CDR Logo", align="center"))
    body.append(docx_paragraph("JOB WORKSHEET", bold=True, size=32, align="center", spacing_after=60))
    body.append(docx_paragraph(
        "CDR M&E Services Ltd\n"
        "6 Mandale Park, Urlay Nook Road, Egglescliffe, Stockton-on-Tees, TS16 0TA\n"
        "Telephone: 01642 057939 | Email: helpdesk@cdrme.co.uk\n"
        "VAT Number: 397715249 | Company No.: 13744971",
        size=17,
        align="center",
        spacing_after=160,
    ))

    body.append(docx_table(
        [["Customer Details", "Site Details"], [customer_details, site_details]],
        header_first=True,
        header_orange=True,
        widths=[5160, 5160],
    ))
    body.append(docx_spacer(110))

    body.append(docx_table(
        [
            ["Job Number", cdr_number, "Customer Order Number", order_number],
            ["Date Logged", date_logged, "Job Category", job_category],
            ["Date Complete", date_complete, "Status", outcome],
        ],
        widths=[1950, 3210, 2450, 2710],
        label_columns={0, 2},
    ))
    body.append(docx_spacer(110))

    body.append(docx_section("Description", task))
    body.append(docx_table(visit_rows, header_first=True, header_orange=True, widths=[1450, 1250, 1250, 2550, 2200, 1620]))
    body.append(docx_spacer(110))
    body.append(docx_section("Engineer Comment", comments))
    body.append(docx_table([["Client Signature"], [signature_text]], header_first=True, header_orange=True, widths=[DOCX_CONTENT_WIDTH]))
    if signature_image_bytes:
        body.append(docx_image_paragraph("rIdSignature1", width_emu=3024000, height_emu=1058400, name="Client Signature", align="left"))
    body.append(docx_spacer(100))
    body.append(docx_paragraph(
        "CDR M&E Services Ltd | 01642 057939 | helpdesk@cdrme.co.uk",
        size=16,
        align="center",
        spacing_after=0,
    ))

    document_xml = f'''<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture'>
<w:body>
{''.join(body)}
<w:sectPr>
<w:pgSz w:w='{DOCX_PAGE_WIDTH}' w:h='16838'/>
<w:pgMar w:top='680' w:right='793' w:bottom='680' w:left='793' w:header='720' w:footer='720' w:gutter='0'/>
</w:sectPr>
</w:body>
</w:document>'''

    content_types = '''<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
<Default Extension='xml' ContentType='application/xml'/>
<Default Extension='png' ContentType='image/png'/>
<Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>
<Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml'/>
<Override PartName='/docProps/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/>
</Types>'''

    root_rels = '''<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>
<Relationship Id='rId2' Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties' Target='docProps/core.xml'/>
<Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties' Target='docProps/app.xml'/>
</Relationships>'''

    doc_relationships = []
    if logo_image_bytes:
        doc_relationships.append("<Relationship Id='rIdLogo1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/cdr_logo.png'/>")
    if signature_image_bytes:
        doc_relationships.append("<Relationship Id='rIdSignature1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/client_signature.png'/>")

    doc_rels = (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        + "".join(doc_relationships)
        + "</Relationships>"
    )

    now_iso = datetime.now(UK_TZ).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/' xmlns:dcmitype='http://purl.org/dc/dcmitype/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
<dc:title>{docx_escape(cdr_number)} Worksheet</dc:title>
<dc:creator>CDR Engineer Bot</dc:creator>
<cp:lastModifiedBy>CDR Engineer Bot</cp:lastModifiedBy>
<dcterms:created xsi:type='dcterms:W3CDTF'>{now_iso}</dcterms:created>
<dcterms:modified xsi:type='dcterms:W3CDTF'>{now_iso}</dcterms:modified>
</cp:coreProperties>'''

    app_props = '''<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'>
<Application>CDR Engineer Bot</Application>
</Properties>'''

    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", root_rels)
        docx.writestr("word/_rels/document.xml.rels", doc_rels)
        docx.writestr("word/document.xml", document_xml)
        if logo_image_bytes:
            docx.writestr("word/media/cdr_logo.png", logo_image_bytes)
        if signature_image_bytes:
            docx.writestr("word/media/client_signature.png", signature_image_bytes)
        docx.writestr("docProps/core.xml", core)
        docx.writestr("docProps/app.xml", app_props)

    return buffer.getvalue()



def generate_and_upload_worksheet_pdf(site_id, jobs_list_id, item_id, worksheet, fields, updated_log, outcome):
    # Generate and upload editable Word worksheet using the professional PDF-style layout.
    # Function name is kept for compatibility with existing worksheet flow and SharePoint fields.
    try:
        docx_bytes = build_worksheet_docx_bytes(worksheet, fields, updated_log, outcome, site_id)
        cdr_number = worksheet.get("cdr_number") or fields.get("CDRNumber", "JOB")
        worksheet_folder_name = safe_folder_name(cdr_number)
        file_name = f"{safe_docx_filename(cdr_number)}_worksheet_{datetime.now(UK_TZ).strftime('%Y%m%d_%H%M%S')}.docx"

        worksheet_link = upload_file_to_sharepoint(
            site_id,
            WORKSHEET_BASE_FOLDER,
            worksheet_folder_name,
            file_name,
            docx_bytes,
        )
        return worksheet_link
    except Exception as e:
        print(f"ERROR generating worksheet Word document: {e}")
        return ""



def get_engineer_by_lookup_id(engineers, lookup_id):
    lookup_id = str(lookup_id or "")
    for engineer in engineers or []:
        fields = engineer.get("fields", {})
        item_lookup_id = str(fields.get("id", "") or engineer.get("id", ""))
        if item_lookup_id == lookup_id:
            return {
                "lookup_id": item_lookup_id,
                "name": str(get_field_value(fields, "EngineerName", "Engineer Name", "Title") or f"Engineer {lookup_id}"),
                "telegram_id": str(get_field_value(fields, "TelegramID", "Telegram ID") or "").strip(),
            }
    return None


def get_engineer_by_name(engineers, name):
    target = str(name or "").strip().lower()
    if not target:
        return None

    for engineer in engineers or []:
        fields = engineer.get("fields", {})
        engineer_name = str(get_field_value(fields, "EngineerName", "Engineer Name", "Title") or "").strip()
        if engineer_name.lower() == target:
            lookup_id = str(fields.get("id", "") or engineer.get("id", ""))
            return {
                "lookup_id": lookup_id,
                "name": engineer_name,
                "telegram_id": str(get_field_value(fields, "TelegramID", "Telegram ID") or "").strip(),
            }

    return None


def apply_manual_visit_times_to_log(fields, worksheet, outcome):
    """Return an EngineerVisitLog that uses the engineer-approved final review times.

    If the engineer edits times, add a Travelling Reverted marker and then clean
    Travelling / On Site / final outcome lines using the approved times. This lets
    the worksheet and timesheet parse the corrected attendance without office review.
    """
    existing_log = get_field_value(fields, "EngineerVisitLog", "Engineer Visit Log") or ""
    engineer_name = worksheet.get("engineer_name", "Engineer")
    manual_times = worksheet.get("EditedVisitTimes") or {}
    comment = build_visit_comment_extra(worksheet)

    if not manual_times:
        return append_engineer_log(fields, engineer_name, outcome, comment)

    today = datetime.now(UK_TZ).strftime("%d/%m/%Y")
    travel = normalise_helpdesk_time(manual_times.get("travel", "")) or manual_times.get("travel", "")
    on_site = normalise_helpdesk_time(manual_times.get("on_site", "")) or manual_times.get("on_site", "")
    off_site = normalise_helpdesk_time(manual_times.get("off_site", "")) or manual_times.get("off_site", "") or datetime.now(UK_TZ).strftime("%H:%M")

    lines = [line for line in str(existing_log).splitlines() if line.strip()]
    lines.append(f"{now_log_time()} - {engineer_name} - Travelling Reverted - Times manually edited on final worksheet review")
    if travel:
        lines.append(f"{today} {travel} - {engineer_name} - Travelling - Time manually confirmed on worksheet")
    if on_site:
        lines.append(f"{today} {on_site} - {engineer_name} - On Site - Time manually confirmed on worksheet")
    lines.append(f"{today} {off_site} - {engineer_name} - {outcome} - {comment}")
    return "\n".join(lines)



def build_worksheet_update_fields(worksheet, fields, updated_log, outcome, is_final_engineer, worksheet_pdf_link=""):
    fields_to_update = {
        "WorkCompleted": worksheet.get("WorkCompleted", ""),
        "MaterialsUsed": worksheet.get("MaterialsUsed", ""),
        "FollowOnRequired": worksheet.get("FollowOnRequired", False),
        "FollowOnNotes": worksheet.get("FollowOnNotes", ""),
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


async def finalise_worksheet_direct(context, worksheet, fields, item_id, site_id, jobs_list_id, actor_id):
    """Finalise an engineer worksheet without Helpdesk approval.

    This is deliberately small and keeps the existing shared worksheet/update helpers.
    The key difference from the old office-review flow is that EngineerVisitLog gets
    the actual final outcome (Completed / Revisit Required / No Access), so dashboard
    productivity can parse the visit as closed and keep the productive time.
    """
    outcome = worksheet.get("JobOutcome", "Completed")
    engineer_lookup_id = str(worksheet.get("engineer_lookup_id", ""))
    assigned_ids_latest = get_assigned_engineer_ids(fields)

    if engineer_lookup_id not in assigned_ids_latest:
        return False, "You are no longer assigned to this job. Worksheet has not been submitted.", ""

    is_final_engineer = len(get_remaining_assigned_engineer_ids(fields, engineer_lookup_id)) == 0
    updated_log = append_engineer_log(
        fields,
        worksheet["engineer_name"],
        outcome,
        build_visit_comment_extra(worksheet),
    )

    worksheet_link = ""
    # Generate the worksheet when this attendance closes the job. For revisit/no-access
    # the job returns to dispatch, but the visit remains in the EngineerVisitLog.
    if is_final_engineer and outcome == "Completed":
        worksheet_link = generate_and_upload_worksheet_pdf(
            site_id,
            jobs_list_id,
            item_id,
            worksheet,
            fields,
            updated_log,
            outcome,
        )

    update_payload = build_worksheet_update_fields(
        worksheet,
        fields,
        updated_log,
        outcome,
        is_final_engineer,
        worksheet_link,
    )


    update_list_item_fields_in_safe_chunks(
        site_id,
        jobs_list_id,
        item_id,
        update_payload,
        required_fields={"Status", "JobOutcome", "EngineerVisitLog"},
    )

    update_active_day_live_status(
        site_id,
        str(actor_id),
        outcome,
        get_job_reference(fields),
    )

    if outcome in ["Completed", "No Access", "Revisit Required"]:
        await notify_trade_group(context, worksheet, fields, updated_log, outcome)

    message = f"Worksheet submitted:\n\n{worksheet['cdr_number']} → {outcome}"
    if worksheet_link:
        message += "\nWorksheet generated."
    return True, message, worksheet_link


async def submit_worksheet_direct(update: Update, context: ContextTypes.DEFAULT_TYPE, sender):
    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await sender.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    site_id = worksheet["site_id"]
    jobs_list_id = worksheet["jobs_list_id"]
    item_id = worksheet["item_id"]

    latest_jobs = get_list_items(site_id, jobs_list_id)
    job = find_job_by_item_id(latest_jobs, item_id)
    fields = job["fields"] if job else worksheet["fields"]

    if is_closed_job(fields):
        context.user_data.pop("worksheet", None)
        await sender.reply_text(
            "This job has already been closed or returned to the office. Worksheet has not been submitted again.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    assigned_ids_latest = get_assigned_engineer_ids(fields)
    if worksheet["engineer_lookup_id"] not in assigned_ids_latest:
        context.user_data.pop("worksheet", None)
        await sender.reply_text(
            "You are no longer assigned to this job. Worksheet has not been submitted.",
            reply_markup=get_main_menu(await get_role_for_update(update)),
        )
        return ConversationHandler.END

    outcome = worksheet.get("JobOutcome", "Completed")
    updated_log = apply_manual_visit_times_to_log(fields, worksheet, outcome)

    is_final_engineer = len(get_remaining_assigned_engineer_ids(fields, worksheet["engineer_lookup_id"])) == 0

    worksheet_link = ""
    if is_final_engineer and outcome == "Completed":
        worksheet_link = generate_and_upload_worksheet_pdf(
            site_id,
            jobs_list_id,
            item_id,
            worksheet,
            fields,
            updated_log,
            outcome,
        )

    update_payload = build_worksheet_update_fields(
        worksheet,
        fields,
        updated_log,
        outcome,
        is_final_engineer,
        worksheet_link,
    )

    # Office approval has been removed from the engineer completion path. Do not
    # write the large approval payload here; keeping the direct submission
    # small avoids SharePoint Invalid request issues.

    update_list_item_fields_in_safe_chunks(
        site_id,
        jobs_list_id,
        item_id,
        update_payload,
        required_fields={"EngineerVisitLog", "WorksheetSubmitted"},
    )

    if outcome in ["Completed", "No Access", "Revisit Required"]:
        await notify_trade_group(context, worksheet, fields, updated_log, outcome)

    context.user_data.pop("worksheet", None)

    if is_final_engineer:
        if outcome == "Completed":
            message = f"Worksheet submitted and job completed:\n\n{worksheet['cdr_number']}"
            if worksheet_link:
                message += "\n\nWorksheet generated."
        else:
            message = f"Worksheet submitted and job returned to office for dispatch:\n\n{worksheet['cdr_number']} → {outcome}"
    else:
        message = f"Your worksheet has been submitted for {worksheet['cdr_number']}. Other assigned engineer(s) remain on the job."

    await sender.reply_text(message, reply_markup=get_main_menu(await get_role_for_update(update)))
    return ConversationHandler.END


async def worksheet_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    menu_result = await handle_menu_during_conversation(update, context, REVIEW)
    if menu_result is not None:
        return menu_result

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await update.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    edit_mode = context.user_data.pop("worksheet_review_edit", "")
    text = str(update.message.text or "").strip()
    answer = text.upper()

    if edit_mode == "times":
        parts = [part.strip() for part in text.replace(";", ",").split(",")]
        if len(parts) != 3:
            await update.message.reply_text("Please reply with three times like: 08:15, 08:45, 15:30")
            context.user_data["worksheet_review_edit"] = "times"
            return REVIEW
        travel, on_site, off_site = [normalise_helpdesk_time(part) for part in parts]
        if not travel or not on_site or not off_site:
            await update.message.reply_text("One of those times was not recognised. Please use 24-hour format, for example: 08:15, 08:45, 15:30")
            context.user_data["worksheet_review_edit"] = "times"
            return REVIEW
        worksheet["EditedVisitTimes"] = {"travel": travel, "on_site": on_site, "off_site": off_site}
        await update.message.reply_text(build_review_text(worksheet), reply_markup=get_review_keyboard())
        return REVIEW

    if edit_mode == "work":
        worksheet["WorkCompleted"] = text
        await update.message.reply_text(build_review_text(worksheet), reply_markup=get_review_keyboard())
        return REVIEW

    if edit_mode == "materials":
        worksheet["MaterialsUsed"] = text
        await update.message.reply_text(build_review_text(worksheet), reply_markup=get_review_keyboard())
        return REVIEW

    if edit_mode == "follow_on":
        lower = text.lower()
        if lower in ["no", "n", "none", "n/a", "na"]:
            worksheet["FollowOnRequired"] = False
            worksheet["FollowOnNotes"] = ""
        else:
            worksheet["FollowOnRequired"] = True
            worksheet["FollowOnNotes"] = text
        await update.message.reply_text(build_review_text(worksheet), reply_markup=get_review_keyboard())
        return REVIEW

    if edit_mode == "no_access_reason":
        worksheet["NoAccessReason"] = text
        await update.message.reply_text(build_review_text(worksheet), reply_markup=get_review_keyboard())
        return REVIEW

    if answer == "CANCEL":
        context.user_data.pop("worksheet", None)
        await update.message.reply_text("Worksheet cancelled. Nothing has been submitted.")
        return ConversationHandler.END

    if answer == "RESTART":
        worksheet["WorkCompleted"] = ""
        worksheet["MaterialsUsed"] = ""
        worksheet["FollowOnRequired"] = False
        worksheet["FollowOnNotes"] = ""
        worksheet["photo_links"] = []
        worksheet["photo_files_for_group"] = []
        worksheet["ClientSignatureRequired"] = False
        worksheet["ClientSignatureReceived"] = False
        worksheet.pop("EditedVisitTimes", None)
        await update.message.reply_text(f"Restarting worksheet for {worksheet['cdr_number']}.\n\nWhat work was completed?")
        return WORK_COMPLETED

    if answer == "SUBMIT":
        return await submit_worksheet_direct(update, context, update.message)

    await update.message.reply_text(
        "Please use the buttons below, or type SUBMIT, RESTART or CANCEL.",
        reply_markup=get_review_keyboard(),
    )
    return REVIEW


async def worksheet_review_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
        worksheet["photo_links"] = []
        worksheet["photo_files_for_group"] = []
        worksheet["ClientSignatureRequired"] = False
        worksheet["ClientSignatureReceived"] = False
        worksheet.pop("EditedVisitTimes", None)
        await query.message.reply_text(f"Restarting worksheet for {worksheet['cdr_number']}.\n\nWhat work was completed?")
        return WORK_COMPLETED

    if action == "submit":
        return await submit_worksheet_direct(update, context, query.message)

    if action == "edit_outcome":
        await query.message.reply_text("Select the correct outcome:", reply_markup=get_outcome_edit_keyboard())
        return REVIEW

    if action == "edit_times":
        current = get_current_visit_times_from_worksheet(worksheet)
        context.user_data["worksheet_review_edit"] = "times"
        await query.message.reply_text(
            "Reply with the correct travel, on site and off site times in this format:\n\n"
            "08:15, 08:45, 15:30\n\n"
            f"Current: {current.get('travel') or 'Missing'}, {current.get('on_site') or 'Missing'}, {current.get('off_site') or 'Missing'}"
        )
        return REVIEW

    if action == "edit_work":
        context.user_data["worksheet_review_edit"] = "work"
        await query.message.reply_text("Type the corrected work completed / comments.")
        return REVIEW

    if action == "edit_materials":
        context.user_data["worksheet_review_edit"] = "materials"
        await query.message.reply_text("Type the corrected materials used, or None.")
        return REVIEW

    if action == "edit_follow_on":
        context.user_data["worksheet_review_edit"] = "follow_on"
        await query.message.reply_text("Type the follow-on required, or type No if no follow-on is required.")
        return REVIEW

    if action == "add_photos":
        await query.message.reply_text(
            "Send any extra job photos now, then type DONE when finished."
        )
        return PHOTOS

    if action == "edit_signature":
        await query.message.reply_text("Is a client signature required?", reply_markup=get_yes_no_keyboard("signature_required"))
        return SIGNATURE_REQUIRED

    await query.message.reply_text("Please use one of the final review buttons.", reply_markup=get_review_keyboard())
    return REVIEW


async def worksheet_outcome_edit_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    worksheet = context.user_data.get("worksheet")
    if not worksheet:
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu(await get_role_for_update(update)))
        return ConversationHandler.END

    outcome = query.data.split("|", 1)[1]
    worksheet["JobOutcome"] = outcome

    if outcome == "No Access" and not worksheet.get("NoAccessReason"):
        context.user_data["worksheet_review_edit"] = "no_access_reason"
        await query.message.reply_text("Type the No Access reason / notes.")
        return REVIEW

    await query.message.reply_text(build_review_text(worksheet), reply_markup=get_review_keyboard())
    return REVIEW


async def worksheet_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

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
                        reply_markup=get_job_buttons(item_id, fields.get("SiteName", "")),
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


async def remind_engineers_to_start_day(app):
    """Daily 07:40 reminder to active engineers/admin engineers to start their day."""
    if is_weekend_today():
        print("Start-day reminders skipped because today is Saturday/Sunday.")
        return

    try:
        site_id = get_site_id()
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        engineers = get_list_items(site_id, engineers_list_id)

        for engineer in engineers:
            fields = engineer.get("fields", {})
            engineer_name = str(get_field_value(fields, "EngineerName", "Engineer Name", "Title") or "").strip()
            telegram_id = str(get_field_value(fields, "TelegramID", "Telegram ID") or "").strip()
            role = str(get_field_value(fields, "Role") or "Engineer").strip().lower()
            active_value = get_field_value(fields, "Active")

            if active_value not in [None, ""] and not bool_field(active_value):
                continue

            if role not in ["engineer", "admin", "apprentice"]:
                continue

            if not telegram_id:
                continue

            try:
                await app.bot.send_message(
                    chat_id=telegram_id,
                    text=(
                        "⏰ Start Day Reminder\n\n"
                        f"Morning {engineer_name or 'Engineer'}, please tap 🟢 Start Day when you begin work today."
                    ),
                )
            except Exception as e:
                print(f"WARNING: Could not send start-day reminder to {telegram_id}: {e}")

    except Exception as e:
        print(f"ERROR sending start-day reminders: {e}")


async def remind_active_engineers_to_end_day(app):
    """Daily reminder so engineers do not forget to close their day."""
    if is_weekend_today():
        print("End-day reminders skipped because today is Saturday/Sunday.")
        return

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
                        reply_markup=get_engineer_menu(),
                    )
                except Exception as e:
                    print(f"WARNING: Could not send end-day reminder to {telegram_id}: {e}")

    except Exception as e:
        print(f"ERROR sending end-day reminders: {e}")



async def check_engineer_idle_alerts(app):
    """Optional helpdesk idle alerts.

    Safe by default:
    - Does not write to SharePoint.
    - Does not message engineers.
    - Only sends to HELPDESK_CHAT_ID when IDLE_ALERTS_ENABLED=true.
    """
    if not IDLE_ALERTS_ENABLED:
        return

    if not HELPDESK_CHAT_ID:
        print("Idle alerts enabled but HELPDESK_CHAT_ID is missing.")
        return

    try:
        site_id = get_site_id()
        engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
        jobs_list_id = get_list_id(site_id, JOBS_LIST)
        day_logs_list_id = get_list_id(site_id, DAY_LOGS_LIST)

        engineers = get_list_items(site_id, engineers_list_id)
        jobs_data = get_list_items(site_id, jobs_list_id)
        day_logs = get_list_items(site_id, day_logs_list_id)
        rows = dashboard_engineer_rows(engineers, day_logs, jobs_data)
        now = datetime.now(UK_TZ)

        active_idle_names = set()

        for row in rows:
            if row.get("css") != "idle":
                continue

            idle_minutes = int(row.get("idle_minutes") or 0)
            if idle_minutes < IDLE_ALERT_THRESHOLD_MINUTES:
                continue

            engineer_name = row.get("name", "Engineer")
            active_idle_names.add(engineer_name)

            last_alert = LAST_IDLE_ALERTS.get(engineer_name)
            if last_alert:
                minutes_since_last = (now - last_alert).total_seconds() / 60
                if minutes_since_last < IDLE_ALERT_REPEAT_MINUTES:
                    continue

            LAST_IDLE_ALERTS[engineer_name] = now
            await app.bot.send_message(
                chat_id=HELPDESK_CHAT_ID,
                text=(
                    "⚠️ Engineer Idle Alert\n\n"
                    f"Engineer: {engineer_name}\n"
                    f"Idle time: {dashboard_duration_text(now.replace(tzinfo=UK_TZ) - __import__('datetime').timedelta(minutes=idle_minutes), now)}\n"
                    f"Last activity: {row.get('last', '-') or '-'}\n"
                    f"Status: {row.get('job', 'No current open job')}\n\n"
                    "This is a dashboard-only alert. No job, worksheet or SharePoint record has been changed."
                ),
            )

        # Clear alert state once the engineer is no longer idle, so future idle periods can alert again.
        for name in list(LAST_IDLE_ALERTS.keys()):
            if name not in active_idle_names:
                LAST_IDLE_ALERTS.pop(name, None)

    except Exception as e:
        print(f"ERROR checking engineer idle alerts: {e}")



MESSAGE_TEMPLATES = {
    "start_day": "Please remember to start your day on the bot.",
    "log_job": "Please remember to log onto your current job on the bot.",
    "end_day": "Please remember to end your day on the bot when finished.",
    "call_office": "Please call the office when you are free.",
    "upload_photos": "Please upload the missing job photos when you have signal.",
    "complete_worksheet": "Please complete your worksheet for the job before leaving site.",
    "custom": "",
}


def get_message_template_keyboard():
    return ReplyKeyboardMarkup(
        [
            ["Start Day", "Log Job"],
            ["End Day", "Call Office"],
            ["Upload Photos", "Complete Worksheet"],
            ["Custom"],
            ["/cancel"],
        ],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def normalise_message_template(value):
    return {
        "start day": "start_day",
        "log job": "log_job",
        "end day": "end_day",
        "call office": "call_office",
        "upload photos": "upload_photos",
        "complete worksheet": "complete_worksheet",
        "custom": "custom",
    }.get(str(value or "").strip().lower())


async def start_message_engineer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    if not user_can_use_helpdesk(role):
        await update.message.reply_text("This option is only available to Helpdesk/Admin users.")
        return ConversationHandler.END

    site_id = get_site_id()
    engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
    engineers = get_list_items(site_id, engineers_list_id)
    assignable = get_active_assignable_engineers(engineers)
    context.user_data["message_engineers"] = assignable

    await update.message.reply_text(
        "Who do you want to message? Reply with engineer number(s), for example 1 or 1,3.\n\n"
        + format_engineer_selection_list(assignable),
        reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
    )
    return MESSAGE_ENGINEER_SELECT


async def message_engineer_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    engineers = context.user_data.get("message_engineers", [])
    selected, error = parse_engineer_selection(update.message.text, engineers)
    if error:
        await update.message.reply_text(error)
        return MESSAGE_ENGINEER_SELECT

    context.user_data["message_engineer_selected"] = selected
    await update.message.reply_text("Choose a template or select Custom.", reply_markup=get_message_template_keyboard())
    return MESSAGE_ENGINEER_TEMPLATE


async def message_engineer_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    template_key = normalise_message_template(update.message.text)
    if not template_key:
        await update.message.reply_text("Please choose one of the template buttons.")
        return MESSAGE_ENGINEER_TEMPLATE

    if template_key == "custom":
        await update.message.reply_text("Type the custom message you want to send.", reply_markup=ReplyKeyboardRemove())
        return MESSAGE_ENGINEER_TEXT

    context.user_data["message_engineer_text"] = MESSAGE_TEMPLATES[template_key]
    selected = context.user_data.get("message_engineer_selected", [])
    names = ", ".join(e["name"] for e in selected)
    await update.message.reply_text(
        f"Review message to: {names}\n\n{MESSAGE_TEMPLATES[template_key]}\n\nReply YES to send or NO to cancel.",
        reply_markup=get_review_reply_keyboard(),
    )
    return MESSAGE_ENGINEER_REVIEW


async def message_engineer_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = str(update.message.text or "").strip()
    if not message:
        await update.message.reply_text("Please type a message.")
        return MESSAGE_ENGINEER_TEXT

    context.user_data["message_engineer_text"] = message
    selected = context.user_data.get("message_engineer_selected", [])
    names = ", ".join(e["name"] for e in selected)
    await update.message.reply_text(
        f"Review message to: {names}\n\n{message}\n\nReply YES to send or NO to cancel.",
        reply_markup=get_review_reply_keyboard(),
    )
    return MESSAGE_ENGINEER_REVIEW


async def message_engineer_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = str(update.message.text or "").strip().lower()
    if answer in ["❌ no", "no", "n", "cancel"]:
        await update.message.reply_text("Message cancelled.", reply_markup=get_helpdesk_menu(include_engineer_menu=True))
        return ConversationHandler.END

    if answer not in ["✅ yes", "yes", "y"]:
        await update.message.reply_text("Reply YES to send or NO to cancel.")
        return MESSAGE_ENGINEER_REVIEW

    selected = context.user_data.get("message_engineer_selected", [])
    message = context.user_data.get("message_engineer_text", "")
    sent = []
    failed = []

    for engineer in selected:
        try:
            await context.bot.send_message(
                chat_id=engineer["telegram_id"],
                text=f"📢 Message from CDR Helpdesk\n\n{message}",
            )
            sent.append(engineer["name"])
        except Exception as e:
            failed.append(f"{engineer['name']}: {e}")

    if HELPDESK_CHAT_ID:
        try:
            await context.bot.send_message(
                chat_id=HELPDESK_CHAT_ID,
                text=f"📢 Bot message sent to: {', '.join(sent) or 'None'}\n\nMessage:\n{message}",
            )
        except Exception:
            pass

    reply = f"Message sent to: {', '.join(sent) or 'None'}"
    if failed:
        reply += "\n\nFailed:\n" + "\n".join(failed)

    await update.message.reply_text(reply, reply_markup=get_helpdesk_menu(include_engineer_menu=True))
    return ConversationHandler.END


async def send_due_task_activity_reminders(app):
    """Send 08:00 reminders for stored Task / Activity items due today.

    Requires SharePoint list: Task Activities.
    Safe: only reads Task Activities and marks ReminderSent on that task record.
    Does not touch jobs, worksheets, photos, signatures or day logs.
    """
    if is_weekend_today():
        print("Task/activity reminders skipped because today is Saturday/Sunday.")
        return

    try:
        site_id = get_site_id()
        tasks_list_id = get_list_id(site_id, TASK_ACTIVITIES_LIST)
        tasks = get_list_items(site_id, tasks_list_id)
        today = get_today_iso()

        for task in tasks:
            fields = task.get("fields", {})
            status = str(get_field_value(fields, "Status") or "Active").strip().lower()
            if status != "active":
                continue

            if bool_field(get_field_value(fields, "ReminderSent", "Reminder Sent")):
                continue

            raw_date = get_field_value(fields, "TaskDate", "Task Date", "Date")
            parsed_date = sharepoint_date_to_uk_date(raw_date)
            task_date = parsed_date.isoformat() if parsed_date else str(raw_date)[:10]
            if task_date != today:
                continue

            telegram_ids_raw = str(get_field_value(fields, "AssignedTelegramIDs", "Assigned Telegram IDs") or "")
            telegram_ids = [x.strip() for x in re.split(r"[\n,;]+", telegram_ids_raw) if x.strip()]
            if not telegram_ids:
                continue

            title = str(get_field_value(fields, "Title") or "Task / Activity")
            time_text = str(get_field_value(fields, "TaskTime", "Task Time") or "")
            details = str(get_field_value(fields, "TaskDetails", "Task Details") or "")

            for telegram_id in telegram_ids:
                try:
                    await app.bot.send_message(
                        chat_id=telegram_id,
                        text=(
                            "📋 Task / Activity Reminder - Today\n\n"
                            f"{title}\n"
                            f"Date: {format_sharepoint_date(raw_date) or task_date}\n"
                            f"Time: {time_text or 'N/A'}\n\n"
                            f"{details}"
                        ),
                    )
                except Exception as e:
                    print(f"WARNING: Could not send task/activity reminder to {telegram_id}: {e}")

            update_list_item_fields(
                site_id,
                tasks_list_id,
                task.get("id"),
                build_field_payload_for_list(site_id, tasks_list_id, {"ReminderSent": True, "Reminder Sent": True}),
            )

    except Exception as e:
        print(f"ERROR sending task/activity reminders: {e}")





def task_date_iso_from_fields(fields):
    raw_date = get_field_value(fields, "TaskDate", "Task Date", "Date")
    parsed_date = sharepoint_date_to_uk_date(raw_date)
    if parsed_date:
        return parsed_date.isoformat()
    return str(raw_date or "")[:10]


def task_is_active_or_future(fields):
    status = str(get_field_value(fields, "Status") or "Active").strip().lower()
    if status != "active":
        return False

    task_date = task_date_iso_from_fields(fields)
    today = get_today_iso()

    # Show current and future active tasks for cancellation.
    # Old tasks stay hidden even if their status was not manually completed.
    return bool(task_date) and task_date >= today


def task_assigned_names_from_fields(fields):
    """Return assigned names for a stored Task / Activity record.

    The previous version accidentally called task_assigned_names_from_fields()
    from inside itself, causing infinite recursion when the Cancel Task / Activity
    button tried to display active tasks.
    """
    assigned = str(get_field_value(fields, "AssignedNames", "Assigned Names") or "").strip()
    if assigned:
        return assigned

    # Older records may only have Telegram IDs. Show a useful fallback rather than blank.
    telegram_ids = task_telegram_ids_from_fields(fields)
    if telegram_ids:
        return ", ".join(telegram_ids)

    return ""


def format_task_activity_option(index, task):
    fields = task.get("fields", {})
    title = str(get_field_value(fields, "Title") or "Task / Activity")
    task_date = task_date_iso_from_fields(fields)
    time_text = str(get_field_value(fields, "TaskTime", "Task Time") or "")
    assigned = task_assigned_names_from_fields(fields)
    details = str(get_field_value(fields, "TaskDetails", "Task Details") or "").strip()

    line = f"{index}. {title} | {task_date or 'No date'}"
    if time_text:
        line += f" | {time_text}"
    if assigned:
        line += f" | {assigned}"
    if details:
        short_details = details.replace("\\n", " ").replace("\n", " ")
        if len(short_details) > 70:
            short_details = short_details[:67] + "..."
        line += f"\n   {short_details}"
    return line


def task_telegram_ids_from_fields(fields):
    telegram_ids_raw = str(get_field_value(fields, "AssignedTelegramIDs", "Assigned Telegram IDs") or "")
    return [x.strip() for x in re.split(r"[\n,;]+", telegram_ids_raw) if x.strip()]


async def start_cancel_task_activity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_role_for_update(update)
    if not user_can_use_helpdesk(role):
        await update.message.reply_text("This option is only available to Helpdesk/Admin users.")
        return ConversationHandler.END

    try:
        site_id = get_site_id()
        tasks_list_id = get_list_id(site_id, TASK_ACTIVITIES_LIST)
        tasks = get_list_items(site_id, tasks_list_id)
    except Exception as e:
        await update.message.reply_text(
            "I could not open the Task Activities list. Check the SharePoint list exists and is named exactly: Task Activities"
        )
        print(f"ERROR opening Task Activities list for cancellation: {e}")
        return ConversationHandler.END

    active_tasks = [task for task in tasks if task_is_active_or_future(task.get("fields", {}))]
    active_tasks.sort(key=lambda t: (task_date_iso_from_fields(t.get("fields", {})), str(get_field_value(t.get("fields", {}), "TaskTime", "Task Time") or "")))

    if not active_tasks:
        await update.message.reply_text("There are no active current/future Task / Activity items to cancel.")
        return ConversationHandler.END

    context.user_data["cancel_task_site_id"] = site_id
    context.user_data["cancel_task_list_id"] = tasks_list_id
    context.user_data["cancel_task_items"] = active_tasks

    options = "\n\n".join(format_task_activity_option(i, task) for i, task in enumerate(active_tasks, start=1))
    await update.message.reply_text(
        "Which Task / Activity do you want to cancel? Reply with the number.\n\n" + options,
        reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
    )
    return CANCEL_TASK_SELECT


async def cancel_task_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text_value = str(update.message.text or "").strip()
    tasks = context.user_data.get("cancel_task_items", [])

    if not text_value.isdigit():
        await update.message.reply_text("Please reply with the task number.")
        return CANCEL_TASK_SELECT

    index = int(text_value)
    if index < 1 or index > len(tasks):
        await update.message.reply_text("That number is not in the list.")
        return CANCEL_TASK_SELECT

    selected = tasks[index - 1]
    context.user_data["cancel_task_selected"] = selected

    await update.message.reply_text(
        "Please type the cancellation reason. This will be sent to the assigned person/people.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return CANCEL_TASK_REASON


async def cancel_task_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    reason = str(update.message.text or "").strip()
    if not reason:
        await update.message.reply_text("Please type a cancellation reason.")
        return CANCEL_TASK_REASON

    context.user_data["cancel_task_reason"] = reason
    task = context.user_data.get("cancel_task_selected", {})
    fields = task.get("fields", {})

    title = str(get_field_value(fields, "Title") or "Task / Activity")
    task_date = task_date_iso_from_fields(fields)
    time_text = str(get_field_value(fields, "TaskTime", "Task Time") or "")
    assigned = task_assigned_names_from_fields(fields)

    await update.message.reply_text(
        "Review cancellation:\n\n"
        f"Task: {title}\n"
        f"Date: {task_date or 'N/A'}\n"
        f"Time: {time_text or 'N/A'}\n"
        f"Assigned: {assigned or 'N/A'}\n\n"
        f"Reason:\n{reason}\n\n"
        "Reply YES to cancel and notify, or NO to stop.",
        reply_markup=get_review_reply_keyboard(),
    )
    return CANCEL_TASK_REVIEW


async def cancel_task_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = str(update.message.text or "").strip().lower()
    if answer in ["❌ no", "no", "n", "cancel"]:
        await update.message.reply_text("Task / Activity cancellation stopped.", reply_markup=get_helpdesk_menu(include_engineer_menu=True))
        return ConversationHandler.END

    if answer not in ["✅ yes", "yes", "y"]:
        await update.message.reply_text("Reply YES to cancel and notify, or NO to stop.")
        return CANCEL_TASK_REVIEW

    site_id = context.user_data.get("cancel_task_site_id")
    tasks_list_id = context.user_data.get("cancel_task_list_id")
    task = context.user_data.get("cancel_task_selected", {})
    reason = context.user_data.get("cancel_task_reason", "")
    fields = task.get("fields", {})
    task_id = task.get("id")

    title = str(get_field_value(fields, "Title") or "Task / Activity")
    raw_date = get_field_value(fields, "TaskDate", "Task Date", "Date")
    task_date = format_sharepoint_date(raw_date) or task_date_iso_from_fields(fields)
    time_text = str(get_field_value(fields, "TaskTime", "Task Time") or "")
    details = str(get_field_value(fields, "TaskDetails", "Task Details") or "")
    telegram_ids = task_telegram_ids_from_fields(fields)

    try:
        payload = build_field_payload_for_list(
            site_id,
            tasks_list_id,
            {
                "Status": "Cancelled",
                "CancelReason": reason,
                "Cancel Reason": reason,
                "CancelledDateTime": graph_datetime_now(),
                "Cancelled Date Time": graph_datetime_now(),
            },
        )
        update_list_item_fields(site_id, tasks_list_id, task_id, payload)
    except Exception as e:
        await update.message.reply_text(f"I could not update the task to Cancelled in SharePoint: {e}")
        return ConversationHandler.END

    sent = []
    failed = []
    for telegram_id in telegram_ids:
        try:
            await context.bot.send_message(
                chat_id=telegram_id,
                text=(
                    "❌ Task / Activity Cancelled\n\n"
                    "Please do not attend the following:\n\n"
                    f"{title}\n"
                    f"Date: {task_date or 'N/A'}\n"
                    f"Time: {time_text or 'N/A'}\n\n"
                    f"{details}\n\n"
                    f"Reason: {reason}"
                ),
            )
            sent.append(telegram_id)
        except Exception as e:
            failed.append(f"{telegram_id}: {e}")

    if HELPDESK_CHAT_ID:
        try:
            await context.bot.send_message(
                chat_id=HELPDESK_CHAT_ID,
                text=(
                    "❌ Task / Activity cancelled\n\n"
                    f"Task: {title}\n"
                    f"Date: {task_date or 'N/A'}\n"
                    f"Time: {time_text or 'N/A'}\n"
                    f"Reason: {reason}\n"
                    f"Notified: {len(sent)}"
                ),
            )
        except Exception:
            pass

    reply = f"Task / Activity cancelled and notification sent to {len(sent)} recipient(s)."
    if failed:
        reply += "\n\nFailed to notify:\n" + "\n".join(failed)

    await update.message.reply_text(reply, reply_markup=get_helpdesk_menu(include_engineer_menu=True))
    return ConversationHandler.END



# -----------------------------
# Admin-only OpenAI assistant
# -----------------------------
def extract_openai_response_text(response_json):
    """Extract plain text from the OpenAI Responses API response shape."""
    if not isinstance(response_json, dict):
        return ""

    if response_json.get("output_text"):
        return str(response_json.get("output_text") or "").strip()

    texts = []
    for item in response_json.get("output", []) or []:
        for content in item.get("content", []) or []:
            if content.get("type") in ["output_text", "text"] and content.get("text"):
                texts.append(str(content.get("text")))

    return "\n".join(texts).strip()


def call_openai_assistant(question):
    """Call OpenAI using a simple server-side API key stored in Railway."""
    if not OPENAI_API_KEY:
        return "OPENAI_API_KEY is missing in Railway variables. Add it, redeploy, then try again."

    question = str(question or "").strip()
    if not question:
        return "Please ask a question."

    payload = {
        "model": OPENAI_MODEL,
        "instructions": (
            "You are an internal assistant for CDR M&E Services Ltd. "
            "Give practical, concise answers for a UK building services / M&E contractor. "
            "Do not claim you have checked SharePoint, Telegram, Railway, email or live company systems unless that information is included in the user's message. "
            "If the user asks for legal, HR, financial, gas, electrical or compliance advice, give cautious general guidance and tell them to verify before acting. "
            "Avoid asking for personal data or customer-sensitive information."
        ),
        "input": question,
        "max_output_tokens": 700,
        "store": False,
    }

    try:
        response = requests.post(
            "https://api.openai.com/v1/responses",
            headers={
                "Authorization": f"Bearer {OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=45,
        )

        if response.status_code >= 400:
            return f"OpenAI request failed ({response.status_code}): {response.text[:900]}"

        answer = extract_openai_response_text(response.json())
        return answer or "OpenAI returned an empty response."

    except Exception as e:
        return f"OpenAI request failed: {e}"


async def ask_chatbot_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)
    if str(role or "").lower() != "admin":
        await update.message.reply_text(
            "Only Admin users can use Ask ChatGPT.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    context.user_data["ask_chatbot_active"] = True
    await update.message.reply_text(
        "Ask ChatGPT is open for Admin testing.\n\n"
        "Type your question now. Avoid sending customer names, tenant details, addresses or other GDPR-sensitive data while testing.\n\n"
        "Type /cancel to close.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ASK_CHATBOT_QUESTION


async def ask_chatbot_question(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    menu_result = await handle_menu_during_conversation(update, context, ASK_CHATBOT_QUESTION)
    if menu_result is not None:
        return menu_result

    role = await get_role_for_update(update)
    if str(role or "").lower() != "admin":
        context.user_data.pop("ask_chatbot_active", None)
        await update.message.reply_text(
            "Only Admin users can use Ask ChatGPT.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    question = str(update.message.text or "").strip()
    if not question:
        await update.message.reply_text("Please type a question, or /cancel to close.")
        return ASK_CHATBOT_QUESTION

    await update.message.reply_text("Thinking...")
    answer = await asyncio.to_thread(call_openai_assistant, question)

    # Telegram has a message length limit. Split long replies into safe chunks.
    chunks = [answer[i:i + 3500] for i in range(0, len(answer), 3500)] or ["No answer returned."]
    for chunk in chunks:
        await update.message.reply_text(chunk)

    await update.message.reply_text(
        "Ask another question, or type /cancel to close Ask ChatGPT.",
    )
    return ASK_CHATBOT_QUESTION


async def ask_chatbot_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    context.user_data.pop("ask_chatbot_active", None)
    role = await get_role_for_update(update)
    await update.message.reply_text(
        "Ask ChatGPT closed.",
        reply_markup=get_main_menu(role),
    )
    return ConversationHandler.END



GLOBAL_SCHEDULER = None


async def post_init(app):
    try:
        await app.bot.delete_webhook(drop_pending_updates=True)
        print("Webhook removed.")
    except Exception as e:
        print(f"Could not remove webhook: {e}")

    global GLOBAL_SCHEDULER
    scheduler = AsyncIOScheduler(timezone=UK_TZ)

    scheduler.add_job(
        send_new_jobs,
        trigger="interval",
        seconds=30,
        args=[app],
    )

    scheduler.add_job(
        remind_engineers_to_start_day,
        trigger="cron",
        day_of_week="mon-fri",
        hour=7,
        minute=40,
        args=[app],
    )

    scheduler.add_job(
        remind_active_engineers_to_end_day,
        trigger="cron",
        day_of_week="mon-fri",
        hour=16,
        minute=45,
        args=[app],
    )

    scheduler.add_job(
        send_due_task_activity_reminders,
        trigger="cron",
        day_of_week="mon-fri",
        hour=8,
        minute=0,
        args=[app],
    )

    scheduler.start()
    GLOBAL_SCHEDULER = scheduler
    print("Scheduler started.")


telegram_app = (
    ApplicationBuilder()
    .token(BOT_TOKEN)
    .connect_timeout(30)
    .read_timeout(60)
    .write_timeout(120)
    .pool_timeout(30)
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
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, startday_confirm),
        ],
        START_DAY_VAN_REG: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, startday_van_reg)],
        START_DAY_START_MILEAGE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, startday_start_mileage)],
        START_DAY_VAN_CHECK: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, startday_van_check)],
        START_DAY_VAN_PHOTOS: [
            MessageHandler(filters.ChatType.PRIVATE & filters.PHOTO, startday_van_photos),
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, startday_van_photos),
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
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, endday_confirm),
        ],
        END_DAY_MILEAGE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, endday_mileage)],
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
        WORK_COMPLETED: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_work_completed)],
        MATERIALS_USED: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_materials_used)],
        FOLLOW_ON_REQUIRED: [
            CallbackQueryHandler(worksheet_follow_on_required_button, pattern=r"^follow_on_required\|"),
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_follow_on_required),
        ],
        FOLLOW_ON_NOTES: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_follow_on_notes)],
        PHOTOS: [
            CallbackQueryHandler(worksheet_photos_done_button, pattern=r"^photos_done\|"),
            MessageHandler(filters.ChatType.PRIVATE & filters.PHOTO, worksheet_photos),
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_photos),
        ],
        SIGNATURE_REQUIRED: [
            CallbackQueryHandler(worksheet_signature_required_button, pattern=r"^signature_required\|"),
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_signature_required),
        ],
        SIGNATURE_WAITING: [
            CallbackQueryHandler(worksheet_signature_waiting_button, pattern=r"^signature_waiting\|"),
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_signature_waiting),
        ],
        REVIEW: [
            CallbackQueryHandler(worksheet_review_button, pattern=r"^review\|"),
            CallbackQueryHandler(worksheet_outcome_edit_button, pattern=r"^review_outcome\|"),
            MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, worksheet_review),
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
        BUG_IDEA_TEXT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, bugidea_text)],
    },
    fallbacks=[CommandHandler("cancel", bugidea_cancel)],
)



reopenjob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("reopenjob", reopenjob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_REOPEN_JOB)}$"), reopenjob_start),
    ],
    states={
        REOPENJOB_CDR_NUMBER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reopenjob_cdr_number)],
        REOPENJOB_DATE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reopenjob_date)],
        REOPENJOB_ASSIGN_ENGINEERS: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reopenjob_assign_engineers)],
        REOPENJOB_REASON: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reopenjob_reason)],
        REOPENJOB_REVIEW: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reopenjob_review)],
    },
    fallbacks=[CommandHandler("cancel", reopenjob_cancel)],
)


reassign_handler = ConversationHandler(
    entry_points=[
        CommandHandler("reassign", reassign_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_REASSIGN_JOB)}$"), reassign_start),
    ],
    states={
        REASSIGN_CDR_NUMBER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reassign_cdr_number)],
        REASSIGN_REMOVE_ENGINEERS: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reassign_remove_engineers)],
        REASSIGN_ASSIGN_ENGINEERS: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reassign_assign_engineers)],
        REASSIGN_DATE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reassign_date)],
        REASSIGN_REASON: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reassign_reason)],
        REASSIGN_REVIEW: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, reassign_review)],
    },
    fallbacks=[CommandHandler("cancel", reassign_cancel)],
)


openjobs_handler = ConversationHandler(
    entry_points=[
        CommandHandler("openjobs", openjobs_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_OPEN_JOBS)}$"), openjobs_start),
    ],
    states={
        OPENJOBS_FILTER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, openjobs_filter)],
        OPENJOBS_SELECT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, openjobs_select)],
    },
    fallbacks=[CommandHandler("cancel", openjobs_cancel)],
)



canceljob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("canceljob", canceljob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_CANCEL_JOB)}$"), canceljob_start),
    ],
    states={
        CANCELJOB_CDR_NUMBER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, canceljob_cdr_number)],
        CANCELJOB_CONFIRM: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, canceljob_confirm)],
    },
    fallbacks=[CommandHandler("cancel", canceljob_cancel)],
)


deletejob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("deletejob", deletejob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_DELETE_JOB)}$"), deletejob_start),
    ],
    states={
        DELETEJOB_CDR_NUMBER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, deletejob_cdr_number)],
        DELETEJOB_CONFIRM: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, deletejob_confirm)],
    },
    fallbacks=[CommandHandler("cancel", deletejob_cancel)],
)


receipt_handler = ConversationHandler(
    entry_points=[
        CommandHandler("receipts", receipt_start),
        CommandHandler("uploadreceipts", receipt_start),
        CommandHandler("returns", receipt_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_UPLOAD_RECEIPTS)}$"), receipt_start),
    ],
    states={
        RECEIPT_TYPE: [CallbackQueryHandler(receipt_type_button, pattern=r"^receipt_type\|")],
        RECEIPT_DATE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, receipt_date)],
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
        FINDJOB_SEARCH: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, findjob_search)],
        FINDJOB_SELECT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, findjob_select)],
    },
    fallbacks=[CommandHandler("cancel", findjob_cancel)],
)


def get_quote_reminder_recipients(site_id):
    engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
    engineers = get_list_items(site_id, engineers_list_id)
    recipients = []

    for item in engineers:
        fields = item.get("fields", {})
        active_value = get_field_value(fields, "Active")
        if active_value not in [None, ""] and not bool_field(active_value):
            continue

        name = str(get_field_value(fields, "EngineerName", "Engineer Name", "Title") or "").strip()
        telegram_id = str(get_field_value(fields, "TelegramID", "Telegram ID") or "").strip()
        role = str(get_field_value(fields, "Role") or "Engineer").strip()

        if name and telegram_id:
            recipients.append({"name": name, "telegram_id": telegram_id, "role": role})

    recipients.sort(key=lambda value: value["name"].lower())
    return recipients


def format_quote_recipient_list(recipients):
    return "\n".join(
        f"{index}. {recipient['name']} ({recipient.get('role', 'User')})"
        for index, recipient in enumerate(recipients, start=1)
    )


def parse_quote_recipient_selection(text, recipients):
    value = str(text or "").replace(";", ",").strip()
    selected = []
    seen_ids = set()

    if not value:
        return None, "Please reply with recipient number(s), e.g. 1 or 1,2."

    for part in [p.strip() for p in value.split(",") if p.strip()]:
        if not part.isdigit():
            return None, "Please reply with recipient number(s), e.g. 1 or 1,2."

        index = int(part)
        if index < 1 or index > len(recipients):
            return None, f"Recipient number {index} is not in the list."

        recipient = recipients[index - 1]
        recipient_id = str(recipient.get("telegram_id") or recipient.get("name") or index)

        if recipient_id not in seen_ids:
            selected.append(recipient)
            seen_ids.add(recipient_id)

    if not selected:
        return None, "Please select at least one recipient."

    return selected, ""


async def quote_reminder_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)

    if not user_can_use_helpdesk(role):
        await update.message.reply_text(
            "You do not have permission to create tasks / activities.",
            reply_markup=get_main_menu(role),
        )
        return ConversationHandler.END

    try:
        site_id = get_site_id()
        recipients = get_quote_reminder_recipients(site_id)
        if not recipients:
            await update.message.reply_text(
                "No active Telegram users were found to send the task / activity to.",
                reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
            )
            return ConversationHandler.END

        context.user_data["quote_reminder"] = {"role": role, "recipients": recipients}
        await update.message.reply_text(
            "Task / Activity.\n\nWho should this be sent to? Reply with the number.\n\n" +
            format_quote_recipient_list(recipients),
            reply_markup=ReplyKeyboardMarkup([["/cancel"]], resize_keyboard=True, one_time_keyboard=False),
        )
        return QUOTE_RECIPIENT

    except Exception as e:
        print(f"ERROR starting task / activity: {e}")
        await update.message.reply_text(
            "There was an error opening Task / Activity. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def quote_reminder_recipient(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    if not reminder:
        return await quote_reminder_start(update, context)

    recipients, error = parse_quote_recipient_selection(update.message.text, reminder.get("recipients", []))
    if error:
        await update.message.reply_text(error + "\n\n" + format_quote_recipient_list(reminder.get("recipients", [])))
        return QUOTE_RECIPIENT

    reminder["recipients_selected"] = recipients
    reminder["recipient"] = recipients[0]
    await update.message.reply_text("Enter the Task / Activity date. Example: Today, Tomorrow, 15/05/2026")
    return QUOTE_DATE



async def quote_reminder_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    if not reminder:
        return await quote_reminder_start(update, context)

    value = update.message.text.strip()
    parsed_date, date_display, time_from_line = parse_task_activity_datetime(value)
    if not parsed_date:
        await update.message.reply_text("Please enter a valid date. Example: Today, Tomorrow, 15/05/2026")
        return QUOTE_DATE

    reminder["date"] = parsed_date
    reminder["date_display"] = date_display or parsed_date
    if time_from_line:
        reminder["time"] = time_from_line
        await update.message.reply_text("Enter the client / site name.")
        return QUOTE_CLIENT

    await update.message.reply_text("Enter the client / site name.")
    return QUOTE_CLIENT

async def quote_reminder_client(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the client name.")
        return QUOTE_CLIENT

    reminder["client"] = value
    await update.message.reply_text("Enter the address / site to attend, or type None if not needed.")
    return QUOTE_ADDRESS


async def quote_reminder_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    value = update.message.text.strip()
    reminder["address"] = "" if is_blank_or_skip(value) else value
    await update.message.reply_text("Enter the time. Example: 08:00, 11:00, 14:00, ASAP")
    return QUOTE_TIME


async def quote_reminder_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the time. Example: 08:00, 11:00, 14:00, ASAP")
        return QUOTE_TIME

    normalised_time = normalise_helpdesk_time(value) or value
    reminder["time"] = normalised_time
    await update.message.reply_text("What is the task / activity details?")
    return QUOTE_SCOPE


async def quote_reminder_scope(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    value = update.message.text.strip()
    if is_blank_or_skip(value):
        await update.message.reply_text("Please enter the task / activity details.")
        return QUOTE_SCOPE

    reminder["scope"] = value
    await update.message.reply_text(
        build_quote_reminder_review(reminder),
        reply_markup=get_review_reply_keyboard(),
    )
    return QUOTE_REVIEW


def build_quote_reminder_review(reminder):
    recipient = reminder.get("recipient", {})
    return (
        "Please review this task / activity:\n\n"
        f"Send to: {recipient.get('name', '')}\n"
        f"Client: {reminder.get('client', '')}\n"
        f"Address: {reminder.get('address', '')}\n"
        f"Time: {reminder.get('time', '')}\n"
        f"Scope: {reminder.get('scope', '')}\n\n"
        "Reply YES to send it now, NO to cancel, or RESTART to start again."
    )


async def send_quote_reminder(bot, chat_id, reminder):
    await bot.send_message(
        chat_id=chat_id,
        text=(
            "📌 Task / Activity\n\n"
            f"Client: {reminder.get('client', '')}\n"
            f"Address: {reminder.get('address', '')}\n"
            f"Time: {reminder.get('time', '')}\n"
            f"Scope: {reminder.get('scope', '')}\n\n"
            f"Sent by: {reminder.get('set_by', 'Helpdesk')}"
        ),
    )


async def quote_reminder_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    reminder = context.user_data.get("quote_reminder")
    role = reminder.get("role", "Helpdesk") if reminder else "Helpdesk"
    answer = update.message.text.strip().lower()
    answer = answer.replace("✅", "").replace("❌", "").replace("🔄", "").strip()

    if answer in ["no", "n", "cancel"]:
        context.user_data.pop("quote_reminder", None)
        await update.message.reply_text(
            "Task / Activity cancelled.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END

    if answer in ["restart", "redo"]:
        context.user_data.pop("quote_reminder", None)
        return await quote_reminder_start(update, context)

    if answer not in ["yes", "y"]:
        await update.message.reply_text("Reply YES to send it, NO to cancel, or RESTART to start again.")
        return QUOTE_REVIEW

    try:
        recipients_to_send = reminder.get("recipients_selected") or [reminder.get("recipient", {})]
        reminder["set_by"] = update.effective_user.full_name or "Helpdesk"

        try:
            site_id = get_site_id()
            create_task_activity_record(site_id, reminder)
        except Exception as store_error:
            print(f"WARNING: Task / Activity was sent but could not be stored for reminder/cancellation: {store_error}")

        sent_names = []
        failed_names = []

        for recipient in recipients_to_send:
            chat_id = recipient.get("telegram_id")
            if not chat_id:
                failed_names.append(recipient.get("name", "Unknown"))
                continue

            try:
                await send_quote_reminder(context.bot, chat_id, reminder)
                sent_names.append(recipient.get("name", "recipient"))
            except Exception as send_error:
                failed_names.append(f"{recipient.get('name', 'Unknown')}: {send_error}")

        if not sent_names:
            raise Exception("No tasks / activities were sent. " + "; ".join(failed_names))

        context.user_data.pop("quote_reminder", None)
        message = f"Task / Activity sent to {', '.join(sent_names)}."
        if failed_names:
            message += f"\n\nFailed: {'; '.join(failed_names)}"

        await update.message.reply_text(
            message,
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR sending task / activity: {e}")
        await update.message.reply_text(
            "There was an error sending the task / activity. Please check Railway logs.",
            reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
        )
        return ConversationHandler.END


async def quote_reminder_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if is_group_chat(update):
        return ConversationHandler.END

    role = await get_role_for_update(update)
    context.user_data.pop("quote_reminder", None)
    await update.message.reply_text(
        "Task / Activity cancelled.",
        reply_markup=get_helpdesk_menu(include_engineer_menu=(role.lower() == "admin")),
    )
    return ConversationHandler.END


logjob_handler = ConversationHandler(
    entry_points=[
        CommandHandler("logjob", logjob_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_LOG_JOB)}$"), logjob_start),
    ],
    states={
        LOGJOB_CDR_NUMBER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_cdr_number)],
        LOGJOB_CUSTOMER_NAME: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_customer_name)],
        LOGJOB_SITE_NAME: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_site_name)],
        LOGJOB_CONTACT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_contact)],
        LOGJOB_TASK: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_task)],
        LOGJOB_NOTES: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_notes)],
        LOGJOB_DATE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_date)],
        LOGJOB_TIME: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_time)],
        LOGJOB_CATEGORY: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_category)],
        LOGJOB_ORDER_NUMBER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_order_number)],
        LOGJOB_ASSIGN_ENGINEERS: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_assign_engineers)],
        LOGJOB_REVIEW: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, logjob_review)],
    },
    fallbacks=[CommandHandler("cancel", logjob_cancel)],
)

cancel_task_activity_handler = ConversationHandler(
    entry_points=[
        CommandHandler("canceltask", start_cancel_task_activity),
        MessageHandler(filters.ChatType.PRIVATE & filters.Regex(f"^{re.escape(MENU_CANCEL_TASK_ACTIVITY)}$"), start_cancel_task_activity),
    ],
    states={
        CANCEL_TASK_SELECT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, cancel_task_select)],
        CANCEL_TASK_REASON: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, cancel_task_reason)],
        CANCEL_TASK_REVIEW: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, cancel_task_review)],
    },
    fallbacks=[CommandHandler("cancel", quote_reminder_cancel)],
)


message_engineer_handler = ConversationHandler(
    entry_points=[
        MessageHandler(filters.Regex(f"^{re.escape(MENU_MESSAGE_ENGINEER)}$"), start_message_engineer),
    ],
    states={
        MESSAGE_ENGINEER_SELECT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, message_engineer_select)],
        MESSAGE_ENGINEER_TEMPLATE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, message_engineer_template)],
        MESSAGE_ENGINEER_TEXT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, message_engineer_text)],
        MESSAGE_ENGINEER_REVIEW: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, message_engineer_review)],
    },
    fallbacks=[CommandHandler("cancel", quote_reminder_cancel)],
)


quote_reminder_handler = ConversationHandler(
    entry_points=[
        CommandHandler("quotereminder", quote_reminder_start),
        MessageHandler(filters.Regex(f"^{re.escape(MENU_QUOTE_REMINDER)}$"), quote_reminder_start),
    ],
    states={
        QUOTE_RECIPIENT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_recipient)],
        QUOTE_DATE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_date)],
        QUOTE_CLIENT: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_client)],
        QUOTE_ADDRESS: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_address)],
        QUOTE_TIME: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_time)],
        QUOTE_SCOPE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_scope)],
        QUOTE_REVIEW: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, quote_reminder_review)],
    },
    fallbacks=[CommandHandler("cancel", quote_reminder_cancel)],
)




asset_handler = ConversationHandler(
    entry_points=[
        CommandHandler("createasset", asset_start_manual),
        MessageHandler(filters.ChatType.PRIVATE & filters.Regex(f"^{re.escape(MENU_CREATE_ASSET)}$"), asset_start_manual),
        CallbackQueryHandler(asset_start_from_job, pattern=r"^asset_create_job\|"),
    ],
    states={
        ASSET_CUSTOMER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_customer)],
        ASSET_SITE_NAME: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_site_name)],
        ASSET_SITE_ADDRESS: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_site_address)],
        ASSET_CATEGORY: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_category)],
        ASSET_TYPE: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_type)],
        ASSET_DESCRIPTION: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_description)],
        ASSET_LOCATION: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_location)],
        ASSET_MANUFACTURER: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_manufacturer)],
        ASSET_MODEL: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_model)],
        ASSET_SERIAL: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_serial)],
        ASSET_PHOTO: [MessageHandler(filters.ChatType.PRIVATE & (filters.PHOTO | filters.TEXT) & ~filters.COMMAND, asset_photo)],
        ASSET_NOTES: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, asset_notes)],
    },
    fallbacks=[CommandHandler("cancel", asset_cancel)],
)


ask_chatbot_handler = ConversationHandler(
    entry_points=[
        MessageHandler(filters.Regex(f"^{re.escape(MENU_ASK_CHATBOT)}$"), ask_chatbot_start),
    ],
    states={
        ASK_CHATBOT_QUESTION: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, ask_chatbot_question)],
    },
    fallbacks=[CommandHandler("cancel", ask_chatbot_cancel)],
)

abortjob_handler = ConversationHandler(
    entry_points=[
        CallbackQueryHandler(abort_job_start, pattern="^abort_job\\|"),
    ],
    states={
        ABORTJOB_REASON: [CallbackQueryHandler(abort_job_reason, pattern="^abort_reason\\|")],
        ABORTJOB_NOTES: [MessageHandler(filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND, abort_job_notes)],
    },
    fallbacks=[CommandHandler("cancel", abort_job_cancel)],
)



telegram_app.add_handler(
    MessageHandler(
        filters.ChatType.GROUPS & (filters.COMMAND | filters.Regex(r"^(🟢 Start Day|📋 My Jobs|🏁 End Day|🐞 Bug / Ideas|🧾 Receipts / Returns|📣 Request Job|🧰 Helpdesk|➕ Log Job|🔁 Reassign Job|♻️ Reopen Job|📋 Open Jobs|🔎 Find Job|❌ Cancel Job|🗑 Delete Job|🤖 Ask ChatGPT|🏷 Create Asset|👷 Engineer Menu|📌 Task / Activity|📢 Message Engineer|❌ Cancel Task / Activity|/start|/my_id|/id|/jobs|/requestjob|/helpdesk|/startday|/endday|/receipts|/uploadreceipts|/logjob|/reassign|/reopenjob|/findjob|/openjobs|/canceljob|/deletejob|/quotereminder)$")),
        group_chat_cleanup,
    ),
    group=0,
)

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("my_id", my_id))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(CommandHandler("requestjob", request_job))
telegram_app.add_handler(CommandHandler("helpdesk", helpdesk_start))
telegram_app.add_handler(cancel_task_activity_handler)
telegram_app.add_handler(startday_handler)
telegram_app.add_handler(endday_handler)
telegram_app.add_handler(worksheet_handler)
telegram_app.add_handler(bugidea_handler)
telegram_app.add_handler(receipt_handler)
telegram_app.add_handler(logjob_handler)
telegram_app.add_handler(message_engineer_handler)
telegram_app.add_handler(quote_reminder_handler)
telegram_app.add_handler(asset_handler)
telegram_app.add_handler(ask_chatbot_handler)
telegram_app.add_handler(reopenjob_handler)
telegram_app.add_handler(reassign_handler)
telegram_app.add_handler(openjobs_handler)
telegram_app.add_handler(canceljob_handler)
telegram_app.add_handler(deletejob_handler)
telegram_app.add_handler(findjob_handler)
telegram_app.add_handler(abortjob_handler)
telegram_app.add_handler(MessageHandler(filters.Regex(f"^({MENU_MY_JOBS}|{MENU_BUG_IDEA}|{MENU_UPLOAD_RECEIPTS}|{MENU_REQUEST_JOB}|{MENU_CREATE_ASSET}|{MENU_QUOTE_REMINDER}|{MENU_MESSAGE_ENGINEER}|{MENU_CANCEL_TASK_ACTIVITY}|{MENU_HELPDESK}|{MENU_LOG_JOB}|{MENU_REASSIGN_JOB}|{MENU_REOPEN_JOB}|{MENU_OPEN_JOBS}|{MENU_FIND_JOB}|{MENU_CANCEL_JOB}|{MENU_DELETE_JOB}|{MENU_ASK_CHATBOT}|{MENU_ENGINEER_MENU})$"), menu_button))
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
