import os
import base64
import secrets
import threading
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

BOT_TOKEN = os.getenv("BOT_TOKEN")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
HELPDESK_CHAT_ID = os.getenv("HELPDESK_CHAT_ID")
SIGNATURE_BASE_URL = os.getenv("SIGNATURE_BASE_URL")
PORT = int(os.getenv("PORT", "8000"))

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"
DAY_LOGS_LIST = "Engineer Day Logs"
BUG_IDEAS_LIST = "Bug Ideas"


PHOTO_LIBRARY = "Documents"
PHOTO_BASE_FOLDER = "15 - ENGINEER JOB PHOTOS"
SIGNATURE_BASE_FOLDER = "16 - CLIENT SIGNATURES"
VAN_CHECK_PHOTO_BASE_FOLDER = "17 - VAN CHECK PHOTOS"

DAY_ACTIVE_STATUS = "Active"
DAY_CLOSED_STATUS = "Closed"

MENU_START_DAY = "🟢 Start Day"
MENU_MY_JOBS = "📋 My Jobs"
MENU_END_DAY = "🏁 End Day"
MENU_BUG_IDEA = "🐞 Bug / Ideas"


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


def get_main_menu():
    return ReplyKeyboardMarkup(
        [
            [MENU_START_DAY, MENU_MY_JOBS],
            [MENU_END_DAY, MENU_BUG_IDEA],
        ],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


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

    return HTMLResponse(
        """
        <html>
            <body style="font-family: Arial; padding: 30px;">
                <h2>Signature saved</h2>
                <p>Thank you. The job has been signed successfully.</p>
                <p>You can now hand the phone back to the engineer.</p>
            </body>
        </html>
        """
    )


def run_signature_web_server():
    uvicorn.run(web_app, host="0.0.0.0", port=PORT, log_level="info")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "CDR Engineer Bot is online. Use the menu below.",
        reply_markup=get_main_menu(),
    )


async def startday_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, _, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        day_logs_list_id = get_list_id(site_id, DAY_LOGS_LIST)
        day_logs = get_list_items(site_id, day_logs_list_id)
        active_day = find_active_day_log(day_logs, user_id)

        if active_day:
            await update.message.reply_text(
                "Your day is already active. You can now use 📋 My Jobs.",
                reply_markup=get_main_menu(),
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
            reply_markup=get_main_menu(),
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
        await update.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu())
        return ConversationHandler.END

    await update.message.reply_text("Starting your day. Please enter the van registration.")
    return START_DAY_VAN_REG


async def startday_confirm_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    answer = query.data.split("|", 1)[1]

    if answer == "no":
        context.user_data.pop("start_day", None)
        await query.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu())
        return ConversationHandler.END

    await query.message.reply_text("Starting your day. Please enter the van registration.")
    return START_DAY_VAN_REG


async def startday_van_reg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, START_DAY_VAN_REG)
    if menu_result is not None:
        return menu_result

    start_day = context.user_data.get("start_day")

    if not start_day:
        await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu())
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
        await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu())
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
            reply_markup=get_main_menu(),
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
        await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu())
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
            await update.message.reply_text("Please try /startday again.", reply_markup=get_main_menu())
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
                reply_markup=get_main_menu(),
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
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END


async def startday_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("start_day", None)
    await update.message.reply_text("Start day cancelled. Your jobs are still locked.", reply_markup=get_main_menu())
    return ConversationHandler.END


async def endday_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, _, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        day_logs_list_id, active_day = get_active_day_for_engineer(site_id, user_id)

        if not active_day:
            await update.message.reply_text(
                "You do not have an active day to end. Tap 🟢 Start Day when you begin work.",
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        _, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        open_jobs = get_open_jobs_for_engineer_today(jobs_data, current_engineer["lookup_id"])

        if open_jobs:
            await update.message.reply_text(
                "You cannot end your day while you still have job(s) assigned for today. "
                "Complete them, mark No Access, or mark Revisit Required first.\n\n"
                f"Open job(s):\n{format_open_jobs_for_end_day(open_jobs)}",
                reply_markup=get_main_menu(),
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
            reply_markup=get_main_menu(),
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
        await update.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu())
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
        await query.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu())
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
            await update.message.reply_text("Please try /endday again.", reply_markup=get_main_menu())
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
                reply_markup=get_main_menu(),
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
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR saving end day: {e}")
        await update.message.reply_text(
            "There was an error saving your end day record. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END


async def endday_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("end_day", None)
    await update.message.reply_text("End day cancelled. Your day is still active.", reply_markup=get_main_menu())
    return ConversationHandler.END


async def mystatus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        site_id, _, _, current_engineer = get_engineer_for_telegram_id(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID.",
                reply_markup=get_main_menu(),
            )
            return

        _, active_day = get_active_day_for_engineer(site_id, user_id)

        if active_day:
            fields = active_day["fields"]
            await update.message.reply_text(
                f"Status: Day active\n"
                f"Engineer: {current_engineer['name']}\n"
                f"Start time: {format_sharepoint_date(fields.get('StartTime', ''))} {str(fields.get('StartTime', ''))[11:16] if fields.get('StartTime') else ''}",
                reply_markup=get_main_menu(),
            )
        else:
            await update.message.reply_text(
                "Status: No active day. Tap 🟢 Start Day before using job buttons.",
                reply_markup=get_main_menu(),
            )

    except Exception as e:
        print(f"ERROR getting status: {e}")
        await update.message.reply_text("There was an error checking your status.", reply_markup=get_main_menu())


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


async def handle_menu_during_conversation(update: Update, context: ContextTypes.DEFAULT_TYPE, current_state):
    """
    Allows safe menu use while an engineer is inside a conversation flow.
    Without this, buttons like 🆔 My ID are treated as answers to the active question.
    """
    text = update.message.text if update.message else ""
    if text == MENU_BUG_IDEA:
        return current_state
    if text == MENU_MY_JOBS:
        await update.message.reply_text(
            "You are currently part-way through another task. Finish it or type /cancel before viewing jobs.",
            reply_markup=get_main_menu(),
        )
        return current_state

    if text in [MENU_START_DAY, MENU_END_DAY, MENU_BUG_IDEA]:
        await update.message.reply_text(
            "You are already part-way through another task. Finish it or type /cancel first.",
            reply_markup=get_main_menu(),
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
                reply_markup=get_main_menu(),
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
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END


async def bugidea_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, BUG_IDEA_TEXT)
    if menu_result is not None:
        return menu_result

    try:
        bug_idea = context.user_data.get("bug_idea")

        if not bug_idea:
            await update.message.reply_text("Please try again from the menu.", reply_markup=get_main_menu())
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
            reply_markup=get_main_menu(),
        )

        return ConversationHandler.END

    except Exception as e:
        print(f"ERROR saving bug/idea: {e}")
        await update.message.reply_text(
            "There was an error saving the bug/idea. Please ask the office to check Railway logs.",
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END


async def bugidea_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("bug_idea", None)
    await update.message.reply_text("Bug/idea cancelled.", reply_markup=get_main_menu())
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
                reply_markup=get_main_menu(),
            )
            return

        if not engineer_has_active_day(site_id, user_id):
            await update.message.reply_text(
                "Please start your day first using 🟢 Start Day or /startday. Your jobs are locked until your day has started.",
                reply_markup=get_main_menu(),
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
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        if not engineer_has_active_day(site_id, user_id):
            await sender.reply_text(
                "Please start your day first using 🟢 Start Day or /startday before updating jobs.",
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        job = find_job_by_item_id(jobs_data, item_id) if item_id else find_job_by_cdr(jobs_data, cdr_number)

        if not job:
            await sender.reply_text("Could not find this job. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu())
            return ConversationHandler.END

        fields = job["fields"]

        if is_closed_job(fields):
            await sender.reply_text(
                "This job has already been closed or returned to the office. No further action is required.",
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        assigned_ids = get_assigned_engineer_ids(fields)
        if current_engineer["lookup_id"] not in assigned_ids:
            await sender.reply_text("You are not assigned to this job.", reply_markup=get_main_menu())
            return ConversationHandler.END

        allowed, reason = can_click_action(fields, current_engineer["name"], outcome)
        if not allowed:
            await sender.reply_text(reason, reply_markup=get_main_menu())
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
        await target.reply_text("There was an error starting the worksheet.", reply_markup=get_main_menu())
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


async def complete_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text(
            "Use 📋 My Jobs, then tap Complete, Revisit or No Access on the job card.\n\n"
            "The old /complete CDR number method still works if needed: /complete CDR00001",
            reply_markup=get_main_menu(),
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

    context.user_data["worksheet"]["WorkCompleted"] = update.message.text
    await update.message.reply_text("What materials were used? Type None if none.")
    return MATERIALS_USED


async def worksheet_materials_used(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, MATERIALS_USED)
    if menu_result is not None:
        return menu_result

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
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu())
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

    context.user_data["worksheet"]["FollowOnNotes"] = update.message.text
    await update.message.reply_text("Any engineer notes? Type None if none.")
    return ENGINEER_NOTES


async def worksheet_engineer_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_result = await handle_menu_during_conversation(update, context, ENGINEER_NOTES)
    if menu_result is not None:
        return menu_result

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
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu())
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
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu())
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
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        assigned_ids_latest = get_assigned_engineer_ids(fields)
        if worksheet["engineer_lookup_id"] not in assigned_ids_latest:
            context.user_data.pop("worksheet", None)
            await update.message.reply_text(
                "You are no longer assigned to this job. Worksheet has not been submitted.",
                reply_markup=get_main_menu(),
            )
            return ConversationHandler.END

        outcome = worksheet.get("JobOutcome", "Completed")

        updated_log = append_engineer_log(
            fields,
            worksheet["engineer_name"],
            outcome,
            "Worksheet submitted",
        )

        assigned_ids = get_assigned_engineer_ids(fields)
        is_final_engineer = len(assigned_ids) <= 1

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

        if outcome == "No Access" and worksheet.get("NoAccessReason"):
            fields_to_update["NoAccessReason"] = worksheet.get("NoAccessReason")

        if is_final_engineer:
            fields_to_update["JobOutcome"] = outcome

            if outcome == "Completed":
                fields_to_update["Status"] = COMPLETED_STATUS
                # Keep TelegramNotified as a pure dispatch flag.
                # It should only be set to True after the bot actually sends the job
                # to an engineer in send_new_jobs(). Completed jobs are blocked from
                # being resent by Status/JobOutcome, not by forcing this flag.
                if is_notified(fields):
                    fields_to_update["TelegramNotified"] = True
            else:
                fields_to_update["Status"] = AWAITING_DEPLOYMENT_STATUS
                # Revisit/No Access goes back to the office for re-dispatch.
                # Leave this False so the next assigned engineer can be notified.
                fields_to_update["TelegramNotified"] = False

            fields_to_update.update(clear_engineer_assignment_payload())
        else:
            fields_to_update.update(
                remove_current_engineer_assignment_payload(
                    fields,
                    worksheet["engineer_lookup_id"],
                )
            )

        update_list_item_fields(site_id, jobs_list_id, item_id, fields_to_update)

        update_active_day_live_status(
            site_id,
            str(update.effective_user.id),
            outcome,
            get_job_reference(fields),
        )

        await update.message.reply_text(
            f"Worksheet submitted:\n\n{worksheet['cdr_number']} → {outcome}",
            reply_markup=get_main_menu(),
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
        await query.message.reply_text("Worksheet not found. Tap 📋 My Jobs and try again.", reply_markup=get_main_menu())
        return ConversationHandler.END

    action = query.data.split("|", 1)[1]

    if action == "cancel":
        context.user_data.pop("worksheet", None)
        await query.message.reply_text("Worksheet cancelled. Nothing has been submitted.", reply_markup=get_main_menu())
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
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END

    assigned_ids_latest = get_assigned_engineer_ids(fields)
    if worksheet["engineer_lookup_id"] not in assigned_ids_latest:
        context.user_data.pop("worksheet", None)
        await query.message.reply_text(
            "You are no longer assigned to this job. Worksheet has not been submitted.",
            reply_markup=get_main_menu(),
        )
        return ConversationHandler.END

    outcome = worksheet.get("JobOutcome", "Completed")

    updated_log = append_engineer_log(
        fields,
        worksheet["engineer_name"],
        outcome,
        "Worksheet submitted",
    )

    assigned_ids = get_assigned_engineer_ids(fields)
    is_final_engineer = len(assigned_ids) <= 1

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

    update_list_item_fields(site_id, jobs_list_id, item_id, fields_to_update)

    update_active_day_live_status(
        site_id,
        str(query.from_user.id),
        outcome,
        get_job_reference(fields),
    )

    await query.message.reply_text(
        f"Worksheet submitted:\n\n{worksheet['cdr_number']} → {outcome}",
        reply_markup=get_main_menu(),
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
                        reply_markup=get_main_menu(),
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

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(startday_handler)
telegram_app.add_handler(endday_handler)
telegram_app.add_handler(worksheet_handler)
telegram_app.add_handler(bugidea_handler)
telegram_app.add_handler(MessageHandler(filters.Regex(f"^({MENU_MY_JOBS}|{MENU_BUG_IDEA})$"), menu_button))
telegram_app.add_handler(CallbackQueryHandler(status_button))

if __name__ == "__main__":
    threading.Thread(target=run_signature_web_server, daemon=True).start()
    print(f"Signature web server running on port {PORT}")
    print(f"Bot running... PID={os.getpid()}")

    telegram_app.run_polling(
        drop_pending_updates=True,
        close_loop=False,
        allowed_updates=Update.ALL_TYPES,
    )
