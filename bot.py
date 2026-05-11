import os
import base64
import secrets
import threading
import requests
import msal
from datetime import datetime
from zoneinfo import ZoneInfo

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
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

PHOTO_LIBRARY = "Documents"
PHOTO_BASE_FOLDER = "15 - ENGINEER JOB PHOTOS"
SIGNATURE_BASE_FOLDER = "16 - CLIENT SIGNATURES"

UK_TZ = ZoneInfo("Europe/London")

AWAITING_DEPLOYMENT_STATUS = "Awaiting Engineer Deployment"
ASSIGNED_STATUS = "Assigned"
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


def update_list_item_fields(site_id, list_id, item_id, fields_to_update):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"

    response = requests.patch(
        url,
        headers=get_headers(),
        json=fields_to_update,
    )

    if response.status_code not in [200, 204]:
        raise Exception(f"Could not update item {item_id}: {response.text}")


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


def engineer_has_logged(fields, engineer_name, action):
    log = fields.get("EngineerVisitLog", "") or ""
    search_text = f" - {engineer_name} - {action}"
    return search_text in log


def can_click_action(fields, engineer_name, action):
    has_travelled = engineer_has_logged(fields, engineer_name, "Travelling")
    has_on_site = engineer_has_logged(fields, engineer_name, "On Site")

    if action == "Travelling":
        return True, ""

    if action == "On Site" and not has_travelled:
        return False, "You need to click Travelling before clicking On Site."

    if action in ["No Access", "Revisit Required", "Completed"] and not has_on_site:
        return False, "You need to click On Site before selecting this option."

    return True, ""


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


def get_job_buttons(item_id):
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("Travelling", callback_data=f"status|{item_id}|Travelling"),
            InlineKeyboardButton("On Site", callback_data=f"status|{item_id}|On Site"),
        ],
        [
            InlineKeyboardButton("Revisit", callback_data=f"confirm_outcome|{item_id}|Revisit Required"),
            InlineKeyboardButton("No Access", callback_data=f"confirm_outcome|{item_id}|No Access"),
        ],
        [
            InlineKeyboardButton("Complete", callback_data=f"complete_help|{item_id}"),
        ],
    ])


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


def find_job_by_cdr(jobs_data, cdr_number):
    for job in jobs_data:
        fields = job["fields"]
        if str(fields.get("CDRNumber", "")).lower() == cdr_number.lower():
            return job
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
            str(fields.get("CDRNumber", "")).lower() == cdr_number.lower()
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
    signature_received = "Yes" if worksheet.get("ClientSignatureReceived") else "No"

    return (
        f"Please review worksheet for {worksheet['cdr_number']}:\n\n"
        f"Work completed:\n{worksheet.get('WorkCompleted', '')}\n\n"
        f"Materials used:\n{worksheet.get('MaterialsUsed', '')}\n\n"
        f"Follow-on required:\n{'Yes' if worksheet.get('FollowOnRequired') else 'No'}\n\n"
        f"Follow-on notes:\n{worksheet.get('FollowOnNotes', '') or 'None'}\n\n"
        f"Engineer notes:\n{worksheet.get('EngineerCompletionNotes', '')}\n\n"
        f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n"
        f"Client signature required: {signature_required}\n"
        f"Client signature received: {signature_received}\n\n"
        f"Type SUBMIT to complete the job.\n"
        f"Type RESTART to redo the worksheet.\n"
        f"Type CANCEL to abandon it."
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
            body {{ font-family: Arial, sans-serif; background: #f4f4f4; padding: 20px; margin: 0; }}
            .container {{ max-width: 650px; margin: auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.12); }}
            h1 {{ color: #f58220; margin-bottom: 5px; }}
            .job-box {{ background: #f7f7f7; padding: 15px; border-radius: 8px; margin: 20px 0; }}
            label {{ font-weight: bold; display: block; margin-top: 15px; }}
            input[type="text"] {{ width: 100%; padding: 12px; font-size: 16px; box-sizing: border-box; }}
            canvas {{ width: 100%; height: 230px; border: 2px solid #333; border-radius: 8px; background: white; margin-top: 10px; }}
            button {{ width: 100%; padding: 14px; margin-top: 15px; font-size: 16px; border: none; border-radius: 8px; cursor: pointer; }}
            .submit {{ background: #f58220; color: white; font-weight: bold; }}
            .clear {{ background: #555; color: white; }}
            .small {{ font-size: 13px; color: #555; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>CDR M&E Services Ltd</h1>
            <h2>Client Signature</h2>
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
            const signaturePad = new SignaturePad(canvas);

            function resizeCanvas() {{
                const ratio = Math.max(window.devicePixelRatio || 1, 1);
                const rect = canvas.getBoundingClientRect();
                canvas.width = rect.width * ratio;
                canvas.height = rect.height * ratio;
                canvas.getContext("2d").scale(ratio, ratio);
                signaturePad.clear();
            }}

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
    await update.message.reply_text("CDR Engineer Bot is online.")


async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Your Telegram ID is: {update.effective_user.id}")


async def jobs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = str(update.effective_user.id)
        today = datetime.now(UK_TZ).date()

        _, _, _, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)

        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID."
            )
            return

        found_any = False

        for job in jobs_data:
            fields = job["fields"]
            item_id = job["id"]

            job_date = sharepoint_date_to_uk_date(fields.get("Date", ""))
            assigned_ids = get_assigned_engineer_ids(fields)

            if current_engineer["lookup_id"] in assigned_ids and job_date == today:
                found_any = True
                await update.message.reply_text(
                    "Today's job:\n\n" + format_job(fields, current_engineer["name"]),
                    reply_markup=get_job_buttons(item_id),
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

        job = find_job_by_item_id(jobs_data, item_id)

        if not job:
            await query.message.reply_text("Could not find this job.")
            return

        fields = job["fields"]
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
                        callback_data=f"outcome|{item_id}|{selected_outcome}",
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
                },
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
                update_fields.update(clear_engineer_assignment_payload())
            else:
                update_fields.update(
                    remove_current_engineer_assignment_payload(
                        fields,
                        current_engineer["lookup_id"],
                    )
                )

            update_list_item_fields(site_id, jobs_list_id, item_id, update_fields)

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


async def complete_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if not context.args:
            await update.message.reply_text("To complete a job, type:\n\n/complete CDR00001")
            return ConversationHandler.END

        cdr_number = context.args[0].strip()
        user_id = str(update.effective_user.id)

        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)

        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID."
            )
            return ConversationHandler.END

        job = find_job_by_cdr(jobs_data, cdr_number)

        if not job:
            await update.message.reply_text(f"No job found with CDR number: {cdr_number}")
            return ConversationHandler.END

        fields = job["fields"]
        assigned_ids = get_assigned_engineer_ids(fields)

        if current_engineer["lookup_id"] not in assigned_ids:
            await update.message.reply_text("You are not assigned to this job.")
            return ConversationHandler.END

        allowed, reason = can_click_action(
            fields,
            current_engineer["name"],
            "Completed",
        )

        if not allowed:
            await update.message.reply_text(reason)
            return ConversationHandler.END

        context.user_data["worksheet"] = {
            "cdr_number": cdr_number,
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "item_id": job["id"],
            "engineer_name": current_engineer["name"],
            "engineer_lookup_id": current_engineer["lookup_id"],
            "fields": fields,
            "photo_links": [],
            "ClientSignatureRequired": False,
            "ClientSignatureReceived": False,
        }

        await update.message.reply_text(
            f"Starting worksheet for {cdr_number}.\n\n"
            f"You can type /cancel at any point before submitting.\n\n"
            f"What work was completed?"
        )

        return WORK_COMPLETED

    except Exception as e:
        print(f"ERROR starting worksheet: {e}")
        await update.message.reply_text("There was an error starting the worksheet.")
        return ConversationHandler.END


async def worksheet_work_completed(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["worksheet"]["WorkCompleted"] = update.message.text
    await update.message.reply_text("What materials were used? Type None if none.")
    return MATERIALS_USED


async def worksheet_materials_used(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["worksheet"]["MaterialsUsed"] = update.message.text
    await update.message.reply_text("Is a follow-on required? Reply Yes or No.")
    return FOLLOW_ON_REQUIRED


async def worksheet_follow_on_required(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = update.message.text.strip().lower()

    if answer not in ["yes", "no", "y", "n"]:
        await update.message.reply_text("Please reply Yes or No.")
        return FOLLOW_ON_REQUIRED

    follow_on_required = answer in ["yes", "y"]
    context.user_data["worksheet"]["FollowOnRequired"] = follow_on_required

    if follow_on_required:
        await update.message.reply_text("What follow-on is required?")
        return FOLLOW_ON_NOTES

    context.user_data["worksheet"]["FollowOnNotes"] = ""
    await update.message.reply_text("Any engineer notes? Type None if none.")
    return ENGINEER_NOTES


async def worksheet_follow_on_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["worksheet"]["FollowOnNotes"] = update.message.text
    await update.message.reply_text("Any engineer notes? Type None if none.")
    return ENGINEER_NOTES


async def worksheet_engineer_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["worksheet"]["EngineerCompletionNotes"] = update.message.text

    await update.message.reply_text(
        "Upload job photos now.\n\n"
        "Send one or more photos, then type DONE when finished.\n"
        "If no photos are needed, type DONE."
    )

    return PHOTOS


async def worksheet_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    worksheet = context.user_data["worksheet"]

    if update.message.text and update.message.text.strip().upper() == "DONE":
        await update.message.reply_text("Is a client signature required? Reply Yes or No.")
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

    await update.message.reply_text("Please send a photo or type DONE.")
    return PHOTOS


async def worksheet_signature_required(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = update.message.text.strip().lower()
    worksheet = context.user_data["worksheet"]

    if answer not in ["yes", "no", "y", "n"]:
        await update.message.reply_text("Please reply Yes or No.")
        return SIGNATURE_REQUIRED

    signature_required = answer in ["yes", "y"]
    worksheet["ClientSignatureRequired"] = signature_required

    if not signature_required:
        await update.message.reply_text(build_review_text(worksheet))
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
            "Once signed, type SIGNED.\n"
            "If no client is available, type SKIP."
        )

        return SIGNATURE_WAITING

    except Exception as e:
        print(f"ERROR creating signature link: {e}")
        await update.message.reply_text(
            "There was an error creating the signature link. Type SKIP to continue without a signature."
        )
        return SIGNATURE_WAITING


async def worksheet_signature_waiting(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = update.message.text.strip().upper()
    worksheet = context.user_data["worksheet"]

    if answer == "SKIP":
        worksheet["ClientSignatureReceived"] = False
        await update.message.reply_text(build_review_text(worksheet))
        return REVIEW

    if answer != "SIGNED":
        await update.message.reply_text("Please type SIGNED once the client has signed, or SKIP to continue without a signature.")
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
        await update.message.reply_text(build_review_text(worksheet))
        return REVIEW

    await update.message.reply_text(
        "I cannot see the signature yet. Make sure the client pressed Submit Signature, then type SIGNED again."
    )
    return SIGNATURE_WAITING


async def worksheet_review(update: Update, context: ContextTypes.DEFAULT_TYPE):
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

        updated_log = append_engineer_log(
            fields,
            worksheet["engineer_name"],
            "Completed",
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
            "JobOutcome": "Completed",
            "EngineerVisitLog": updated_log,
            "ClientSignatureRequired": worksheet.get("ClientSignatureRequired", False),
        }

        if is_final_engineer:
            fields_to_update["Status"] = COMPLETED_STATUS
            fields_to_update.update(clear_engineer_assignment_payload())
        else:
            fields_to_update.update(
                remove_current_engineer_assignment_payload(
                    fields,
                    worksheet["engineer_lookup_id"],
                )
            )

        update_list_item_fields(site_id, jobs_list_id, item_id, fields_to_update)

        await update.message.reply_text(
            f"Worksheet submitted and job completed:\n\n{worksheet['cdr_number']}"
        )

        await notify_helpdesk(
            context,
            (
                f"Worksheet submitted\n\n"
                f"CDR Number: {worksheet['cdr_number']}\n"
                f"Engineer: {worksheet['engineer_name']}\n"
                f"Outcome: Completed\n"
                f"Final engineer: {'Yes' if is_final_engineer else 'No'}\n"
                f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n"
                f"Client signature required: {'Yes' if worksheet.get('ClientSignatureRequired') else 'No'}\n"
                f"Client signature received: {'Yes' if worksheet.get('ClientSignatureReceived') else 'No'}"
            ),
        )

        context.user_data.pop("worksheet", None)
        return ConversationHandler.END

    await update.message.reply_text("Please type SUBMIT, RESTART or CANCEL.")
    return REVIEW


async def worksheet_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("worksheet", None)
    await update.message.reply_text("Worksheet cancelled. Nothing has been submitted.")
    return ConversationHandler.END


async def send_new_jobs(app):
    try:
        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        _, engineers_by_lookup = build_engineer_maps(engineers)

        sent_job_ids = set()

        for job in jobs_data:
            fields = job["fields"]
            item_id = job["id"]

            if is_notified(fields):
                continue

            assigned_ids = get_assigned_engineer_ids(fields)

            if not assigned_ids:
                continue

            for engineer_id in assigned_ids:
                engineer = engineers_by_lookup.get(engineer_id)

                if not engineer:
                    continue

                await app.bot.send_message(
                    chat_id=engineer["telegram_id"],
                    text="New job assigned:\n\n" + format_job(fields, engineer["name"]),
                    reply_markup=get_job_buttons(item_id),
                )

            sent_job_ids.add(item_id)

        for item_id in sent_job_ids:
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
        print(f"ERROR sending new jobs: {e}")


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

    scheduler.start()
    print("Scheduler started.")


telegram_app = (
    ApplicationBuilder()
    .token(BOT_TOKEN)
    .post_init(post_init)
    .build()
)

worksheet_handler = ConversationHandler(
    entry_points=[CommandHandler("complete", complete_start)],
    states={
        WORK_COMPLETED: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_work_completed)],
        MATERIALS_USED: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_materials_used)],
        FOLLOW_ON_REQUIRED: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_follow_on_required)],
        FOLLOW_ON_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_follow_on_notes)],
        ENGINEER_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_engineer_notes)],
        PHOTOS: [
            MessageHandler(filters.PHOTO, worksheet_photos),
            MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_photos),
        ],
        SIGNATURE_REQUIRED: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_signature_required)],
        SIGNATURE_WAITING: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_signature_waiting)],
        REVIEW: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_review)],
    },
    fallbacks=[CommandHandler("cancel", worksheet_cancel)],
)

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(worksheet_handler)
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
