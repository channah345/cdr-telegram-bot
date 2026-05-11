import os
import requests
import msal
from datetime import datetime
from zoneinfo import ZoneInfo

from apscheduler.schedulers.asyncio import AsyncIOScheduler
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

BOT_TOKEN = os.getenv("BOT_TOKEN")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
HELPDESK_CHAT_ID = os.getenv("HELPDESK_CHAT_ID")

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"

PHOTO_LIBRARY = "Documents"
PHOTO_BASE_FOLDER = "15 - ENGINEER JOB PHOTOS"

UK_TZ = ZoneInfo("Europe/London")

AWAITING_DEPLOYMENT_STATUS = "Awaiting Engineer Deployment"
ASSIGNED_STATUS = "Assigned"
COMPLETED_STATUS = "Completed"

WORK_COMPLETED, MATERIALS_USED, FOLLOW_ON_REQUIRED, FOLLOW_ON_NOTES, ENGINEER_NOTES, PHOTOS, REVIEW = range(7)

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=authority,
    client_credential=CLIENT_SECRET,
)


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

            if lookup_id != str(current_lookup_id):
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


def ensure_photo_folder(drive_id, folder_path):
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


def upload_photo_to_sharepoint(drive_id, folder_path, file_name, file_bytes):
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
        raise Exception(f"Could not upload photo: {response.text}")

    return response.json().get("webUrl", "")


def build_review_text(worksheet):
    return (
        f"Please review worksheet for {worksheet['cdr_number']}:\n\n"
        f"Work completed:\n{worksheet.get('WorkCompleted', '')}\n\n"
        f"Materials used:\n{worksheet.get('MaterialsUsed', '')}\n\n"
        f"Follow-on required:\n{'Yes' if worksheet.get('FollowOnRequired') else 'No'}\n\n"
        f"Follow-on notes:\n{worksheet.get('FollowOnNotes', '') or 'None'}\n\n"
        f"Engineer notes:\n{worksheet.get('EngineerCompletionNotes', '')}\n\n"
        f"Photos uploaded: {len(worksheet.get('photo_links', []))}\n\n"
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

            update_fields = {
                "Status": selected_status,
                "EngineerVisitLog": updated_log,
            }

            update_list_item_fields(site_id, jobs_list_id, item_id, update_fields)

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

            if is_final_engineer:
                await query.message.reply_text(
                    f"Updated:\n\n"
                    f"{fields.get('CDRNumber', '')} → {selected_outcome}\n"
                    f"Final engineer removed. Job returned to Awaiting Engineer Deployment."
                )
            else:
                await query.message.reply_text(
                    f"Updated:\n\n"
                    f"{fields.get('CDRNumber', '')} → {selected_outcome}\n"
                    f"You have been removed from this job. Other assigned engineer(s) remain."
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
        await update.message.reply_text(build_review_text(worksheet))
        return REVIEW

    if update.message.photo:
        site_id = worksheet["site_id"]
        cdr_number = worksheet["cdr_number"]

        drive_id = get_drive_id(site_id, PHOTO_LIBRARY)

        folder_path = f"{PHOTO_BASE_FOLDER}/{cdr_number}"
        ensure_photo_folder(drive_id, folder_path)

        photo = update.message.photo[-1]
        telegram_file = await context.bot.get_file(photo.file_id)
        file_bytes = await telegram_file.download_as_bytearray()

        timestamp = datetime.now(UK_TZ).strftime("%Y%m%d_%H%M%S")
        file_name = f"{cdr_number}_{timestamp}_{photo.file_unique_id}.jpg"

        photo_link = upload_photo_to_sharepoint(
            drive_id,
            folder_path,
            file_name,
            bytes(file_bytes),
        )

        worksheet["photo_links"].append(photo_link)

        return PHOTOS

    await update.message.reply_text("Please send a photo or type DONE.")
    return PHOTOS


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

        if is_final_engineer:
            await update.message.reply_text(
                f"Worksheet submitted and job completed:\n\n{worksheet['cdr_number']}"
            )
        else:
            await update.message.reply_text(
                f"Worksheet submitted:\n\n"
                f"{worksheet['cdr_number']}\n\n"
                f"You have been removed from this job. Other assigned engineer(s) remain."
            )

        await notify_helpdesk(
            context,
            (
                f"Worksheet submitted\n\n"
                f"CDR Number: {worksheet['cdr_number']}\n"
                f"Engineer: {worksheet['engineer_name']}\n"
                f"Outcome: Completed\n"
                f"Final engineer: {'Yes' if is_final_engineer else 'No'}\n"
                f"Photos uploaded: {len(worksheet.get('photo_links', []))}"
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
        minutes=2,
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
        REVIEW: [MessageHandler(filters.TEXT & ~filters.COMMAND, worksheet_review)],
    },
    fallbacks=[CommandHandler("cancel", worksheet_cancel)],
)

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(worksheet_handler)
telegram_app.add_handler(CallbackQueryHandler(status_button))

print("Bot running...")

telegram_app.run_polling(
    drop_pending_updates=True,
    close_loop=False,
)