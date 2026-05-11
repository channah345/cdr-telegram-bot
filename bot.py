import os
import requests
import msal
from datetime import datetime
from zoneinfo import ZoneInfo

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, CallbackQueryHandler,
    ContextTypes, ConversationHandler, MessageHandler, filters
)

BOT_TOKEN = os.getenv("BOT_TOKEN")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
HELPDESK_CHAT_ID = os.getenv("HELPDESK_CHAT_ID")

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"
PHOTO_LIBRARY = "15 - ENGINEER JOB PHOTOS"

UK_TZ = ZoneInfo("Europe/London")

WORK_COMPLETED, MATERIALS_USED, FOLLOW_ON_REQUIRED, FOLLOW_ON_NOTES, ENGINEER_NOTES, PHOTOS = range(6)

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
            InlineKeyboardButton("Need Parts", callback_data=f"status|{item_id}|Need Parts"),
            InlineKeyboardButton("Revisit", callback_data=f"status|{item_id}|Revisit Required"),
        ],
        [
            InlineKeyboardButton("No Access", callback_data=f"status|{item_id}|No Access"),
        ],
    ])


def format_job(fields, engineer_name=None):
    return (
        f"CDR Number: {fields.get('CDRNumber', '')}\n"
        f"Date: {format_sharepoint_date(fields.get('Date', ''))}\n"
        f"Time: {fields.get('StartTime', '')}\n"
        f"Engineer: {engineer_name or ''}\n"
        f"Status: {fields.get('Status', 'Assigned')}\n"
        f"Site: {fields.get('SiteName', '')}\n"
        f"Address: {fields.get('Address', '')}\n"
        f"Task: {fields.get('Task', '')}\n"
        f"Notes: {fields.get('Notes', '')}\n"
        f"Contact: {fields.get('ContactName', '')}\n"
        f"Phone: {fields.get('ContactNumber', '')}"
    )


def find_job_by_cdr(jobs_data, cdr_number):
    for job in jobs_data:
        fields = job["fields"]
        if str(fields.get("CDRNumber", "")).lower() == cdr_number.lower():
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


def ensure_photo_folder(drive_id, folder_name):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"

    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace",
    }

    response = requests.post(url, headers=get_headers(), json=body)

    if response.status_code not in [200, 201]:
        raise Exception(f"Could not create photo folder: {response.text}")


def upload_photo_to_sharepoint(drive_id, folder_name, file_name, file_bytes):
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:"
        f"/{folder_name}/{file_name}:/content"
    )

    response = requests.put(
        url,
        headers=get_headers(content_type=False),
        data=file_bytes,
    )

    if response.status_code not in [200, 201]:
        raise Exception(f"Could not upload photo: {response.text}")

    return response.json().get("webUrl", "")


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
            status = fields.get("Status", "Assigned")

            if status == "Completed":
                continue

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
        item_id = data[1]
        new_status = data[2]

        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)

        user_id = str(query.from_user.id)
        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await query.message.reply_text("You are not set up as an engineer.")
            return

        job = None
        for item in jobs_data:
            if str(item.get("id")) == str(item_id):
                job = item
                break

        if not job:
            await query.message.reply_text("Could not find this job.")
            return

        fields = job["fields"]
        assigned_ids = get_assigned_engineer_ids(fields)

        if current_engineer["lookup_id"] not in assigned_ids:
            await query.message.reply_text("You are not assigned to this job.")
            return

        update_list_item_fields(site_id, jobs_list_id, item_id, {"Status": new_status})

        await query.message.reply_text(
            f"Status updated:\n\n{fields.get('CDRNumber', '')} → {new_status}"
        )

        await notify_helpdesk(
            context,
            (
                f"Job status updated\n\n"
                f"CDR Number: {fields.get('CDRNumber', '')}\n"
                f"Engineer: {current_engineer['name']}\n"
                f"Status: {new_status}\n"
                f"Site: {fields.get('SiteName', '')}"
            ),
        )

    except Exception as e:
        print(f"ERROR updating status: {e}")
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

        context.user_data["worksheet"] = {
            "cdr_number": cdr_number,
            "site_id": site_id,
            "jobs_list_id": jobs_list_id,
            "item_id": job["id"],
            "engineer_name": current_engineer["name"],
            "fields": fields,
            "photo_links": [],
        }

        await update.message.reply_text(
            f"Starting worksheet for {cdr_number}.\n\nWhat work was completed?"
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
        "Upload job photos now.\n\nSend one or more photos, then type DONE when finished."
    )

    return PHOTOS


async def worksheet_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    worksheet = context.user_data["worksheet"]

    if update.message.text and update.message.text.strip().upper() == "DONE":
        fields_to_update = {
            "WorkCompleted": worksheet.get("WorkCompleted", ""),
            "MaterialsUsed": worksheet.get("MaterialsUsed", ""),
            "FollowOnRequired": worksheet.get("FollowOnRequired", False),
            "FollowOnNotes": worksheet.get("FollowOnNotes", ""),
            "EngineerCompletionNotes": worksheet.get("EngineerCompletionNotes", ""),
            "WorksheetSubmitted": True,
            "Status": "Completed",
        }

        update_list_item_fields(
            worksheet["site_id"],
            worksheet["jobs_list_id"],
            worksheet["item_id"],
            fields_to_update,
        )

        await update.message.reply_text(
            f"Worksheet submitted and job completed:\n\n{worksheet['cdr_number']}"
        )

        await notify_helpdesk(
            context,
            (
                f"Worksheet submitted\n\n"
                f"CDR Number: {worksheet['cdr_number']}\n"
                f"Engineer: {worksheet['engineer_name']}\n"
                f"Photos uploaded: {len(worksheet.get('photo_links', []))}"
            ),
        )

        context.user_data.pop("worksheet", None)
        return ConversationHandler.END

    if update.message.photo:
        site_id = worksheet["site_id"]
        cdr_number = worksheet["cdr_number"]

        drive_id = get_drive_id(site_id, PHOTO_LIBRARY)
        ensure_photo_folder(drive_id, cdr_number)

        photo = update.message.photo[-1]
        telegram_file = await context.bot.get_file(photo.file_id)
        file_bytes = await telegram_file.download_as_bytearray()

        timestamp = datetime.now(UK_TZ).strftime("%Y%m%d_%H%M%S")
        file_name = f"{cdr_number}_{timestamp}_{photo.file_unique_id}.jpg"

        photo_link = upload_photo_to_sharepoint(
            drive_id,
            cdr_number,
            file_name,
            bytes(file_bytes),
        )

        worksheet["photo_links"].append(photo_link)

        await update.message.reply_text(
            f"Photo uploaded. Total photos: {len(worksheet['photo_links'])}\n\nSend another photo or type DONE."
        )

        return PHOTOS

    await update.message.reply_text("Please send a photo or type DONE.")
    return PHOTOS


async def worksheet_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("worksheet", None)
    await update.message.reply_text("Worksheet cancelled.")
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
                {"TelegramNotified": True},
            )

    except Exception as e:
        print(f"ERROR sending new jobs: {e}")


async def post_init(app):
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
    },
    fallbacks=[CommandHandler("cancel", worksheet_cancel)],
)

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(worksheet_handler)
telegram_app.add_handler(CallbackQueryHandler(status_button))

print("Bot running...")

telegram_app.run_polling()