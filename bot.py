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
)

BOT_TOKEN = os.getenv("BOT_TOKEN")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
HELPDESK_CHAT_ID = os.getenv("HELPDESK_CHAT_ID")

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"

UK_TZ = ZoneInfo("Europe/London")

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=authority,
    client_credential=CLIENT_SECRET,
)


def get_headers():
    token_result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in token_result:
        raise Exception(f"Could not get Microsoft token: {token_result}")

    return {
        "Authorization": f"Bearer {token_result['access_token']}",
        "Content-Type": "application/json",
    }


def format_sharepoint_date(value):
    if not value:
        return ""

    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        uk_time = dt.astimezone(UK_TZ)
        return uk_time.strftime("%d/%m/%Y")
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
        raise Exception(
            f"Could not get SharePoint site. "
            f"Status: {response.status_code}. Response: {response.text}"
        )

    return response.json()["id"]


def get_list_id(site_id, list_name):
    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    response = requests.get(lists_url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(
            f"Could not get SharePoint lists. "
            f"Status: {response.status_code}. Response: {response.text}"
        )

    for lst in response.json()["value"]:
        if lst["name"] == list_name:
            return lst["id"]

    raise Exception(f"List not found: {list_name}")


def get_list_items(site_id, list_id):
    items_url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{site_id}/lists/{list_id}/items?expand=fields"
    )

    response = requests.get(items_url, headers=get_headers())

    if response.status_code != 200:
        raise Exception(
            f"Could not get list items. "
            f"Status: {response.status_code}. Response: {response.text}"
        )

    return response.json()["value"]


def update_list_item_fields(site_id, list_id, item_id, fields_to_update):
    update_url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{site_id}/lists/{list_id}/items/{item_id}/fields"
    )

    response = requests.patch(
        update_url,
        headers=get_headers(),
        json=fields_to_update,
    )

    if response.status_code not in [200, 204]:
        raise Exception(
            f"Could not update item {item_id}. "
            f"Status: {response.status_code}. Response: {response.text}"
        )


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
    engineer_lookup_values = fields.get("Engineer", [])
    assigned_engineer_ids = []

    if isinstance(engineer_lookup_values, list):
        for engineer in engineer_lookup_values:
            assigned_engineer_ids.append(str(engineer.get("LookupId")))

    return assigned_engineer_ids


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
        f"Status: {fields.get('Status', 'Assigned')}\n"
        f"Site: {fields.get('SiteName', '')}\n"
        f"Address: {fields.get('Address', '')}\n"
        f"Task: {fields.get('Task', '')}\n"
        f"Notes: {fields.get('Notes', '')}\n"
        f"Contact: {fields.get('ContactName', '')}\n"
        f"Phone: {fields.get('ContactNumber', '')}"
    )


def find_job_by_item_id(jobs_data, item_id):
    for job in jobs_data:
        if str(job.get("id")) == str(item_id):
            return job
    return None


async def notify_helpdesk(context, text):
    if HELPDESK_CHAT_ID:
        try:
            await context.bot.send_message(
                chat_id=HELPDESK_CHAT_ID,
                text=text,
            )
        except Exception as e:
            print(f"Could not notify helpdesk: {e}")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("CDR Engineer Bot is online.")


async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Your Telegram ID is: {update.effective_user.id}"
    )


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

            if (
                current_engineer["lookup_id"] in assigned_ids
                and job_date == today
            ):
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
            await query.edit_message_text(
                "You are not set up as an engineer. Please ask the office to add your Telegram ID."
            )
            return

        job = find_job_by_item_id(jobs_data, item_id)

        if not job:
            await query.edit_message_text("Could not find this job in SharePoint.")
            return

        fields = job["fields"]
        assigned_ids = get_assigned_engineer_ids(fields)

        if current_engineer["lookup_id"] not in assigned_ids:
            await query.edit_message_text("You are not assigned to this job.")
            return

        if action == "complete_help":
            cdr_number = fields.get("CDRNumber", "")
            await query.message.reply_text(
                f"To complete this job, type:\n\n/complete {cdr_number}"
            )
            return

        if action == "status":
            new_status = data[2]

            update_list_item_fields(
                site_id,
                jobs_list_id,
                item_id,
                {"Status": new_status},
            )

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
        await query.message.reply_text(
            "There was an error updating the job status. Please ask the office to check Railway logs."
        )


async def complete(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if not context.args:
            await update.message.reply_text(
                "To complete a job, type:\n\n/complete CDR00001"
            )
            return

        cdr_number_requested = context.args[0].strip()

        user_id = str(update.effective_user.id)

        site_id, _, jobs_list_id, engineers, jobs_data = get_sharepoint_data()
        engineers_by_telegram, _ = build_engineer_maps(engineers)

        current_engineer = engineers_by_telegram.get(user_id)

        if not current_engineer:
            await update.message.reply_text(
                "You are not set up as an engineer yet. Please ask the office to add your Telegram ID."
            )
            return

        for job in jobs_data:
            fields = job["fields"]
            item_id = job["id"]

            if str(fields.get("CDRNumber", "")).lower() != cdr_number_requested.lower():
                continue

            assigned_ids = get_assigned_engineer_ids(fields)

            if current_engineer["lookup_id"] not in assigned_ids:
                await update.message.reply_text("You are not assigned to this job.")
                return

            update_list_item_fields(
                site_id,
                jobs_list_id,
                item_id,
                {"Status": "Completed"},
            )

            await update.message.reply_text(
                f"Job completed:\n\n{fields.get('CDRNumber', '')}"
            )

            await notify_helpdesk(
                context,
                (
                    f"Job completed\n\n"
                    f"CDR Number: {fields.get('CDRNumber', '')}\n"
                    f"Engineer: {current_engineer['name']}\n"
                    f"Site: {fields.get('SiteName', '')}"
                ),
            )

            return

        await update.message.reply_text(
            f"No job found with CDR number: {cdr_number_requested}"
        )

    except Exception as e:
        print(f"ERROR completing job: {e}")
        await update.message.reply_text(
            "There was an error completing the job. Please ask the office to check Railway logs."
        )


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

                telegram_id = engineer["telegram_id"]
                engineer_name = engineer["name"]

                await app.bot.send_message(
                    chat_id=telegram_id,
                    text="New job assigned:\n\n" + format_job(fields, engineer_name),
                    reply_markup=get_job_buttons(item_id),
                )

                print(
                    f"Sent job {fields.get('CDRNumber', '')} "
                    f"to {engineer_name} ({telegram_id})"
                )

            sent_job_ids.add(item_id)

        for item_id in sent_job_ids:
            update_list_item_fields(
                site_id,
                jobs_list_id,
                item_id,
                {"TelegramNotified": True},
            )

        if sent_job_ids:
            print(f"Marked {len(sent_job_ids)} job(s) as TelegramNotified.")

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

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))
telegram_app.add_handler(CommandHandler("complete", complete))
telegram_app.add_handler(CallbackQueryHandler(status_button))

print("Bot running...")

telegram_app.run_polling()