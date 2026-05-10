import os
import requests
import msal

from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

BOT_TOKEN = os.getenv("BOT_TOKEN")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")

JOBS_LIST = "Engineer Jobs"
ENGINEERS_LIST = "Engineers"

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=authority,
    client_credential=CLIENT_SECRET,
)

token_result = msal_app.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)

access_token = token_result["access_token"]

headers = {
    "Authorization": f"Bearer {access_token}"
}


def get_site_id():
    site_hostname = SHAREPOINT_SITE.split("/")[2]
    site_path = "/" + "/".join(SHAREPOINT_SITE.split("/")[3:])
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"

    response = requests.get(site_url, headers=headers)
    response.raise_for_status()

    return response.json()["id"]


def get_list_id(site_id, list_name):
    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    response = requests.get(lists_url, headers=headers)
    response.raise_for_status()

    for lst in response.json()["value"]:
        if lst["name"] == list_name:
            return lst["id"]

    raise Exception(f"List not found: {list_name}")


def get_list_items(site_id, list_id):
    items_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
        f"?expand=fields"
    )

    response = requests.get(items_url, headers=headers)
    response.raise_for_status()

    return response.json()["value"]


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("CDR Engineer Bot is online.")


async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Your Telegram ID is: {update.effective_user.id}"
    )


async def jobs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = str(update.effective_user.id)

    site_id = get_site_id()

    engineers_list_id = get_list_id(site_id, ENGINEERS_LIST)
    jobs_list_id = get_list_id(site_id, JOBS_LIST)

    engineers = get_list_items(site_id, engineers_list_id)
    jobs = get_list_items(site_id, jobs_list_id)

    # Build engineer name → Telegram ID map
    engineer_telegram_ids = {}

    for engineer in engineers:
        fields = engineer["fields"]

        name = fields.get("EngineerName", "")
        telegram_id = str(fields.get("TelegramID", ""))

        if name and telegram_id:
            engineer_telegram_ids[name] = telegram_id

    found_jobs = []

    for job in jobs:
        fields = job["fields"]
	print(fields)

        # Lookup columns often appear as the selected name in Graph
        selected_engineer = fields.get("EngineerLookupValue") or fields.get("Engineer")

        if not selected_engineer:
            continue

        assigned_telegram_id = engineer_telegram_ids.get(selected_engineer)

        if assigned_telegram_id == user_id:
            found_jobs.append(
                f"CDR Number: {fields.get('CDRNumber', '')}\n"
                f"Date: {fields.get('Date', '')}\n"
                f"Time: {fields.get('StartTime', '')}\n"
                f"Engineer: {selected_engineer}\n"
                f"Site: {fields.get('SiteName', '')}\n"
                f"Address: {fields.get('Address', '')}\n"
                f"Task: {fields.get('Task', '')}\n"
                f"Notes: {fields.get('Notes', '')}\n"
                f"Contact: {fields.get('ContactName', '')}\n"
                f"Phone: {fields.get('ContactNumber', '')}"
            )

    if found_jobs:
        await update.message.reply_text(
            "Today's jobs:\n\n" + "\n\n--------------------\n\n".join(found_jobs)
        )
    else:
        await update.message.reply_text("No jobs assigned today.")


telegram_app = ApplicationBuilder().token(BOT_TOKEN).build()

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))

print("Bot running...")

telegram_app.run_polling()