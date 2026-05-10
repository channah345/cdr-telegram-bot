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

LIST_NAME = "Engineer Jobs"

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

app_msal = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=authority,
    client_credential=CLIENT_SECRET,
)

token_result = app_msal.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)

access_token = token_result["access_token"]

headers = {
    "Authorization": f"Bearer {access_token}"
}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("CDR Engineer Bot is online.")

async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Your Telegram ID is: {update.effective_user.id}"
    )

async def jobs(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user_id = str(update.effective_user.id)

    site_hostname = SHAREPOINT_SITE.split("/")[2]
    site_path = "/" + "/".join(SHAREPOINT_SITE.split("/")[3:])

    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"

    site_response = requests.get(site_url, headers=headers)
    site_id = site_response.json()["id"]

    lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"

    lists_response = requests.get(lists_url, headers=headers)

    lists = lists_response.json()["value"]

    list_id = None

    for lst in lists:
        if lst["name"] == LIST_NAME:
            list_id = lst["id"]

    items_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields"

    items_response = requests.get(items_url, headers=headers)

    items = items_response.json()["value"]

    found_jobs = []

    for item in items:

        fields = item["fields"]

        if str(fields.get("TelegramID", "")) == user_id:

            found_jobs.append(
    f"CDR Number: {fields.get('CDRNumber', '')}\n"
    f"Date: {fields.get('Date', '')}\n"
    f"Time: {fields.get('StartTime', '')}\n"
    f"Engineer: {fields.get('EngineerName', '')}\n"
    f"Site: {fields.get('SiteName', '')}\n"
    f"Address: {fields.get('Address', '')}\n"
    f"Task: {fields.get('Task', '')}\n"
    f"Notes: {fields.get('Notes', '')}\n"
    f"Contact: {fields.get('ContactName', '')}\n"
    f"Phone: {fields.get('ContactNumber', '')}"
	)

    if found_jobs:

        await update.message.reply_text(
            "Today's jobs:\n\n" + "\n\n".join(found_jobs)
        )

    else:

        await update.message.reply_text(
            "No jobs assigned today."
        )

telegram_app = ApplicationBuilder().token(BOT_TOKEN).build()

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(CommandHandler("id", id))
telegram_app.add_handler(CommandHandler("jobs", jobs))

print("Bot running...")

telegram_app.run_polling()