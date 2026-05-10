from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

BOT_TOKEN = "8712221208:AAHkc129LjmkZP2N74KkyKOkzuf1iYhX99E"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "CDR Engineer Bot is working."
    )

async def jobs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    jobs_data = {
        123456789: [
            "08:30 - Stockton Fire Station",
            "13:00 - Boiler Fault Darlington"
        ]
    }

    if user_id in jobs_data:
        jobs = "\n".join(jobs_data[user_id])

        await update.message.reply_text(
            f"Today's jobs:\n\n{jobs}"
        )

    else:
        await update.message.reply_text(
            "No jobs assigned today."
        )

async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    await update.message.reply_text(
        f"Your Telegram ID is: {user_id}"
    )

app = ApplicationBuilder().token(BOT_TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("jobs", jobs))
app.add_handler(CommandHandler("id", id))

print("Bot running...")

app.run_polling()