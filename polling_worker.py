"""Telegram polling process. Run exactly one replica of this service."""

import os

os.environ["BOT_RUN_SIGNATURE_SERVER"] = "false"
os.environ["BOT_RUN_SCHEDULER"] = "false"

from telegram import Update

from bot import telegram_app


if __name__ == "__main__":
    telegram_app.run_polling(
        drop_pending_updates=False,
        close_loop=False,
        allowed_updates=Update.ALL_TYPES,
    )
