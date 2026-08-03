"""Independent guarded scheduler process for reminders and notifications."""

import asyncio
import os
import signal

os.environ["BOT_RUN_SIGNATURE_SERVER"] = "false"
os.environ["BOT_RUN_SCHEDULER"] = "false"

from bot import start_scheduler, telegram_app


async def run_scheduler():
    await telegram_app.initialize()
    await telegram_app.start()
    scheduler = start_scheduler(telegram_app)
    stopped = asyncio.Event()
    loop = asyncio.get_running_loop()
    for signum in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(signum, stopped.set)
        except NotImplementedError:
            pass
    try:
        await stopped.wait()
    finally:
        if scheduler and scheduler.running:
            scheduler.shutdown(wait=False)
        await telegram_app.stop()
        await telegram_app.shutdown()


if __name__ == "__main__":
    asyncio.run(run_scheduler())
