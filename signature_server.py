"""Independent FastAPI process for client signatures and health checks."""

import os

os.environ["BOT_RUN_SIGNATURE_SERVER"] = "false"
os.environ["BOT_RUN_SCHEDULER"] = "false"

import uvicorn

from bot import web_app


if __name__ == "__main__":
    uvicorn.run(
        web_app,
        host="0.0.0.0",
        port=int(os.getenv("PORT", "8000")),
        proxy_headers=True,
    )
