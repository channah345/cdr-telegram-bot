import os
from dataclasses import dataclass


@dataclass
class CDRConfig:
    bot_token: str = os.getenv("BOT_TOKEN", "")
    tenant_id: str = os.getenv("TENANT_ID", "")
    client_id: str = os.getenv("CLIENT_ID", "")
    client_secret: str = os.getenv("CLIENT_SECRET", "")
    sharepoint_site: str = os.getenv("SHAREPOINT_SITE", "")
    port: int = int(os.getenv("PORT", "8000"))
