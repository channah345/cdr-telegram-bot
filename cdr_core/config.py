"""Environment-backed configuration shared by the portal and bot."""

from dataclasses import dataclass
import os


@dataclass(frozen=True)
class CDRConfig:
    bot_token: str = os.getenv("BOT_TOKEN", "")
    tenant_id: str = os.getenv("TENANT_ID", "")
    client_id: str = os.getenv("CLIENT_ID", "")
    client_secret: str = os.getenv("CLIENT_SECRET", "")
    sharepoint_site: str = os.getenv("SHAREPOINT_SITE", "")
    database_url: str = os.getenv("DATABASE_URL", "")
    photo_library: str = os.getenv("PHOTO_LIBRARY", "Documents")
    jobs_list: str = os.getenv("JOBS_LIST", "Engineer Jobs")
    engineers_list: str = os.getenv("ENGINEERS_LIST", "Engineers")
    day_logs_list: str = os.getenv("DAY_LOGS_LIST", "Engineer Day Logs")
    task_activities_list: str = os.getenv("TASK_ACTIVITIES_LIST", "Task Activities")
    port: int = int(os.getenv("PORT", "8000"))

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"
