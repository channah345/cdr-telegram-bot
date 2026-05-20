from datetime import datetime
from zoneinfo import ZoneInfo

UK_TZ = ZoneInfo("Europe/London")

def now_log_time() -> str:
    return datetime.now(UK_TZ).strftime("%d/%m/%Y %H:%M")

def graph_datetime_now() -> str:
    return datetime.now(UK_TZ).isoformat()

def format_sharepoint_date(value) -> str:
    if not value:
        return ""
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return dt.astimezone(UK_TZ).strftime("%d/%m/%Y")
    except Exception:
        return str(value)

def sharepoint_date_to_uk_date(value):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return dt.astimezone(UK_TZ).date()
    except Exception:
        return None

def parse_sharepoint_date_to_date(value):
    return sharepoint_date_to_uk_date(value)
