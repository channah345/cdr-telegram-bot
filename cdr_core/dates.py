from datetime import datetime
from zoneinfo import ZoneInfo

UK_TZ = ZoneInfo("Europe/London")


def now_log_time():
    return datetime.now(UK_TZ).strftime("%d/%m/%Y %H:%M")


def graph_datetime_now():
    return datetime.now(UK_TZ).isoformat()


def get_today_iso():
    return datetime.now(UK_TZ).date().isoformat()
