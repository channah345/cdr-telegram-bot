"""Pure job rules shared by web and Telegram interfaces."""

from .statuses import (
    ASSIGNED_STATUS, AWAITING_DEPLOYMENT_STATUS, COMPLETED_STATUS,
    LEGACY_AWAITING_DEPLOYMENT_STATUS, ON_SITE_STATUS, TRAVELLING_STATUS,
)


def get_assigned_engineer_ids(fields: dict) -> list[str]:
    assigned, seen = [], set()

    def add(value):
        value = str(value or "").strip()
        if value and value.lower() not in {"none", "null"} and value not in seen:
            assigned.append(value)
            seen.add(value)

    engineers = fields.get("Engineer", [])
    for engineer in engineers if isinstance(engineers, list) else []:
        add(engineer.get("LookupId") if isinstance(engineer, dict) else engineer)
    for key in ["EngineerLookupId", "Engineer Lookup Id", "EngineerId", "Engineer ID"]:
        value = fields.get(key)
        if isinstance(value, list):
            for item in value:
                add(item)
        else:
            add(value)
    return assigned


def is_closed_job(fields: dict) -> bool:
    status = str(fields.get("Status", "") or "").strip()
    outcome = str(fields.get("JobOutcome", "") or "").strip()
    open_statuses = {
        "", AWAITING_DEPLOYMENT_STATUS, LEGACY_AWAITING_DEPLOYMENT_STATUS,
        ASSIGNED_STATUS, TRAVELLING_STATUS, ON_SITE_STATUS,
    }
    if status == COMPLETED_STATUS or outcome == "Completed":
        return True
    if status in open_statuses:
        return False
    return status in {"No Access", "Revisit Required"} or outcome in {"No Access", "Revisit Required"}


def engineer_has_logged(fields: dict, engineer_name: str, action: str) -> bool:
    return f" - {engineer_name} - {action}" in (fields.get("EngineerVisitLog", "") or "")


def validate_job_action(fields: dict, engineer_name: str, action: str):
    if is_closed_job(fields):
        return False, "This job has already been closed or returned to the office."
    has_travelled = engineer_has_logged(fields, engineer_name, "Travelling")
    has_on_site = engineer_has_logged(fields, engineer_name, "On Site")
    if action == "Travelling" and has_travelled:
        return False, "Travelling has already been logged for this job."
    if action == "On Site":
        if not has_travelled:
            return False, "You need to click Travelling before clicking On Site."
        if has_on_site:
            return False, "On Site has already been logged for this job."
    if action in {"No Access", "Revisit Required", "Completed"} and not has_on_site:
        return False, "You need to click On Site before selecting this option."
    return True, ""
