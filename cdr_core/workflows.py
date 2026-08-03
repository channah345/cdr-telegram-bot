"""Canonical engineer workflow rules shared by the portal and Telegram bot.

This module is intentionally free of FastAPI, Telegram and SharePoint calls.  Both
interfaces can therefore make the same decision before they persist a change.
"""

from __future__ import annotations

from datetime import datetime

from .jobs import get_assigned_engineer_ids, is_closed_job
from .statuses import (
    ASSIGNED_STATUS,
    AWAITING_DEPLOYMENT_STATUS,
    COMPLETED_STATUS,
    ON_SITE_STATUS,
    TRAVELLING_STATUS,
)


RESET_ACTIONS = {
    "Completed",
    "No Access",
    "Revisit Required",
    "Aborted Attendance",
    "Travelling Reverted",
}
WORKSHEET_OUTCOMES = {"Completed", "No Access", "Revisit Required"}


def as_bool(value) -> bool:
    return value is True or str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def append_visit_log(fields: dict, engineer_name: str, action: str, timestamp: str, extra: str = "") -> str:
    existing = str(fields.get("EngineerVisitLog", "") or "").strip()
    line = f"{timestamp} - {engineer_name} - {action}"
    if str(extra or "").strip():
        line += f" - {str(extra).strip()}"
    return f"{existing}\n{line}" if existing else line


def current_visit_lines(fields: dict, engineer_name: str) -> list[str]:
    """Return only this engineer's open visit, retaining the full audit log."""
    lines = [line for line in str(fields.get("EngineerVisitLog", "") or "").splitlines() if line.strip()]
    marker = f" - {engineer_name} - "
    last_reset = -1
    for index, line in enumerate(lines):
        if marker not in line:
            continue
        action_text = line.split(marker, 1)[1].split(" - ", 1)[0].strip()
        if action_text in RESET_ACTIONS or action_text.startswith("Submitted for Helpdesk Review"):
            last_reset = index
    return lines[last_reset + 1 :]


def has_action(fields: dict, engineer_name: str, action: str) -> bool:
    marker = f" - {engineer_name} - {action}"
    return any(marker in line for line in current_visit_lines(fields, engineer_name))


def validate_action(fields: dict, engineer_name: str, action: str) -> tuple[bool, str]:
    if is_closed_job(fields):
        return False, "This job has already been closed or returned to the office."

    travelled = has_action(fields, engineer_name, "Travelling")
    on_site = has_action(fields, engineer_name, "On Site")

    if action == "Travelling":
        return (False, "Travelling has already been logged for this visit.") if travelled else (True, "")
    if action == "Undo Travelling":
        if not travelled:
            return False, "There is no Travelling action to undo."
        if on_site:
            return False, "Travelling cannot be undone after On Site has been logged."
        return True, ""
    if action == "On Site":
        if not travelled:
            return False, "You need to log Travelling before On Site."
        return (False, "On Site has already been logged for this visit.") if on_site else (True, "")
    if action == "Aborted Attendance":
        return True, ""
    if action in WORKSHEET_OUTCOMES:
        if not on_site:
            return False, "You need to log On Site before selecting this outcome."
        return True, ""
    return False, "Unsupported job action."


def aggregate_live_status(log_text: str) -> str:
    """Calculate the shared job status from every engineer's latest open visit."""
    states: dict[str, str] = {}
    for raw_line in str(log_text or "").splitlines():
        parts = raw_line.strip().split(" - ", 3)
        if len(parts) < 3:
            continue
        engineer = parts[1].strip().lower()
        action = parts[2].strip()
        if not engineer:
            continue
        if action == "Travelling":
            states[engineer] = TRAVELLING_STATUS
        elif action == "On Site":
            states[engineer] = ON_SITE_STATUS
        elif action in RESET_ACTIONS or action.startswith("Submitted for Helpdesk Review"):
            states.pop(engineer, None)
    if ON_SITE_STATUS in states.values():
        return ON_SITE_STATUS
    if TRAVELLING_STATUS in states.values():
        return TRAVELLING_STATUS
    return ASSIGNED_STATUS


def progress_update(fields: dict, engineer_name: str, action: str, timestamp: str) -> dict:
    """Build a Travelling, On Site or engineer-specific undo update."""
    ok, message = validate_action(fields, engineer_name, action)
    if not ok:
        raise ValueError(message)
    logged_action = "Travelling Reverted" if action == "Undo Travelling" else action
    extra = "Reverted by engineer" if action == "Undo Travelling" else ""
    updated_log = append_visit_log(fields, engineer_name, logged_action, timestamp, extra)
    return {
        "EngineerVisitLog": updated_log,
        "Status": aggregate_live_status(updated_log),
    }


def remaining_engineer_ids(fields: dict, current_lookup_id: str) -> list[int]:
    current = str(current_lookup_id or "").strip()
    result: list[int] = []
    for lookup_id in get_assigned_engineer_ids(fields):
        value = str(lookup_id or "").strip()
        if value and value != current and value.isdigit() and int(value) not in result:
            result.append(int(value))
    return result


def assignment_payload(ids: list[int]) -> dict:
    return {
        "EngineerLookupId@odata.type": "Collection(Edm.Int32)",
        "EngineerLookupId": list(ids),
    }


def worksheet_update(
    fields: dict,
    engineer_lookup_id: str,
    engineer_name: str,
    outcome: str,
    timestamp: str,
    work_completed: str,
    materials_used: str = "",
    follow_on_required=False,
    follow_on_notes: str = "",
    no_access_reason: str = "",
    client_signature_required=False,
    client_signature_received=False,
    signature_link: str = "",
    worksheet_link: str = "",
    visit_comment: str = "",
) -> tuple[dict, bool]:
    """Build the canonical close-out payload and preserve other engineers."""
    if outcome not in WORKSHEET_OUTCOMES:
        raise ValueError("Unsupported worksheet outcome.")
    ok, message = validate_action(fields, engineer_name, outcome)
    if not ok:
        raise ValueError(message)
    if not str(work_completed or "").strip() and outcome != "No Access":
        raise ValueError("Work completed is required.")
    if outcome == "No Access" and not str(no_access_reason or "").strip():
        raise ValueError("Select or enter a No Access reason.")
    follow_on = as_bool(follow_on_required)
    if follow_on and not str(follow_on_notes or "").strip():
        raise ValueError("Add the follow-on work required.")

    remaining = remaining_engineer_ids(fields, engineer_lookup_id)
    is_final = not remaining
    updated_log = append_visit_log(fields, engineer_name, outcome, timestamp, visit_comment)
    payload = {
        "WorkCompleted": str(work_completed or "").strip(),
        "MaterialsUsed": str(materials_used or "").strip(),
        "FollowOnRequired": follow_on,
        "FollowOnNotes": str(follow_on_notes or "").strip(),
        "WorksheetSubmitted": True,
        "EngineerVisitLog": updated_log,
        "ClientSignatureRequired": as_bool(client_signature_required),
        "ClientSignatureReceived": as_bool(client_signature_received),
    }
    if no_access_reason:
        payload["NoAccessReason"] = str(no_access_reason).strip()
    if signature_link:
        payload["ClientSignatureLink"] = signature_link
    if worksheet_link:
        payload.update({"WorksheetPDFLink": worksheet_link, "WorksheetGenerated": True})

    if is_final:
        payload["JobOutcome"] = outcome
        payload["Status"] = COMPLETED_STATUS if outcome == "Completed" else AWAITING_DEPLOYMENT_STATUS
        payload["TelegramNotified"] = outcome == "Completed"
        payload.update(assignment_payload([]))
    else:
        payload.update(assignment_payload(remaining))
        payload["Status"] = aggregate_live_status(updated_log)
    return payload, is_final


def visit_times(fields: dict, engineer_name: str, now: datetime | None = None) -> dict[str, str]:
    """Extract the current visit times for the final engineer review screen."""
    result = {"travel": "", "on_site": "", "off_site": ""}
    marker = f" - {engineer_name} - "
    for line in current_visit_lines(fields, engineer_name):
        if marker not in line:
            continue
        prefix, action_extra = line.split(marker, 1)
        time_value = prefix.strip().split(" ")[-1]
        action = action_extra.split(" - ", 1)[0].strip()
        if action == "Travelling":
            result["travel"] = time_value
        elif action == "On Site":
            result["on_site"] = time_value
    result["off_site"] = (now or datetime.now()).strftime("%H:%M")
    return result
