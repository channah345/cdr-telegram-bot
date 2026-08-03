import pytest

from cdr_core.workflows import aggregate_live_status, progress_update, validate_action, worksheet_update


def fields_with_two_engineers():
    return {"Status": "Assigned", "EngineerLookupId": ["1", "2"], "EngineerVisitLog": ""}


def log_progress(fields, engineer, action, minute):
    updated = dict(fields)
    updated.update(progress_update(updated, engineer, action, f"03/08/2026 09:{minute:02d}"))
    return updated


def test_actions_are_ordered_and_idempotent():
    fields = fields_with_two_engineers()
    assert validate_action(fields, "Alice", "On Site")[0] is False
    fields = log_progress(fields, "Alice", "Travelling", 0)
    assert validate_action(fields, "Alice", "Travelling")[0] is False
    fields = log_progress(fields, "Alice", "On Site", 20)
    assert validate_action(fields, "Alice", "Undo Travelling")[0] is False
    assert validate_action(fields, "Alice", "Completed")[0] is True


def test_undo_is_per_engineer_and_preserves_other_engineer_status():
    fields = fields_with_two_engineers()
    fields = log_progress(fields, "Alice", "Travelling", 0)
    fields = log_progress(fields, "Bob", "Travelling", 1)
    fields.update(progress_update(fields, "Alice", "Undo Travelling", "03/08/2026 09:05"))
    assert aggregate_live_status(fields["EngineerVisitLog"]) == "Travelling"
    assert validate_action(fields, "Alice", "Travelling")[0] is True
    assert validate_action(fields, "Bob", "Travelling")[0] is False


def test_multi_engineer_worksheet_holds_job_until_final_submission():
    fields = fields_with_two_engineers()
    for engineer, base in [("Alice", 0), ("Bob", 2)]:
        fields = log_progress(fields, engineer, "Travelling", base)
        fields = log_progress(fields, engineer, "On Site", base + 10)
    first, is_final = worksheet_update(fields, "1", "Alice", "Completed", "03/08/2026 10:00", "Completed Alice's part", "Cable", False)
    assert is_final is False
    assert first["EngineerLookupId"] == [2]
    assert first["Status"] == "On Site"
    assert "JobOutcome" not in first
    latest = dict(fields); latest.update(first)
    second, is_final = worksheet_update(latest, "2", "Bob", "Completed", "03/08/2026 10:10", "Completed Bob's part", "", "No")
    assert is_final is True
    assert second["EngineerLookupId"] == []
    assert second["Status"] == "Completed"


def test_no_access_returns_final_engineer_to_dispatch_and_requires_reason():
    fields = {"Status": "Assigned", "EngineerLookupId": [1], "EngineerVisitLog": ""}
    fields = log_progress(fields, "Alice", "Travelling", 0)
    fields = log_progress(fields, "Alice", "On Site", 10)
    with pytest.raises(ValueError, match="No Access reason"):
        worksheet_update(fields, "1", "Alice", "No Access", "03/08/2026 09:20", "")
    payload, is_final = worksheet_update(fields, "1", "Alice", "No Access", "03/08/2026 09:20", "", no_access_reason="Customer unavailable")
    assert is_final is True
    assert payload["Status"] == "Awaiting Dispatch"
    assert payload["TelegramNotified"] is False


def test_previous_revisit_does_not_block_new_visit():
    fields = {"Status": "Assigned", "JobOutcome": "Revisit Required", "EngineerLookupId": [1], "EngineerVisitLog": "01/08/2026 09:00 - Alice - Travelling\n01/08/2026 09:20 - Alice - On Site\n01/08/2026 09:40 - Alice - Revisit Required"}
    assert validate_action(fields, "Alice", "Travelling") == (True, "")
