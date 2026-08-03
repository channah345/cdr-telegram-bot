"""Versioned payload contracts used between CDR interfaces."""

from dataclasses import asdict, dataclass


API_VERSION = "2026-08-03"


@dataclass(frozen=True)
class JobNotification:
    job_item_id: str
    cdr_number: str
    engineer_lookup_id: str
    chat_id: str
    text: str
    site_name: str = ""

    def payload(self):
        return {"api_version": API_VERSION, **asdict(self)}
