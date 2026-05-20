"""Shared CDR Core helpers used by CDR bot, operations portal and future systems."""

from .config import CDRConfig
from .dates import UK_TZ, now_log_time, graph_datetime_now, format_sharepoint_date, sharepoint_date_to_uk_date
from .sharepoint import SharePointClient
from .fields import normalise_field_name, get_field_value, bool_field, normalise_cdr
from .jobs import (
    AWAITING_DEPLOYMENT_STATUS,
    LEGACY_AWAITING_DEPLOYMENT_STATUS,
    ASSIGNED_STATUS,
    TRAVELLING_STATUS,
    ON_SITE_STATUS,
    COMPLETED_STATUS,
    get_assigned_engineer_ids,
    is_closed_job,
    validate_job_action,
)
from .roles import user_can_use_helpdesk, role_counts_for_utilisation, is_vehicle_check_exempt_role
