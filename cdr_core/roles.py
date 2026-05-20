AWAITING_DEPLOYMENT_STATUS = "Awaiting Dispatch"
LEGACY_AWAITING_DEPLOYMENT_STATUS = "Awaiting Deployment"
ASSIGNED_STATUS = "Assigned"
TRAVELLING_STATUS = "Travelling"
ON_SITE_STATUS = "On Site"
COMPLETED_STATUS = "Completed"

    
def user_can_use_helpdesk(role: str) -> bool:
    return str(role or "").strip().lower() in ["helpdesk", "admin"]


def role_counts_for_utilisation(role: str) -> bool:
    return str(role or "").strip().lower() == "engineer"


def is_vehicle_check_exempt_role(role: str) -> bool:
    return str(role or "").strip().lower() in ["admin", "apprentice", "helpdesk"]
