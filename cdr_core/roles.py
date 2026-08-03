"""Role rules shared by all interfaces."""


def normalise_role(role: str) -> str:
    return str(role or "").strip().lower()


def user_can_use_helpdesk(role: str) -> bool:
    return normalise_role(role) in {"helpdesk", "admin"}


def role_counts_for_utilisation(role: str) -> bool:
    return normalise_role(role) == "engineer"


def is_vehicle_check_exempt_role(role: str) -> bool:
    return normalise_role(role) in {"admin", "apprentice", "helpdesk"}
