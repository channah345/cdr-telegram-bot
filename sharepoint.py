def normalise_field_name(value) -> str:
    return "".join(str(value or "").lower().replace("_x0020_", "").split())

def get_field_value(fields: dict, *names):
    fields = fields or {}
    for name in names:
        if name in fields and fields.get(name) not in [None, ""]:
            return fields.get(name)
    wanted = {normalise_field_name(name) for name in names}
    for key, value in fields.items():
        if normalise_field_name(key) in wanted and value not in [None, ""]:
            return value
    return None

def bool_field(value) -> bool:
    return value in [True, "true", "True", "Yes", "yes", "1", 1]

def normalise_cdr(value) -> str:
    value = str(value or "").strip().lower()
    for prefix in ["cdr:", "cdr number:", "cdrnumber:"]:
        if value.startswith(prefix):
            value = value[len(prefix):].strip()
    return value.replace(" ", "").replace("-", "").replace("_", "")
