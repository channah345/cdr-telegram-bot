from dataclasses import dataclass
import os

@dataclass(frozen=True)
class CDRConfig:
    tenant_id: str = os.getenv("TENANT_ID", "")
    client_id: str = os.getenv("CLIENT_ID", "")
    client_secret: str = os.getenv("CLIENT_SECRET", "")
    sharepoint_site: str = os.getenv("SHAREPOINT_SITE", "")
    photo_library: str = os.getenv("PHOTO_LIBRARY", "Documents")

    jobs_list: str = os.getenv("JOBS_LIST", "Engineer Jobs")
    engineers_list: str = os.getenv("ENGINEERS_LIST", "Engineers")
    day_logs_list: str = os.getenv("DAY_LOGS_LIST", "Engineer Day Logs")
    task_activities_list: str = os.getenv("TASK_ACTIVITIES_LIST", "Task Activities")

    photo_base_folder: str = os.getenv("PHOTO_BASE_FOLDER", "15 - ENGINEER JOB PHOTOS")
    signature_base_folder: str = os.getenv("SIGNATURE_BASE_FOLDER", "16 - CLIENT SIGNATURES")
    van_check_photo_base_folder: str = os.getenv("VAN_CHECK_PHOTO_BASE_FOLDER", "17 - VAN CHECK PHOTOS")
    worksheet_base_folder: str = os.getenv("WORKSHEET_BASE_FOLDER", "18 - JOB WORKSHEETS")
    receipt_base_folder: str = os.getenv("RECEIPT_BASE_FOLDER", "19 - RECEIPTS")

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"
