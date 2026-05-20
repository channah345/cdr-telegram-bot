import re
import requests
import msal
from .config import CDRConfig
from .fields import normalise_field_name

class SharePointClient:
    def __init__(self, config: CDRConfig | None = None):
        self.config = config or CDRConfig()
        self._app = msal.ConfidentialClientApplication(
            self.config.client_id,
            authority=self.config.authority,
            client_credential=self.config.client_secret,
        )
        self._site_id = None
        self._list_ids = {}
        self._drive_ids = {}

    def headers(self, content_type: bool = True) -> dict:
        token_result = self._app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in token_result:
            raise RuntimeError(f"Could not get Microsoft token: {token_result}")
        headers = {"Authorization": f"Bearer {token_result['access_token']}"}
        if content_type:
            headers["Content-Type"] = "application/json"
        return headers

    def site_id(self) -> str:
        if self._site_id:
            return self._site_id
        site = self.config.sharepoint_site
        site_hostname = site.split("/")[2]
        site_path = "/" + "/".join(site.split("/")[3:])
        url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
        response = requests.get(url, headers=self.headers(), timeout=30)
        if response.status_code != 200:
            raise RuntimeError(f"Could not get SharePoint site: {response.text}")
        self._site_id = response.json()["id"]
        return self._site_id

    def list_id(self, list_name: str) -> str:
        if list_name in self._list_ids:
            return self._list_ids[list_name]
        site_id = self.site_id()
        response = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists", headers=self.headers(), timeout=30)
        if response.status_code != 200:
            raise RuntimeError(f"Could not get lists: {response.text}")
        for lst in response.json().get("value", []):
            if lst.get("name") == list_name:
                self._list_ids[list_name] = lst["id"]
                return lst["id"]
        raise RuntimeError(f"List not found: {list_name}")

    def list_items(self, list_name_or_id: str, expand_fields: bool = True) -> list[dict]:
        site_id = self.site_id()
        list_id = self._list_ids.get(list_name_or_id) or list_name_or_id
        expand = "?expand=fields" if expand_fields else ""
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items{expand}"
        response = requests.get(url, headers=self.headers(), timeout=30)
        if response.status_code != 200:
            # Treat input as a list name if it was not a valid id.
            list_id = self.list_id(list_name_or_id)
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items{expand}"
            response = requests.get(url, headers=self.headers(), timeout=30)
        if response.status_code != 200:
            raise RuntimeError(f"Could not get items: {response.text}")
        return response.json().get("value", [])

    def list_columns(self, list_name_or_id: str) -> list[dict]:
        site_id = self.site_id()
        list_id = self._list_ids.get(list_name_or_id) or list_name_or_id
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"
        response = requests.get(url, headers=self.headers(), timeout=30)
        if response.status_code != 200:
            list_id = self.list_id(list_name_or_id)
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"
            response = requests.get(url, headers=self.headers(), timeout=30)
        if response.status_code != 200:
            raise RuntimeError(f"Could not get list columns: {response.text}")
        return response.json().get("value", [])

    def build_field_payload(self, list_name_or_id: str, fields: dict) -> dict:
        columns = self.list_columns(list_name_or_id)
        lookup = {"title": "Title"}
        for column in columns:
            internal_name = column.get("name", "")
            display_name = column.get("displayName", "")
            if column.get("readOnly") or internal_name in ["LinkTitle", "LinkTitleNoMenu"]:
                continue
            for key in [internal_name, display_name]:
                normalised = normalise_field_name(key)
                if normalised and normalised not in lookup:
                    lookup[normalised] = internal_name
        payload = {}
        for desired_name, value in (fields or {}).items():
            internal_name = "Title" if desired_name == "Title" else lookup.get(normalise_field_name(desired_name))
            if internal_name:
                payload[internal_name] = value
        return payload

    def update_item_fields(self, list_name_or_id: str, item_id: str, fields_to_update: dict):
        site_id = self.site_id()
        list_id = self._list_ids.get(list_name_or_id) or self.list_id(list_name_or_id)
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        payload = dict(fields_to_update or {})
        while True:
            response = requests.patch(url, headers=self.headers(), json=payload, timeout=30)
            if response.status_code in [200, 204]:
                return
            match = re.search(r"Field '([^']+)' is not recognized", response.text or "")
            if match and match.group(1) in payload:
                missing = match.group(1)
                payload.pop(missing, None)
                payload.pop(f"{missing}@odata.type", None)
                continue
            raise RuntimeError(f"Could not update item {item_id}: {response.text}")
