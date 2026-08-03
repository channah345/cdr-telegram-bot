"""Pooled, paginated Microsoft Graph client with async adapters."""

import asyncio
import re
import threading
from urllib.parse import urlencode

import msal
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from .config import CDRConfig
from .fields import normalise_field_name


class SharePointClient:
    def __init__(self, config: CDRConfig | None = None):
        self.config = config or CDRConfig()
        self._app = None
        self._app_lock = threading.RLock()
        self._site_id = None
        self._list_ids = {}
        self._columns = {}
        self.session = requests.Session()
        self.session.mount("https://", HTTPAdapter(
            pool_connections=20, pool_maxsize=20,
            max_retries=Retry(total=3, connect=3, read=2, backoff_factor=.4,
                              status_forcelist=(429, 502, 503, 504),
                              allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
                              respect_retry_after_header=True),
        ))

    def _msal(self):
        if self._app is None:
            with self._app_lock:
                if self._app is None:
                    self._app = msal.ConfidentialClientApplication(
                        self.config.client_id, authority=self.config.authority,
                        client_credential=self.config.client_secret,
                    )
        return self._app

    def headers(self, content_type: bool = True) -> dict:
        token_result = self._msal().acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in token_result:
            raise RuntimeError(f"Could not get Microsoft token: {token_result}")
        headers = {"Authorization": f"Bearer {token_result['access_token']}"}
        if content_type:
            headers["Content-Type"] = "application/json"
        return headers

    def _request(self, method, url, **kwargs):
        kwargs.setdefault("timeout", (10, 60))
        response = self.session.request(method, url, headers=self.headers(), **kwargs)
        if response.status_code >= 400:
            raise RuntimeError(f"Microsoft Graph {method} failed ({response.status_code}): {response.text}")
        return response

    def site_id(self) -> str:
        if self._site_id:
            return self._site_id
        site = self.config.sharepoint_site
        if not site:
            raise RuntimeError("SHAREPOINT_SITE is not configured")
        site_hostname = site.split("/")[2]
        site_path = "/" + "/".join(site.split("/")[3:])
        response = self._request("GET", f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}")
        self._site_id = response.json()["id"]
        return self._site_id

    def list_id(self, list_name: str) -> str:
        if list_name in self._list_ids:
            return self._list_ids[list_name]
        response = self._request("GET", f"https://graph.microsoft.com/v1.0/sites/{self.site_id()}/lists")
        for item in response.json().get("value", []):
            if item.get("name") == list_name:
                self._list_ids[list_name] = item["id"]
                return item["id"]
        raise RuntimeError(f"List not found: {list_name}")

    def list_items(self, list_name_or_id: str, *, filter_query: str = "", select_fields: str = "") -> list[dict]:
        list_id = self._list_ids.get(list_name_or_id) or list_name_or_id
        params = {"$expand": "fields", "$top": "999"}
        if filter_query:
            params["$filter"] = filter_query
        if select_fields:
            params["$select"] = select_fields
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id()}/lists/{list_id}/items?{urlencode(params)}"
        items = []
        while url:
            response = self._request("GET", url)
            payload = response.json()
            items.extend(payload.get("value", []))
            url = payload.get("@odata.nextLink")
        return items

    def list_columns(self, list_name_or_id: str) -> list[dict]:
        list_id = self._list_ids.get(list_name_or_id) or list_name_or_id
        if list_id not in self._columns:
            response = self._request("GET", f"https://graph.microsoft.com/v1.0/sites/{self.site_id()}/lists/{list_id}/columns")
            self._columns[list_id] = response.json().get("value", [])
        return self._columns[list_id]

    def build_field_payload(self, list_name_or_id: str, fields: dict) -> dict:
        lookup = {"title": "Title"}
        for column in self.list_columns(list_name_or_id):
            internal = column.get("name", "")
            if column.get("readOnly") or internal in {"LinkTitle", "LinkTitleNoMenu"}:
                continue
            for key in [internal, column.get("displayName", "")]:
                normalised = normalise_field_name(key)
                if normalised and normalised not in lookup:
                    lookup[normalised] = internal
        return {
            ("Title" if name == "Title" else lookup[normalise_field_name(name)]): value
            for name, value in (fields or {}).items()
            if name == "Title" or normalise_field_name(name) in lookup
        }

    def update_item_fields(self, list_name_or_id: str, item_id: str, fields_to_update: dict):
        list_id = self._list_ids.get(list_name_or_id) or list_name_or_id
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id()}/lists/{list_id}/items/{item_id}/fields"
        payload = dict(fields_to_update or {})
        while payload:
            response = self.session.patch(url, headers=self.headers(), json=payload, timeout=(10, 60))
            if response.status_code in {200, 204}:
                return
            match = re.search(r"Field '([^']+)' is not recognized", response.text or "")
            if match and match.group(1) in payload:
                missing = match.group(1)
                payload.pop(missing, None)
                payload.pop(f"{missing}@odata.type", None)
                continue
            raise RuntimeError(f"Could not update item {item_id}: {response.text}")

    async def async_site_id(self):
        return await asyncio.to_thread(self.site_id)

    async def async_list_id(self, list_name):
        return await asyncio.to_thread(self.list_id, list_name)

    async def async_list_items(self, list_name_or_id, **kwargs):
        return await asyncio.to_thread(self.list_items, list_name_or_id, **kwargs)

    async def async_update_item_fields(self, list_name_or_id, item_id, fields):
        return await asyncio.to_thread(self.update_item_fields, list_name_or_id, item_id, fields)
