from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Tuple

import pandas as pd
import requests

from boletin_core import consolidate_frames, read_normas_from_bytes

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_BASE = "https://login.microsoftonline.com"


@dataclass
class OneDriveConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    folder_path: str = "/"
    drive_id: str | None = None
    site_id: str | None = None
    site_hostname: str | None = None
    site_path: str | None = None
    user_id: str | None = None
    recursive: bool = False

    @property
    def is_configured(self) -> bool:
        base = bool(self.tenant_id and self.client_id and self.client_secret)
        target = bool(self.drive_id or self.site_id or (self.site_hostname and self.site_path) or self.user_id)
        return base and target


def _normalize_folder_path(path: str) -> str:
    path = (path or "/").strip()
    if not path or path == "/":
        return ""
    return path if path.startswith("/") else f"/{path}"


def get_access_token(config: OneDriveConfig) -> str:
    url = f"{TOKEN_BASE}/{config.tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": config.client_id,
        "client_secret": config.client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    resp = requests.post(url, data=payload, timeout=30)
    if not resp.ok:
        raise RuntimeError(f"No pude obtener token de Microsoft Graph: {resp.status_code} {resp.text[:500]}")
    data = resp.json()
    token = data.get("access_token")
    if not token:
        raise RuntimeError("Microsoft Graph no devolvió access_token.")
    return token


def graph_get(url: str, token: str) -> Dict[str, Any]:
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=60)
    if not resp.ok:
        raise RuntimeError(f"Error Microsoft Graph {resp.status_code}: {resp.text[:700]}")
    return resp.json()


def resolve_site_id(config: OneDriveConfig, token: str) -> str:
    if config.site_id:
        return config.site_id
    if not (config.site_hostname and config.site_path):
        raise RuntimeError("Falta SITE_ID o SITE_HOSTNAME + SITE_PATH para SharePoint.")
    site_path = config.site_path if config.site_path.startswith("/") else f"/{config.site_path}"
    url = f"{GRAPH_BASE}/sites/{config.site_hostname}:{site_path}"
    data = graph_get(url, token)
    site_id = data.get("id")
    if not site_id:
        raise RuntimeError("No pude resolver el SITE_ID de SharePoint.")
    return site_id


def root_children_url(config: OneDriveConfig, token: str) -> str:
    folder_path = _normalize_folder_path(config.folder_path)
    if config.drive_id:
        if folder_path:
            return f"{GRAPH_BASE}/drives/{config.drive_id}/root:{folder_path}:/children"
        return f"{GRAPH_BASE}/drives/{config.drive_id}/root/children"

    if config.site_id or (config.site_hostname and config.site_path):
        site_id = resolve_site_id(config, token)
        if folder_path:
            return f"{GRAPH_BASE}/sites/{site_id}/drive/root:{folder_path}:/children"
        return f"{GRAPH_BASE}/sites/{site_id}/drive/root/children"

    if config.user_id:
        if folder_path:
            return f"{GRAPH_BASE}/users/{config.user_id}/drive/root:{folder_path}:/children"
        return f"{GRAPH_BASE}/users/{config.user_id}/drive/root/children"

    raise RuntimeError("Configuración incompleta: indicá DRIVE_ID, SITE_ID/SITE_HOSTNAME+SITE_PATH o USER_ID.")


def list_children(url: str, token: str) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    next_url = url
    while next_url:
        data = graph_get(next_url, token)
        items.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return items


def list_excel_items(config: OneDriveConfig) -> List[Dict[str, Any]]:
    token = get_access_token(config)
    first_url = root_children_url(config, token)
    items = list_children(first_url, token)
    excel_items: List[Dict[str, Any]] = []

    def walk(current_items: List[Dict[str, Any]]) -> None:
        for item in current_items:
            name = item.get("name", "")
            if item.get("file") and not name.startswith("~$") and name.lower().endswith((".xlsx", ".xlsm")):
                excel_items.append(item)
            elif config.recursive and item.get("folder"):
                drive_id = item.get("parentReference", {}).get("driveId")
                item_id = item.get("id")
                if drive_id and item_id:
                    child_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
                    walk(list_children(child_url, token))

    walk(items)
    return excel_items


def download_item(item: Dict[str, Any]) -> bytes:
    download_url = item.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise RuntimeError(f"El archivo {item.get('name', '')} no tiene downloadUrl.")
    resp = requests.get(download_url, timeout=120)
    if not resp.ok:
        raise RuntimeError(f"No pude descargar {item.get('name', '')}: {resp.status_code}")
    return resp.content


def load_onedrive_folder(config: OneDriveConfig) -> Tuple[pd.DataFrame, List[str]]:
    if not config.is_configured:
        return pd.DataFrame(), ["Faltan datos de configuración de Microsoft Graph."]

    rows = []
    errors: List[str] = []
    try:
        items = list_excel_items(config)
    except Exception as exc:  # noqa: BLE001
        return pd.DataFrame(), [f"No pude listar la carpeta de OneDrive/SharePoint: {exc}"]

    for item in items:
        name = item.get("name", "archivo.xlsx")
        try:
            content = download_item(item)
            rows.append(read_normas_from_bytes(name, content))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{name}: {exc}")

    return consolidate_frames(rows), errors
