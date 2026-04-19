"""
auth.py – Authentification Microsoft via Device Code Flow.
Aucun redirect URI requis — fonctionne dans les iframes Streamlit Cloud.
"""

import requests

SCOPES = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/User.Read offline_access"


def start_device_flow(tenant_id: str, client_id: str) -> dict:
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode",
        data={"client_id": client_id, "scope": SCOPES},
        timeout=15,
    )
    r.raise_for_status()
    return r.json()


def poll_token(tenant_id: str, client_id: str, device_code: str) -> tuple[str | None, str | None]:
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id":   client_id,
            "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
        },
        timeout=15,
    )
    d = r.json()
    return d.get("access_token"), d.get("refresh_token")


def refresh_access_token(tenant_id: str, client_id: str, refresh_token: str) -> str | None:
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id":     client_id,
            "grant_type":    "refresh_token",
            "refresh_token": refresh_token,
            "scope":         SCOPES,
        },
        timeout=15,
    )
    return r.json().get("access_token")


def get_user_info(access_token: str) -> dict | None:
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=15,
    )
    return r.json() if r.status_code == 200 else None