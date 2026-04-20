"""
auth.py – Authentification Microsoft via Device Code Flow multi-tenant.

Configuration Azure AD requise :
  - App Registration → Authentification → Types de comptes pris en charge
    → "Comptes dans un annuaire organisationnel (tout Azure AD — multi-tenant)"
  - Secrets Streamlit : AZURE_CLIENT_ID + APP_PASSWORD (plus besoin de AZURE_TENANT_ID)

L'endpoint "organizations" accepte n'importe quel tenant Microsoft 365,
ce qui permet de connecter plusieurs boîtes de domaines différents.
"""

import requests

SCOPES   = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/User.Read offline_access"
ENDPOINT = "https://login.microsoftonline.com/organizations/oauth2/v2.0"


def start_device_flow(client_id: str) -> dict:
    """Démarre le device code flow (multi-tenant, accepte tout domaine M365)."""
    r = requests.post(
        f"{ENDPOINT}/devicecode",
        data={"client_id": client_id, "scope": SCOPES},
        timeout=15,
    )
    r.raise_for_status()
    return r.json()


def poll_token(client_id: str, device_code: str) -> tuple[str | None, str | None]:
    """Interroge le token endpoint. Retourne (access_token, refresh_token)."""
    r = requests.post(
        f"{ENDPOINT}/token",
        data={
            "client_id":   client_id,
            "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
        },
        timeout=15,
    )
    d = r.json()
    return d.get("access_token"), d.get("refresh_token")


def refresh_access_token(client_id: str, refresh_token: str) -> str | None:
    """Rafraîchit silencieusement un access token."""
    r = requests.post(
        f"{ENDPOINT}/token",
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
    """Récupère le profil Microsoft Graph (/me)."""
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=15,
    )
    return r.json() if r.status_code == 200 else None