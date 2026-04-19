"""
auth.py – Authentification en deux étapes :
  1. Mot de passe applicatif (APP_PASSWORD dans les secrets Streamlit)
  2. Device Code Flow Microsoft — aucun redirect URI requis,
     fonctionne parfaitement dans les iframes Streamlit Cloud.
"""

import requests

TENANT_ID_KEY  = "AZURE_TENANT_ID"
CLIENT_ID_KEY  = "AZURE_CLIENT_ID"
PASSWORD_KEY   = "APP_PASSWORD"

SCOPES = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/User.Read offline_access"

# ── Device Code Flow ──────────────────────────────────────────────────────────

def start_device_flow(tenant_id: str, client_id: str) -> dict:
    """
    Démarre le device code flow.
    Retourne le dict contenant 'user_code', 'verification_uri', 'device_code', 'expires_in'.
    """
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode",
        data={"client_id": client_id, "scope": SCOPES},
        timeout=15,
    )
    r.raise_for_status()
    return r.json()


def poll_token(tenant_id: str, client_id: str, device_code: str) -> tuple[str | None, str | None]:
    """
    Interroge le endpoint token.
    Retourne (access_token, refresh_token) ou (None, None) si pas encore validé.
    """
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id":   client_id,
            "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
        },
        timeout=15,
    )
    data = r.json()
    return data.get("access_token"), data.get("refresh_token")


def refresh_access_token(tenant_id: str, client_id: str, refresh_token: str) -> str | None:
    """Rafraîchit silencieusement l'access token avec le refresh token."""
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
    data = r.json()
    return data.get("access_token")


# ── Infos utilisateur ─────────────────────────────────────────────────────────

def get_user_info(access_token: str) -> dict | None:
    """Récupère le profil Microsoft Graph (/me)."""
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=15,
    )
    return r.json() if r.status_code == 200 else None