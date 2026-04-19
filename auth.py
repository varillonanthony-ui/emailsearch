"""
auth.py – Gestion de l'authentification Microsoft OAuth 2.0 via MSAL.
Seuls les utilisateurs du tenant configuré peuvent se connecter.
"""

import os
import msal
import requests
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID     = os.getenv("AZURE_TENANT_ID")
REDIRECT_URI  = os.getenv("REDIRECT_URI", "http://localhost:8501")

# Autorité spécifique au tenant → seuls les comptes du tenant peuvent se connecter
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/User.Read",
    "offline_access",
]


# ── Helpers MSAL ──────────────────────────────────────────────────────────────

def _build_app(cache: msal.SerializableTokenCache | None = None):
    """Crée un ConfidentialClientApplication avec cache optionnel."""
    if cache is None:
        cache = msal.SerializableTokenCache()
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )
    return app, cache


# ── Auth URL ──────────────────────────────────────────────────────────────────

def get_auth_url() -> str:
    """Génère l'URL d'autorisation Microsoft pour la page de connexion."""
    app, _ = _build_app()
    return app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
    )


# ── Échange code → token ──────────────────────────────────────────────────────

def acquire_token_by_code(code: str) -> tuple[dict, str]:
    """
    Échange un code d'autorisation contre un access token.
    Retourne (result_dict, serialized_cache).
    """
    app, cache = _build_app()
    result = app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    return result, cache.serialize()


# ── Refresh silencieux ────────────────────────────────────────────────────────

def acquire_token_silent(cache_state: str) -> tuple[dict | None, str]:
    """
    Rafraîchit silencieusement le token depuis le cache.
    Retourne (result_dict_or_None, new_serialized_cache).
    """
    cache = msal.SerializableTokenCache()
    cache.deserialize(cache_state)
    app, cache = _build_app(cache)

    accounts = app.get_accounts()
    if not accounts:
        return None, cache_state

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    return result, cache.serialize()


# ── Infos utilisateur ─────────────────────────────────────────────────────────

def get_user_info(access_token: str) -> dict | None:
    """Récupère le profil utilisateur via Microsoft Graph (/me)."""
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=15,
    )
    if resp.status_code == 200:
        return resp.json()
    return None