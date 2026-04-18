import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import requests, json, re, time, datetime, msal
import sys, os

sys.path.append(os.path.dirname(__file__))

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

st.set_page_config(
    page_title="📧 Email Search",
    page_icon="📧",
    layout="wide"
)

# ── SESSION STATE ─────────────────────────────────────────────
for k, v in [("token", None), ("device_flow", None), ("msal_app", None),
             ("msal_cache", None), ("token_cache", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── CONFIG ────────────────────────────────────────────────────
def get_config():
    return {
        "azure": {
            "client_id": st.secrets["AZURE_CLIENT_ID"],
            "tenant_id": st.secrets["AZURE_TENANT_ID"],
            "scopes": ["Mail.Read", "Mail.ReadWrite"]
        }
    }

# ── AUTH ──────────────────────────────────────────────────────
def init_auth(config):
    cache = msal.SerializableTokenCache()
    if st.session_state.token_cache:
        cache.deserialize(st.session_state.token_cache)
    app = msal.PublicClientApplication(
        client_id=config["azure"]["client_id"],
        authority=f"https://login.microsoftonline.com/{config['azure']['tenant_id']}",
        token_cache=cache
    )
    return app, cache

def try_silent_auth(config):
    app, cache = init_auth(config)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(config["azure"]["scopes"], account=accounts[0])
        if result and "access_token" in result:
            st.session_state.token_cache = cache.serialize()
            return result["access_token"]
    return None

def start_device_flow(config):
    app, cache = init_auth(config)
    flow = app.initiate_device_flow(scopes=config["azure"]["scopes"])
    st.session_state.device_flow = flow
    st.session_state.msal_app = app
    st.session_state.msal_cache = cache
    return flow

def complete_device_flow():
    app = st.session_state.msal_app
    flow = st.session_state.device_flow
    cache = st.session_state.msal_cache
    if not app or not flow:
        return None
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        st.session_state.token_cache = cache.serialize()
        return result["access_token"]
    return None

# ── PAGE LOGIN ────────────────────────────────────────────────
def page_login(config):
    st.title("📧 Email Search")
    st.markdown("---")
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("### Connexion Office 365")

        # Tentative silencieuse
        token = try_silent_auth(config)
        if token:
            st.session_state.token = token
            st.rerun()

        if not st.session_state.device_flow:
            if st.button("Se connecter avec Office 365", use_container_width=True, type="primary"):
                with st.spinner("Initialisation..."):
                    start_device_flow(config)
                st.rerun()
        else:
            flow = st.session_state.device_flow
            code_match = re.search(r'enter the code ([A-Z0-9]+)', flow.get("message", ""))
            code = code_match.group(1) if code_match else ""

            st.info("**Etape 1** — Ouvrez : [https://microsoft.com/devicelogin](https://microsoft.com/devicelogin)")
            st.info("**Etape 2** — Entrez ce code :")
            st.code(code, language=None)
            st.markdown("**Etape 3** — Connectez-vous puis cliquez :")

            if st.button("J'ai entré le code, continuer", use_container_width=True, type="primary"):
                with st.spinner("Vérification..."):
                    token = complete_device_flow()
                if token:
                    st.session_state.token = token
                    st.session_state.device_flow = None
                    st.rerun()
                else:
                    st.error("Échec — réessayez.")
                    st.session_state.device_flow = None

# ── GRAPH HELPERS ─────────────────────────────────────────────
def graph_get(token, endpoint):
    r = requests.get(
        f"{GRAPH_BASE}{endpoint}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=15
    )
    r.raise_for_status()
    return r.json()

def clean_html(html):
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', html)).strip()

# ── PAGES DE VOTRE APP ────────────────────────────────────────
def page_emails(token):
    st.title("📧 Recherche d'emails")
    
    col1, col2 = st.columns([3, 1])
    with col1:
        search_query = st.text_input("🔍 Rechercher dans les emails", placeholder="Entrez un mot-clé...")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        search_btn = st.button("Rechercher", type="primary", use_container_width=True)

    if search_btn and search_query:
        with st.spinner("Recherche en cours..."):
            try:
                data = graph_get(token, f"/me/messages?$search=\"{search_query}\"&$top=20&$select=subject,from,receivedDateTime,bodyPreview")
                emails = data.get("value", [])
                
                if not emails:
                    st.info("Aucun email trouvé.")
                else:
                    st.success(f"✅ {len(emails)} email(s) trouvé(s)")
                    for email in emails:
                        subject = email.get("subject", "(sans objet)")
                        sender = email.get("from", {}).get("emailAddress", {})
                        sender_name = sender.get("name", "")
                        sender_email = sender.get("address", "")
                        date = email.get("receivedDateTime", "")[:10]
                        preview = email.get("bodyPreview", "")

                        with st.expander(f"📧 **{subject}** — {sender_name} ({date})"):
                            st.markdown(f"**De :** {sender_name} <{sender_email}>")
                            st.markdown(f"**Date :** {date}")
                            st.markdown(f"**Aperçu :** {preview}")

            except Exception as e:
                st.error(f"Erreur : {e}")

def page_stats(token):
    st.title("📊 Statistiques")
    with st.spinner("Chargement..."):
        try:
            data = graph_get(token, "/me/messages?$top=50&$select=receivedDateTime,from,isRead")
            emails = data.get("value", [])
            
            if emails:
                df = pd.DataFrame([{
                    "date": e.get("receivedDateTime", "")[:10],
                    "sender": e.get("from", {}).get("emailAddress", {}).get("address", ""),
                    "isRead": e.get("isRead", False)
                } for e in emails])

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total emails", len(df))
                with col2:
                    st.metric("Non lus", len(df[df["isRead"] == False]))
                with col3:
                    st.metric("Lus", len(df[df["isRead"] == True]))

                fig = px.histogram(df, x="date", title="Emails par date")
                st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.error(f"Erreur : {e}")

# ── MAIN ──────────────────────────────────────────────────────
def main():
    config = get_config()

    # Sidebar
    with st.sidebar:
        st.title("📧 Email Search")
        st.markdown("---")

        if st.session_state.token:
            # Infos utilisateur
            try:
                me = graph_get(st.session_state.token, "/me")
                st.success(f"✅ {me.get('displayName', '')}")
                st.caption(me.get('mail', ''))
            except:
                st.success("✅ Connecté")

            st.markdown("---")
            page = st.radio("Navigation", ["📧 Emails", "📊 Statistiques"])
            st.markdown("---")

            if st.button("🚪 Se déconnecter"):
                for k in ["token", "token_cache", "device_flow"]:
                    st.session_state[k] = None
                st.rerun()
        else:
            st.info("Connectez-vous pour accéder à l'application.")
            page = "login"

    # Routing
    if not st.session_state.token:
        page_login(config)
    elif page == "📧 Emails":
        page_emails(st.session_state.token)
    elif page == "📊 Statistiques":
        page_stats(st.session_state.token)

if __name__ == "__main__":
    main()
