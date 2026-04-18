import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import json
import requests, re, time, datetime, msal
import sys, os

sys.path.append(os.path.dirname(__file__))
import config

st.set_page_config(
    page_title="📧 Email Search",
    page_icon="📧",
    layout="wide"
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ── SESSION STATE ─────────────────────────────────────────────
for k, v in [("token", None), ("device_flow", None), ("msal_app", None),
             ("msal_cache", None), ("token_cache", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── CONFIG AUTH ───────────────────────────────────────────────
def get_auth_config():
    return {
        "client_id": st.secrets["AZURE_CLIENT_ID"],
        "tenant_id": st.secrets["AZURE_TENANT_ID"],
        "scopes": ["Mail.Read", "Mail.ReadWrite", "User.Read"]
    }

# ── AUTH FUNCTIONS ────────────────────────────────────────────
def init_auth(cfg):
    cache = msal.SerializableTokenCache()
    if st.session_state.token_cache:
        cache.deserialize(st.session_state.token_cache)
    app = msal.PublicClientApplication(
        client_id=cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        token_cache=cache
    )
    return app, cache

def try_silent_auth(cfg):
    app, cache = init_auth(cfg)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(cfg["scopes"], account=accounts[0])
        if result and "access_token" in result:
            st.session_state.token_cache = cache.serialize()
            return result["access_token"]
    return None

def start_device_flow(cfg):
    app, cache = init_auth(cfg)
    flow = app.initiate_device_flow(scopes=cfg["scopes"])
    st.session_state.device_flow = flow
    st.session_state.msal_app = app
    st.session_state.msal_cache = cache
    return flow

def complete_device_flow():
    app   = st.session_state.msal_app
    flow  = st.session_state.device_flow
    cache = st.session_state.msal_cache
    if not app or not flow:
        return None
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        st.session_state.token_cache = cache.serialize()
        return result["access_token"]
    return None

# ── PAGE LOGIN ────────────────────────────────────────────────
def page_login(cfg):
    st.title("📧 Email Search")
    st.markdown("---")
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("### Connexion Office 365")

        token = try_silent_auth(cfg)
        if token:
            st.session_state.token = token
            st.rerun()

        if not st.session_state.device_flow:
            if st.button("Se connecter avec Office 365", use_container_width=True, type="primary"):
                with st.spinner("Initialisation..."):
                    start_device_flow(cfg)
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
        timeout=30
    )
    r.raise_for_status()
    return r.json()

def graph_get_all_pages(token, endpoint):
    results = []
    url = f"{GRAPH_BASE}{endpoint}"
    while url:
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results

# ── DOSSIERS ──────────────────────────────────────────────────
FOLDER_OPTIONS = {
    "📥 Boîte de réception":   "inbox",
    "📤 Éléments envoyés":     "sentitems",
    "🗑️ Éléments supprimés":  "deleteditems",
    "📁 Tous les dossiers":    "all",
    "📂 Courrier indésirable": "junkemail",
    "📝 Brouillons":           "drafts",
}

def search_emails_api(token, query, folder_key):
    select = "$select=subject,from,receivedDateTime,bodyPreview,isRead,toRecipients"
    if folder_key == "all":
        endpoint = f"/me/messages?$search=\"{query}\"&$top=999&{select}"
    else:
        endpoint = f"/me/mailFolders/{folder_key}/messages?$search=\"{query}\"&$top=999&{select}"
    return graph_get_all_pages(token, endpoint)

# ── VÉRIFICATION TOKEN ────────────────────────────────────────
cfg = get_auth_config()

if not st.session_state.token:
    page_login(cfg)
    st.stop()

token = st.session_state.token

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.title("📧 Email Search")
    st.markdown("---")
    try:
        me = graph_get(token, "/me")
        st.success(f"👤 {me.get('displayName', 'Utilisateur')}")
        st.caption(me.get('mail', ''))
    except:
        st.success("✅ Connecté")

    if st.button("🚪 Déconnexion"):
        for k in ["token", "token_cache", "device_flow"]:
            st.session_state[k] = None
        st.rerun()

    st.markdown("---")
    menu = st.radio("Navigation", [
        "🔍 Recherche",
        "📊 Statistiques",
        "🔄 Synchronisation",
        "🔧 Debug"
    ])
    st.markdown("---")
    try:
        conn  = sqlite3.connect(config.DB_PATH)
        count = conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0]
        conn.close()
        st.metric("📧 Emails indexés", f"{count:,}")
    except:
        st.warning("⚠️ Base non initialisée")

# ── PAGE RECHERCHE ────────────────────────────────────────────
if menu == "🔍 Recherche":
    st.title("🔍 Recherche d'emails")

    tab_local, tab_live = st.tabs(["🗄️ Recherche locale (SQLite)", "☁️ Recherche live (API)"])

    # ── RECHERCHE LOCALE ──────────────────────────────────────
    with tab_local:
        col1, col2 = st.columns([3, 1])
        with col1:
            query = st.text_input("Rechercher", placeholder="Ex: facture janvier...")
        with col2:
            search_in = st.multiselect(
                "Chercher dans",
                ["subject", "body", "sender"],
                default=["subject", "body"]
            )

        with st.expander("🔧 Filtres avancés"):
            col3, col4 = st.columns(2)
            with col3:
                date_debut = st.date_input("Date début", value=None, key="local_debut")
            with col4:
                date_fin   = st.date_input("Date fin",   value=None, key="local_fin")
            sender_filter = st.text_input("Expéditeur")

        if query:
            try:
                from src.search_engine import SearchEngine
                engine   = SearchEngine()
                keywords = [kw.strip() for kw in query.split() if kw.strip()]

                all_sets         = []
                all_results_map  = {}
                for keyword in keywords:
                    res = engine.search(keyword, fields=search_in)
                    all_sets.append(set(r['id'] for r in res))
                    for r in res:
                        all_results_map[r['id']] = r

                common_ids = all_sets[0].intersection(*all_sets[1:]) if all_sets else set()
                results    = [all_results_map[i] for i in common_ids]

                if sender_filter:
                    results = [r for r in results if
                               sender_filter.lower() in r['sender'].lower() or
                               sender_filter.lower() in r['sender_email'].lower()]
                if date_debut:
                    results = [r for r in results if r['date'][:10] >= str(date_debut)]
                if date_fin:
                    results = [r for r in results if r['date'][:10] <= str(date_fin)]

                st.markdown(f"🔑 **Mots clés :** {' | '.join(f'`{kw}`' for kw in keywords)}")
                st.markdown(f"**{len(results)} résultat(s) trouvé(s)**")
                st.markdown("---")

                for email in results:
                    with st.expander(f"📧 {email['subject']} | 👤 {email['sender']} | 📅 {email['date'][:10]}"):
                        st.markdown(f"**De :** {email['sender']} ({email['sender_email']})")
                        to_raw = email.get('to', '')
                        try:
                            to_list    = json.loads(to_raw)
                            to_display = ", ".join(
                                f"{r['emailAddress']['name']} <{r['emailAddress']['address']}>"
                                for r in to_list
                            )
                        except:
                            to_display = to_raw or "N/A"
                        st.markdown(f"**À :** {to_display}")
                        st.markdown(f"**Date :** {email['date']}")
                        st.markdown(f"**Aperçu :** {email['body_preview']}")
                        st.markdown(f"**Score :** {email['score']:.2f}")
                        if st.button("Voir email complet", key=email['id']):
                            detail = engine.get_email_detail(email['id'])
                            st.markdown("---")
                            st.markdown(detail['body'], unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Erreur : {e}")
                st.info("💡 Lancez d'abord une synchronisation !")

    # ── RECHERCHE LIVE API ────────────────────────────────────
    with tab_live:
        col1, col2, col3 = st.columns([3, 2, 1])
        with col1:
            live_query = st.text_input("Mot-clé", placeholder="Entrez un mot-clé...", key="live_q")
        with col2:
            folder_label = st.selectbox("📂 Dossier", list(FOLDER_OPTIONS.keys()))
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            search_btn = st.button("Rechercher", type="primary", use_container_width=True)

        if search_btn and live_query:
            folder_key = FOLDER_OPTIONS[folder_label]
            with st.spinner(f"Recherche dans « {folder_label} »..."):
                try:
                    emails = search_emails_api(token, live_query, folder_key)
                    if not emails:
                        st.info("Aucun email trouvé.")
                    else:
                        st.success(f"✅ {len(emails)} email(s) trouvé(s) dans {folder_label}")
                        for email in emails:
                            subject     = email.get("subject", "(sans objet)")
                            sender      = email.get("from", {}).get("emailAddress", {})
                            sender_name = sender.get("name", "")
                            sender_addr = sender.get("address", "")
                            date        = email.get("receivedDateTime", "")[:10]
                            preview     = email.get("bodyPreview", "")
                            is_read     = email.get("isRead", True)
                            unread      = "🔵 " if not is_read else ""
                            to_list     = email.get("toRecipients", [])
                            to_str      = ", ".join([r.get("emailAddress", {}).get("address", "") for r in to_list])

                            with st.expander(f"{unread}📧 **{subject}** — {sender_name} ({date})"):
                                st.markdown(f"**De :** {sender_name} <{sender_addr}>")
                                if to_str:
                                    st.markdown(f"**À :** {to_str}")
                                st.markdown(f"**Date :** {date}")
                                st.markdown(f"**Aperçu :** {preview}")
                except Exception as e:
                    st.error(f"Erreur : {e}")
        elif search_btn:
            st.warning("Entrez un mot-clé.")

# ── PAGE STATISTIQUES ─────────────────────────────────────────
elif menu == "📊 Statistiques":
    st.title("📊 Statistiques")
    try:
        conn = sqlite3.connect(config.DB_PATH)
        df   = pd.read_sql_query(
            "SELECT sender, sender_email, date, is_read, has_attachments FROM emails", conn
        )
        conn.close()

        df["date"] = pd.to_datetime(df["date"])
        df["mois"] = df["date"].dt.to_period("M").astype(str)

        col1, col2, col3 = st.columns(3)
        col1.metric("Total emails",        len(df))
        col2.metric("Emails non lus",      len(df[df["is_read"] == 0]))
        col3.metric("Avec pièces jointes", len(df[df["has_attachments"] == 1]))
        st.markdown("---")

        fig1 = px.bar(df.groupby("mois").size().reset_index(name="count"),
                      x="mois", y="count", title="📅 Emails reçus par mois")
        st.plotly_chart(fig1, use_container_width=True)

        top_senders = (df.groupby("sender_email").size()
                       .reset_index(name="count")
                       .sort_values("count", ascending=False)
                       .head(10))
        fig2 = px.bar(top_senders, x="count", y="sender_email",
                      orientation="h", title="👥 Top 10 expéditeurs")
        st.plotly_chart(fig2, use_container_width=True)

    except Exception as e:
        st.error(f"Erreur : {e}")

# ── PAGE SYNCHRONISATION ──────────────────────────────────────
elif menu == "🔄 Synchronisation":
    st.title("🔄 Synchronisation des emails")

    st.info("La synchronisation récupère **tous** vos emails sans limite via l'API Microsoft Graph.")

    if st.button("🚀 Lancer la synchronisation", type="primary"):
        try:
            from src.email_fetcher import EmailFetcher
            from src.indexer import EmailIndexer

            with st.spinner("📥 Récupération des emails (sans limite)..."):
                fetcher = EmailFetcher(token=token)        # passe le token device flow
                total   = fetcher.fetch_all_emails()       # plus de max_emails
                st.success(f"✅ {total} emails récupérés !")

            with st.spinner("🔍 Indexation en cours..."):
                indexer = EmailIndexer()
                indexed = indexer.index_all_emails()
                st.success(f"✅ {indexed} emails indexés !")

            st.balloons()

        except Exception as e:
            st.error(f"❌ Erreur : {e}")

# ── PAGE DEBUG ────────────────────────────────────────────────
elif menu == "🔧 Debug":
    st.title("🔧 Test connexion API")

    if st.button("Tester la connexion"):
        try:
            me = graph_get(token, "/me")
            st.success("✅ Token valide !")
            st.json(me)

            st.markdown("---")
            st.markdown("**Test récupération emails (5 premiers) :**")
            emails = graph_get(token, "/me/messages?$top=5&$select=subject,receivedDateTime,from")
            st.json(emails)

        except Exception as e:
            st.error(f"❌ Exception : {e}")
