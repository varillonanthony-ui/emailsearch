import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import json
import sys, os

sys.path.append(os.path.dirname(__file__))
import config

st.set_page_config(
    page_title="📧 Email Search",
    page_icon="📧",
    layout="wide"
)

# ── AUTHENTIFICATION MICROSOFT ───────────────────────────────
from msal import ConfidentialClientApplication

REDIRECT_URI = "https://votre-app.streamlit.app"  # ← changez ici

def get_auth_url():
    app = ConfidentialClientApplication(
        config.CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
        client_credential=config.CLIENT_SECRET
    )
    return app.get_authorization_request_url(
        scopes=["User.Read"],
        redirect_uri=REDIRECT_URI
    )

# Vérifier si l'utilisateur est connecté
if "user" not in st.session_state:
    st.title("🔐 Connexion requise")
    st.markdown("Connectez-vous avec votre compte Microsoft pour accéder à l'application.")
    
    auth_url = get_auth_url()
    st.markdown(f'<a href="{auth_url}" target="_self"><button style="background-color:#0078d4;color:white;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-size:16px;">🔑 Se connecter avec Microsoft</button></a>', unsafe_allow_html=True)
    
    # Récupérer le code de retour
    params = st.query_params
    if "code" in params:
        with st.spinner("Connexion en cours..."):
            try:
                import requests
                msal_app = ConfidentialClientApplication(
                    config.CLIENT_ID,
                    authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
                    client_credential=config.CLIENT_SECRET
                )
                result = msal_app.acquire_token_by_authorization_code(
                    params["code"],
                    scopes=["User.Read"],
                    redirect_uri=REDIRECT_URI
                )
                if "access_token" in result:
                    # Récupérer infos utilisateur
                    headers = {"Authorization": f"Bearer {result['access_token']}"}
                    me = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers).json()
                    st.session_state["user"] = me
                    st.query_params.clear()
                    st.rerun()
                else:
                    st.error(f"❌ Erreur : {result.get('error_description')}")
            except Exception as e:
                st.error(f"❌ Exception : {e}")
    st.stop()

# ── UTILISATEUR CONNECTÉ ─────────────────────────────────────
user = st.session_state["user"]

# ── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.title("📧 Email Search")
    st.markdown("---")
    st.success(f"👤 {user.get('displayName', 'Utilisateur')}")
    st.caption(user.get('mail', ''))
    if st.button("🚪 Déconnexion"):
        del st.session_state["user"]
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

# ── PAGE RECHERCHE ───────────────────────────────────────────
if menu == "🔍 Recherche":
    st.title("🔍 Recherche d'emails")

    col1, col2 = st.columns([3, 1])
    with col1:
        query = st.text_input(
            "Rechercher",
            placeholder="Ex: facture janvier, réunion projet..."
        )
    with col2:
        search_in = st.multiselect(
            "Chercher dans",
            ["subject", "body", "sender"],
            default=["subject", "body"]
        )

    with st.expander("🔧 Filtres avancés"):
        col3, col4 = st.columns(2)
        with col3:
            date_debut = st.date_input("Date début", value=None)
        with col4:
            date_fin   = st.date_input("Date fin", value=None)
        sender_filter = st.text_input("Expéditeur")

    if query:
        try:
            from src.search_engine import SearchEngine
            engine = SearchEngine()

            keywords = [kw.strip() for kw in query.split() if kw.strip()]

            all_sets = []
            all_results_map = {}
            for keyword in keywords:
                res = engine.search(keyword, fields=search_in)
                all_sets.append(set(r['id'] for r in res))
                for r in res:
                    all_results_map[r['id']] = r

            if all_sets:
                common_ids = all_sets[0].intersection(*all_sets[1:])
            else:
                common_ids = set()

            results = [all_results_map[id] for id in common_ids]

            if sender_filter:
                results = [
                    r for r in results
                    if sender_filter.lower() in r['sender'].lower()
                    or sender_filter.lower() in r['sender_email'].lower()
                ]

            if date_debut:
                results = [
                    r for r in results
                    if r['date'][:10] >= str(date_debut)
                ]

            if date_fin:
                results = [
                    r for r in results
                    if r['date'][:10] <= str(date_fin)
                ]

            st.markdown(f"🔑 **Mots clés :** {' | '.join(f'`{kw}`' for kw in keywords)}")
            st.markdown(f"**{len(results)} résultat(s) trouvé(s)**")
            st.markdown("---")

            for email in results:
                with st.expander(
                    f"📧 {email['subject']} | 👤 {email['sender']} | 📅 {email['date'][:10]}"
                ):
                    st.markdown(f"**De :** {email['sender']} ({email['sender_email']})")

                    to_raw = email.get('to', '')
                    try:
                        to_list = json.loads(to_raw)
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

# ── PAGE STATISTIQUES ────────────────────────────────────────
elif menu == "📊 Statistiques":
    st.title("📊 Statistiques")
    try:
        conn = sqlite3.connect(config.DB_PATH)
        df   = pd.read_sql_query(
            "SELECT sender, sender_email, date, is_read, has_attachments FROM emails",
            conn
        )
        conn.close()

        df["date"] = pd.to_datetime(df["date"])
        df["mois"] = df["date"].dt.to_period("M").astype(str)

        col1, col2, col3 = st.columns(3)
        col1.metric("Total emails",        len(df))
        col2.metric("Emails non lus",      len(df[df["is_read"] == 0]))
        col3.metric("Avec pièces jointes", len(df[df["has_attachments"] == 1]))

        st.markdown("---")

        emails_par_mois = df.groupby("mois").size().reset_index(name="count")
        fig1 = px.bar(
            emails_par_mois, x="mois", y="count",
            title="📅 Emails reçus par mois"
        )
        st.plotly_chart(fig1, use_container_width=True)

        top_senders = (
            df.groupby("sender_email")
            .size()
            .reset_index(name="count")
            .sort_values("count", ascending=False)
            .head(10)
        )
        fig2 = px.bar(
            top_senders, x="count", y="sender_email",
            orientation="h", title="👥 Top 10 expéditeurs"
        )
        st.plotly_chart(fig2, use_container_width=True)

    except Exception as e:
        st.error(f"Erreur : {e}")

# ── PAGE SYNCHRONISATION ─────────────────────────────────────
elif menu == "🔄 Synchronisation":
    st.title("🔄 Synchronisation des emails")

    st.info("""
    **Avant de synchroniser, vérifiez que les secrets sont configurés :**
    - AZURE_CLIENT_ID
    - AZURE_CLIENT_SECRET
    - AZURE_TENANT_ID
    - USER_EMAIL
    """)

    max_emails = st.slider("Nombre max d'emails à récupérer", 100, 5000, 1000, 100)

    if st.button("🚀 Lancer la synchronisation", type="primary"):
        try:
            from src.email_fetcher import EmailFetcher
            from src.indexer import EmailIndexer

            with st.spinner("📥 Récupération des emails..."):
                fetcher = EmailFetcher()
                total   = fetcher.fetch_all_emails(max_emails=max_emails)
                st.success(f"✅ {total} emails récupérés !")

            with st.spinner("🔍 Indexation en cours..."):
                indexer = EmailIndexer()
                indexed = indexer.index_all_emails()
                st.success(f"✅ {indexed} emails indexés !")

            st.balloons()

        except Exception as e:
            st.error(f"❌ Erreur : {e}")

# ── PAGE DEBUG ───────────────────────────────────────────────
elif menu == "🔧 Debug":
    st.title("🔧 Test connexion API")

    if st.button("Tester la connexion"):
        try:
            import requests
            from msal import ConfidentialClientApplication

            app = ConfidentialClientApplication(
                config.CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
                client_credential=config.CLIENT_SECRET
            )
            result = app.acquire_token_for_client(scopes=config.SCOPES)

            if "access_token" in result:
                st.success("✅ Token obtenu !")

                headers  = {"Authorization": f"Bearer {result['access_token']}"}
                url      = f"{config.GRAPH_ENDPOINT}/users/{config.USER_EMAIL}/messages?$top=5"
                response = requests.get(url, headers=headers)

                st.write("**Status HTTP :**", response.status_code)
                st.json(response.json())
            else:
                st.error(f"❌ Erreur token : {result.get('error_description')}")

        except Exception as e:
            st.error(f"❌ Exception : {e}")
