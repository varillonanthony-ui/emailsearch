import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import sys, os

sys.path.append(os.path.dirname(__file__))
import config

st.set_page_config(
    page_title="📧 Email Search",
    page_icon="📧",
    layout="wide"
)

# ── SIDEBAR ──────────────────────────────────────────────────
with st.sidebar:
    st.title("📧 Email Search")
    st.markdown("---")
    menu = st.radio("Navigation", [
        "🔍 Recherche",
        "📊 Statistiques",
        "🔄 Synchronisation"
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
            date_debut = st.date_input("Date début")
        with col4:
            date_fin   = st.date_input("Date fin")
        sender_filter = st.text_input("Expéditeur")

    if query:
        try:
            from src.search_engine import SearchEngine
            engine  = SearchEngine()
            results = engine.search(query, fields=search_in)

            st.markdown(f"**{len(results)} résultat(s) trouvé(s)**")
            st.markdown("---")

            for email in results:
                with st.expander(
                    f"📧 {email['subject']} | 👤 {email['sender']} | 📅 {email['date'][:10]}"
                ):
                    st.markdown(f"**De :** {email['sender']} ({email['sender_email']})")
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
        col1.metric("Total emails",      len(df))
        col2.metric("Emails non lus",    len(df[df["is_read"] == 0]))
        col3.metric("Avec pièces jointes", len(df[df["has_attachments"] == 1]))

        st.markdown("---")

        # Emails par mois
        emails_par_mois = df.groupby("mois").size().reset_index(name="count")
        fig1 = px.bar(
            emails_par_mois, x="mois", y="count",
            title="📅 Emails reçus par mois"
        )
        st.plotly_chart(fig1, use_container_width=True)

        # Top expéditeurs
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
                indexer  = EmailIndexer()
                indexed  = indexer.index_all_emails()
                st.success(f"✅ {indexed} emails indexés !")

            st.balloons()

        except Exception as e:
            st.error(f"❌ Erreur : {e}")
