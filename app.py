"""
app.py – Application Streamlit principale.
• Authentification Microsoft OAuth 2.0 (même tenant uniquement)
• Synchronisation complète + incrémentale via Graph API
• Recherche multi-mots-clés (logique ET) avec filtres par dossier
• Base de données isolée par utilisateur
"""

import streamlit as st
from datetime import datetime

from auth import get_auth_url, acquire_token_by_code, acquire_token_silent, get_user_info
from database import Database
from email_indexer import EmailIndexer

# ── Configuration page ────────────────────────────────────────────────────────

st.set_page_config(
    page_title="📧 Email Search",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* Cartes email */
.email-card {
    border: 1px solid #dde3ea;
    border-radius: 10px;
    padding: 14px 18px;
    margin: 6px 0;
    background: #ffffff;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
}
.email-card:hover { box-shadow: 0 3px 10px rgba(0,0,0,.12); }

/* Badges */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: .72em;
    font-weight: 600;
    margin-right: 4px;
    vertical-align: middle;
}
.badge-unread     { background:#e3f2fd; color:#1565c0; }
.badge-attachment { background:#e8f5e9; color:#2e7d32; }
.badge-high       { background:#ffebee; color:#b71c1c; }
.badge-low        { background:#f3f3f3; color:#757575; }

/* Bouton Microsoft */
.ms-btn {
    display:block; width:100%; padding:12px 0; text-align:center;
    background:#0078d4; color:#fff !important; border-radius:6px;
    font-size:1em; font-weight:600; text-decoration:none;
    transition:background .2s;
}
.ms-btn:hover { background:#106ebe; }
</style>
""", unsafe_allow_html=True)


# ── Helpers auth ──────────────────────────────────────────────────────────────

def is_logged_in() -> bool:
    return "user_info" in st.session_state and "token_cache" in st.session_state


def get_access_token() -> str | None:
    """Retourne un access token valide, rafraîchi si nécessaire."""
    if not is_logged_in():
        return None
    result, new_cache = acquire_token_silent(st.session_state["token_cache"])
    if result and "access_token" in result:
        st.session_state["token_cache"] = new_cache
        return result["access_token"]
    # Fallback : token stocké en session
    return st.session_state.get("access_token")


def handle_oauth_callback():
    """Détecte le code OAuth dans l'URL et l'échange contre un token."""
    params = st.query_params
    if "code" not in params or is_logged_in():
        return

    code = params["code"]
    try:
        result, cache = acquire_token_by_code(code)
    except Exception as e:
        st.error(f"Erreur lors de la connexion : {e}")
        st.query_params.clear()
        return

    if "access_token" not in result:
        st.error(f"Échec OAuth : {result.get('error_description', 'Erreur inconnue')}")
        st.query_params.clear()
        return

    user_info = get_user_info(result["access_token"])
    if not user_info:
        st.error("Impossible de récupérer les informations utilisateur.")
        st.query_params.clear()
        return

    st.session_state["token_cache"]  = cache
    st.session_state["access_token"] = result["access_token"]
    st.session_state["user_info"]    = user_info
    st.query_params.clear()
    st.rerun()


# ── Page de connexion ─────────────────────────────────────────────────────────

def page_login():
    col = st.columns([1, 2, 1])[1]
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## 📧 Email Search")
        st.markdown("---")
        st.markdown("### Connexion requise")
        st.markdown(
            "Connectez-vous avec votre compte **Microsoft Office 365** "
            "pour accéder à la recherche dans vos emails."
        )
        st.markdown("<br>", unsafe_allow_html=True)
        auth_url = get_auth_url()
        st.markdown(
            f'<a href="{auth_url}" target="_self" class="ms-btn">'
            f'🔐 &nbsp; Se connecter avec Microsoft</a>',
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)
        st.caption(
            "🔒 Seuls les comptes du tenant Microsoft configuré peuvent accéder à cette application. "
            "Chaque utilisateur dispose de sa propre base de données isolée."
        )


# ── Formatage ─────────────────────────────────────────────────────────────────

def fmt_date(dt_str: str) -> str:
    if not dt_str:
        return "—"
    try:
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return dt_str


def highlight_keywords(text: str, keywords: list[str]) -> str:
    """Met en gras les mots-clés dans un texte (HTML simple)."""
    import html
    safe = html.escape(text)
    for kw in keywords:
        if kw:
            import re
            pattern = re.compile(re.escape(kw), re.IGNORECASE)
            safe = pattern.sub(
                lambda m: f"<mark style='background:#fff176;border-radius:3px'>{m.group()}</mark>",
                safe,
            )
    return safe


# ── Synchronisation ───────────────────────────────────────────────────────────

def run_sync(access_token: str, user_id: str):
    """Lance la synchronisation et affiche la progression."""
    st.markdown("### 🔄 Synchronisation en cours…")
    status_ph = st.empty()
    counter_ph = st.empty()
    progress_bar = st.progress(0.0)

    indexer = EmailIndexer(access_token, user_id)

    def on_status(msg: str, count: int):
        status_ph.info(f"📬 {msg}")
        counter_ph.metric("Emails indexés", f"{count:,}")

    try:
        total = indexer.full_sync(on_status=on_status)
        progress_bar.progress(1.0)
        st.success(f"✅ Synchronisation terminée — **{total:,}** emails indexés / mis à jour.")
        st.session_state.pop("show_sync", None)
        st.rerun()
    except PermissionError as e:
        st.error(str(e))
        # Force re-login
        for k in ["user_info", "token_cache", "access_token"]:
            st.session_state.pop(k, None)
    except Exception as e:
        st.error(f"Erreur de synchronisation : {e}")
        st.session_state.pop("show_sync", None)


# ── Résultats de recherche ────────────────────────────────────────────────────

PAGE_SIZE = 25


def show_results(
    db: Database,
    keywords: list[str],
    folder_filter: str,
    folder_ids: list[str] | None,
    tab_key: str,
):
    """Affiche les résultats paginés avec mise en évidence des mots-clés."""
    page_key = f"page_{tab_key}"
    if page_key not in st.session_state:
        st.session_state[page_key] = 0

    page = st.session_state[page_key]

    if not keywords:
        st.info("ℹ️ Entrez au moins un mot-clé pour lancer la recherche.")
        return

    with st.spinner("Recherche en cours…"):
        results, total = db.search_emails(
            keywords=keywords,
            folder_filter=folder_filter,
            folder_ids=folder_ids,
            limit=PAGE_SIZE,
            offset=page * PAGE_SIZE,
        )

    if total == 0:
        st.warning(
            f"Aucun résultat pour : **{' + '.join(keywords)}**\n\n"
            "Tous les mots-clés doivent être présents dans l'email."
        )
        return

    total_pages = max(1, (total - 1) // PAGE_SIZE + 1)
    st.markdown(
        f"**{total:,} résultat(s)** — mots-clés : "
        + " &nbsp;`ET`&nbsp; ".join(f"`{k}`" for k in keywords)
        + f" &nbsp;|&nbsp; page **{page+1}** / {total_pages}",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    for email in results:
        badges = ""
        if not email["is_read"]:
            badges += '<span class="badge badge-unread">Non lu</span>'
        if email["has_attachments"]:
            badges += '<span class="badge badge-attachment">📎 PJ</span>'
        if email["importance"] == "high":
            badges += '<span class="badge badge-high">🔴 Urgent</span>'

        subject_hl = highlight_keywords(email["subject"] or "(Sans objet)", keywords)
        preview_hl = highlight_keywords(email["body_preview"] or "", keywords)

        with st.expander(
            f"{'🔵 ' if not email['is_read'] else '⚪ '}"
            f"{email['subject'] or '(Sans objet)'} "
            f"— {email['sender_name'] or email['sender_email']} "
            f"— {fmt_date(email['received_datetime'])}",
            expanded=False,
        ):
            c1, c2 = st.columns([4, 1])

            with c1:
                st.markdown(
                    f"{badges}<br>"
                    f"<b>Objet :</b> {subject_hl}<br>"
                    f"<b>De :</b> {email['sender_name']} &lt;{email['sender_email']}&gt;<br>"
                    f"<b>À :</b> {(email['recipients'] or '—')[:300]}<br>"
                    f"<b>Dossier :</b> 📁 {email['folder_name']}&nbsp;&nbsp;"
                    f"<b>Date :</b> {fmt_date(email['received_datetime'])}",
                    unsafe_allow_html=True,
                )

            with c2:
                if email.get("web_link"):
                    st.markdown(
                        f'<a href="{email["web_link"]}" target="_blank">'
                        f'<button style="background:#0078d4;color:#fff;border:none;'
                        f'padding:8px 14px;border-radius:6px;cursor:pointer;width:100%;font-size:.9em">'
                        f'📬 Ouvrir Outlook</button></a>',
                        unsafe_allow_html=True,
                    )

            st.markdown("---")
            st.markdown(
                f"<div style='color:#333;line-height:1.6'>{preview_hl}</div>",
                unsafe_allow_html=True,
            )

            if st.button("📄 Voir le contenu complet", key=f"full_{tab_key}_{email['id']}"):
                full = db.get_email_detail(email["id"])
                if full and full.get("body"):
                    st.text_area(
                        "Contenu complet",
                        full["body"],
                        height=350,
                        key=f"body_{tab_key}_{email['id']}",
                    )

    # ── Pagination ────────────────────────────────────────────────────────────
    if total_pages > 1:
        st.markdown("---")
        pc1, pc2, pc3 = st.columns([1, 4, 1])
        with pc1:
            if page > 0 and st.button("← Précédent", key=f"prev_{tab_key}"):
                st.session_state[page_key] -= 1
                st.rerun()
        with pc2:
            st.markdown(
                f"<p style='text-align:center;margin:8px 0'>Page <b>{page+1}</b> / {total_pages}</p>",
                unsafe_allow_html=True,
            )
        with pc3:
            if page < total_pages - 1 and st.button("Suivant →", key=f"next_{tab_key}"):
                st.session_state[page_key] += 1
                st.rerun()


# ── Page principale ───────────────────────────────────────────────────────────

def page_main():
    user_info = st.session_state["user_info"]
    user_id   = user_info.get("id") or user_info.get("mail") or "unknown"

    db      = Database(user_id)
    stats   = db.get_stats()
    folders = db.get_folders()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### 👤 {user_info.get('displayName', 'Utilisateur')}")
        st.caption(user_info.get("mail") or user_info.get("userPrincipalName", ""))
        st.markdown("---")

        col_a, col_b = st.columns(2)
        with col_a:
            st.metric("Emails", f"{stats['total_emails']:,}")
        with col_b:
            st.metric("Dossiers", stats["total_folders"])

        if stats.get("last_sync"):
            try:
                dt = datetime.fromisoformat(stats["last_sync"])
                st.caption(f"🔄 Dernière sync : {dt.strftime('%d/%m/%Y %H:%M')}")
            except Exception:
                pass

        st.markdown("---")

        if st.button("🔄 Synchroniser les emails", use_container_width=True, type="primary"):
            st.session_state["show_sync"] = True
            st.rerun()

        # Stats par dossier (sidebar déroulante)
        with st.expander("📊 Emails par dossier"):
            for row in stats["by_folder"][:15]:
                st.markdown(f"**{row['folder_name']}** : {row['cnt']:,}")

        st.markdown("---")
        if st.button("🚪 Se déconnecter", use_container_width=True):
            for k in ["user_info", "token_cache", "access_token", "show_sync"]:
                st.session_state.pop(k, None)
            st.rerun()

    # ── Sync panel ────────────────────────────────────────────────────────────
    if st.session_state.get("show_sync"):
        token = get_access_token()
        if token:
            run_sync(token, user_id)
        else:
            st.error("Session expirée, veuillez vous reconnecter.")
            st.session_state.pop("show_sync", None)
        return  # Ne pas afficher le reste pendant la sync

    # ── En-tête + barre de recherche ──────────────────────────────────────────
    st.markdown("# 📧 Recherche d'emails")

    with st.form("search_form"):
        kw_input = st.text_input(
            "🔍 Mots-clés (séparés par des virgules)",
            placeholder="Ex: réunion, budget, 2024",
            help="Logique ET : tous les mots-clés doivent être présents dans l'email.",
        )
        submitted = st.form_submit_button("Rechercher", type="primary", use_container_width=True)

    keywords: list[str] = []
    if kw_input:
        keywords = [k.strip() for k in kw_input.split(",") if k.strip()]

    if submitted and not keywords:
        st.warning("Veuillez entrer au moins un mot-clé.")

    # ── Onglets filtres dossiers ──────────────────────────────────────────────
    tab_all, tab_no_sent, tab_specific = st.tabs([
        "📬 Toute la boîte mail",
        "📥 Hors Envoyés / Supprimés",
        "📁 Dossier spécifique",
    ])

    with tab_all:
        if submitted:
            show_results(db, keywords, "all", None, "all")
        elif not submitted and "search_kw" in st.session_state:
            show_results(db, st.session_state["search_kw"], "all", None, "all")

    with tab_no_sent:
        if submitted:
            show_results(db, keywords, "no_sent_deleted", None, "no_sent")
        elif not submitted and "search_kw" in st.session_state:
            show_results(db, st.session_state["search_kw"], "no_sent_deleted", None, "no_sent")

    with tab_specific:
        if not folders:
            st.info("Aucun dossier indexé. Lancez une synchronisation d'abord.")
        else:
            folder_map = {
                f["display_path"]: f["id"]
                for f in sorted(folders, key=lambda x: x["display_path"])
            }
            selected = st.selectbox(
                "Choisir un dossier",
                options=list(folder_map.keys()),
                key="folder_select",
            )
            folder_id = folder_map.get(selected)

            if submitted and folder_id:
                show_results(db, keywords, "specific", [folder_id], "specific")
            elif not submitted and "search_kw" in st.session_state and folder_id:
                show_results(
                    db, st.session_state["search_kw"], "specific", [folder_id], "specific"
                )

    # Mémorise les derniers mots-clés pour conserver l'affichage lors des paginations
    if submitted and keywords:
        st.session_state["search_kw"] = keywords


# ── Point d'entrée ────────────────────────────────────────────────────────────

def main():
    handle_oauth_callback()

    if is_logged_in():
        page_main()
    else:
        page_login()


if __name__ == "__main__":
    main()