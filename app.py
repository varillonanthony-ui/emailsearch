"""
app.py – Email Search sur Office 365.
Auth : mot de passe applicatif + Device Code Flow Microsoft.
"""

import re
import html
import time
import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime

from auth import start_device_flow, poll_token, refresh_access_token, get_user_info
from database import Database
from email_indexer import EmailIndexer, SyncResult

PAGE_SIZE = 25

st.set_page_config(page_title="📧 Email Search", page_icon="📧",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
.badge { display:inline-block; padding:2px 9px; border-radius:20px;
         font-size:.70em; font-weight:600; margin-right:3px; }
.badge-unread { background:#e3f2fd; color:#1565c0; }
.badge-high   { background:#ffebee; color:#b71c1c; }
mark { background:#fff176; border-radius:2px; padding:0 1px; }
</style>""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _creds():
    return st.secrets["AZURE_TENANT_ID"], st.secrets["AZURE_CLIENT_ID"]

def is_logged_in():
    return "user_info" in st.session_state and "access_token" in st.session_state

def get_access_token() -> str | None:
    if not is_logged_in():
        return None
    ref = st.session_state.get("refresh_token")
    if ref:
        t, c = _creds()
        new = refresh_access_token(t, c, ref)
        if new:
            st.session_state["access_token"] = new
            return new
    return st.session_state.get("access_token")

def fmt_date(s: str) -> str:
    if not s:
        return "—"
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return s

def _fr(iso: str) -> str:
    try:
        y, m, d = iso[:10].split("-")
        return f"{d}/{m}/{y}"
    except Exception:
        return iso

def highlight(text: str, keywords: list[str]) -> str:
    safe = html.escape(text)
    for kw in keywords:
        if kw:
            safe = re.compile(re.escape(kw), re.IGNORECASE).sub(
                lambda m: f"<mark>{m.group()}</mark>", safe)
    return safe

def parse_keywords(raw: str) -> list[str]:
    return [k.strip() for k in re.split(r"[,;]+|\s{2,}", raw.strip()) if k.strip()]


# ══════════════════════════════════════════
# AUTHENTIFICATION
# ══════════════════════════════════════════

def check_password() -> bool:
    if st.session_state.get("app_authenticated"):
        return True
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## 📧 Email Search")
        st.markdown("---")
        st.markdown("### 🔒 Accès sécurisé")
        pwd = st.text_input("Mot de passe", type="password",
                            placeholder="Entrez le mot de passe…",
                            label_visibility="collapsed")
        if st.button("Continuer →", type="primary", use_container_width=True):
            if pwd == st.secrets.get("APP_PASSWORD", ""):
                st.session_state["app_authenticated"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
    return False


def page_login():
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## 📧 Email Search")
        st.markdown("---")
        st.markdown("### Connexion Microsoft Office 365")
        st.markdown("<br>", unsafe_allow_html=True)
        tenant_id, client_id = _creds()

        if "device_flow" not in st.session_state:
            if st.button("🔐 Se connecter avec Microsoft",
                         type="primary", use_container_width=True):
                with st.spinner("Initialisation…"):
                    try:
                        st.session_state["device_flow"] = start_device_flow(tenant_id, client_id)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")
            st.caption("🔒 Accès réservé au tenant Microsoft configuré.")
            return

        flow      = st.session_state["device_flow"]
        user_code = flow.get("user_code", "")
        verify    = flow.get("verification_uri", "https://microsoft.com/devicelogin")

        st.markdown("**Suivez ces 3 étapes :**")
        st.markdown(
            f"**1.** Ouvrez : <a href='{verify}' target='_blank' style='"
            "background:#0078d4;color:#fff;padding:5px 12px;border-radius:5px;"
            f"text-decoration:none;font-weight:600'>🌐 {verify}</a>",
            unsafe_allow_html=True)
        st.markdown(
            f"<div style='font-size:2.2em;font-weight:800;letter-spacing:8px;"
            "background:#f0f4ff;border:2px solid #0078d4;border-radius:8px;"
            f"padding:14px 0;text-align:center;color:#0078d4;margin:12px 0'>{user_code}</div>",
            unsafe_allow_html=True)
        st.markdown("**3.** Connectez-vous avec votre compte Office 365, puis revenez ici.")
        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ J'ai validé le code", type="primary", use_container_width=True):
                with st.spinner("Vérification…"):
                    for _ in range(20):
                        tok, ref = poll_token(tenant_id, client_id, flow["device_code"])
                        if tok:
                            ui = get_user_info(tok)
                            if ui:
                                st.session_state.update({
                                    "access_token": tok, "refresh_token": ref, "user_info": ui
                                })
                                st.session_state.pop("device_flow", None)
                                st.rerun()
                            else:
                                st.error("Impossible de récupérer le profil.")
                            break
                        time.sleep(3)
                    else:
                        st.warning("Code non validé ou expiré. Réessayez.")
        with c2:
            if st.button("↩ Recommencer", use_container_width=True):
                st.session_state.pop("device_flow", None)
                st.rerun()


# ══════════════════════════════════════════
# SYNCHRONISATION
# ══════════════════════════════════════════

def run_sync(access_token: str, user_id: str, force_full: bool = False):
    mode = "complète" if force_full else "incrémentale"
    st.markdown(f"### 🔄 Synchronisation {mode} en cours…")

    status_ph = st.empty()
    c1, c2, c3, c4 = st.columns(4)
    new_ph  = c1.empty()
    upd_ph  = c2.empty()
    fold_ph = c3.empty()
    warn_ph = c4.empty()
    bar     = st.progress(0.0)
    log_ph  = st.empty()
    log: list[str] = []

    def on_status(msg: str, r: SyncResult):
        status_ph.info(f"📬 {msg}")
        new_ph.metric("🆕 Nouveaux",  f"{r.emails_new:,}")
        upd_ph.metric("✏️ Mis à jour", f"{r.emails_updated:,}")
        done = r.folders_done + r.folders_skip
        fold_ph.metric("📁 Dossiers",  f"{done}/{r.total_folders}")
        warn_ph.metric("⚠️ Alertes",   str(len(r.warnings)))
        if r.total_folders:
            bar.progress(min(done / r.total_folders, 1.0))
        log.append(msg)
        log_ph.code("\n".join(log[-8:]))

    try:
        result = EmailIndexer(access_token, user_id).sync(
            force_full=force_full, on_status=on_status)
        bar.progress(1.0)
        log_ph.empty()

        parts = []
        if result.emails_new:     parts.append(f"**{result.emails_new:,}** nouveaux")
        if result.emails_updated: parts.append(f"**{result.emails_updated:,}** mis à jour")
        st.success(f"✅ Terminée — {', '.join(parts) or 'aucun changement'}")

        if result.folder_details:
            with st.expander(f"📊 Détail ({len(result.folder_details)} dossiers)",
                             expanded=bool(result.warnings)):
                for info in result.folder_details:
                    gap = info.expected - info.indexed
                    st.markdown(
                        f"{'✅' if gap <= 10 else '⚠️'} **{info.name}** — "
                        f"{info.indexed:,}/{info.expected:,} (+{info.new})")

        if result.warnings:
            st.warning("Certains dossiers semblent incomplets. Relancez une sync complète.")
            with st.expander(f"⚠️ {len(result.warnings)} avertissement(s)"):
                for w in result.warnings:
                    st.caption(w)

        if result.errors:
            with st.expander(f"🔴 {len(result.errors)} erreur(s)"):
                for e in result.errors:
                    st.caption(e)

        st.session_state.pop("show_sync", None)
        st.session_state.pop("sync_force_full", None)
        # Invalide le cache stats/folders
        st.cache_data.clear()
        st.rerun()

    except PermissionError as e:
        st.error(f"🔑 {e}")
        for k in ["user_info", "access_token", "refresh_token"]:
            st.session_state.pop(k, None)
    except Exception as e:
        st.error(f"❌ {e}")
        st.caption("Les dossiers déjà terminés sont sauvegardés. Relancez pour reprendre.")
        st.session_state.pop("show_sync", None)
        st.session_state.pop("sync_force_full", None)


# ══════════════════════════════════════════
# RECHERCHE
# ══════════════════════════════════════════

@st.cache_data(ttl=120, show_spinner=False)
def _cached_search(db_path: str, keywords: tuple, folder_filter: str,
                   folder_ids_key: str, date_from, date_to, offset: int):
    """Cache des résultats de recherche (2 min) pour fluidifier la navigation."""
    from database import Database as _DB
    db = _DB.__new__(_DB)
    db.db_path = __import__("pathlib").Path(db_path)
    db._conn = None
    db._init_db()
    folder_ids = list(folder_ids_key.split(",")) if folder_ids_key else None
    return db.search_emails(
        keywords=list(keywords),
        folder_filter=folder_filter,
        folder_ids=folder_ids,
        date_from=date_from,
        date_to=date_to,
        limit=PAGE_SIZE,
        offset=offset,
    )


def show_results(db, keywords, folder_filter, folder_ids, tab_key,
                 date_from=None, date_to=None, user_id=None):
    page_key = f"page_{tab_key}"
    if page_key not in st.session_state:
        st.session_state[page_key] = 0
    page = st.session_state[page_key]

    folder_ids_key = ",".join(folder_ids) if folder_ids else ""
    with st.spinner("Recherche…"):
        results, total = _cached_search(
            str(db.db_path), tuple(keywords), folder_filter,
            folder_ids_key, date_from, date_to, page * PAGE_SIZE)

    if total == 0:
        st.warning(f"Aucun résultat — tous les mots-clés doivent être présents.")
        return

    total_pages = max(1, (total - 1) // PAGE_SIZE + 1)
    kw_str = " ET ".join(f"`{k}`" for k in keywords)
    st.markdown(f"**{total:,} résultat(s)** — {kw_str} &nbsp;|&nbsp; "
                f"page **{page+1}/{total_pages}**", unsafe_allow_html=True)
    st.markdown("---")

    for email in results:
        badges = ""
        if not email["is_read"]:
            badges += '<span class="badge badge-unread">Non lu</span>'
        if email["importance"] == "high":
            badges += '<span class="badge badge-high">🔴 Urgent</span>'

        subj_hl    = highlight(email["subject"] or "(Sans objet)", keywords)
        preview_hl = highlight(email["body_preview"] or "", keywords)

        title = (
            f"{'🔵 ' if not email['is_read'] else '⚪ '}"
            f"{'📎 ' if email['has_attachments'] else ''}"
            f"{email['subject'] or '(Sans objet)'} "
            f"— {email['sender_name'] or email['sender_email']} "
            f"— {fmt_date(email['received_datetime'])}"
        )
        with st.expander(title):
            c1, c2 = st.columns([4, 1])
            with c1:
                st.markdown(
                    f"{badges}<br>"
                    f"<b>De :</b> {html.escape(email['sender_name'] or '')} "
                    f"&lt;{email['sender_email']}&gt;<br>"
                    f"<b>À :</b> {html.escape((email['recipients'] or '—')[:300])}<br>"
                    f"<b>Dossier :</b> 📁 {email['folder_name']} &nbsp; "
                    f"<b>Date :</b> {fmt_date(email['received_datetime'])}",
                    unsafe_allow_html=True)
            with c2:
                if email.get("web_link"):
                    st.link_button("📬 Ouvrir", email["web_link"],
                                   use_container_width=True)

            st.markdown(f"<div style='color:#444;line-height:1.6;margin-top:8px'>"
                        f"{preview_hl}</div>", unsafe_allow_html=True)

            if st.button("📄 Contenu complet", key=f"full_{tab_key}_{email['id']}"):
                full = db.get_email_detail(email["id"])
                body = (full.get("body") or "") if full else ""
                if not body:
                    token = get_access_token()
                    uid = user_id or st.session_state.get("user_info", {}).get("id", "x")
                    if token:
                        with st.spinner("Chargement…"):
                            body = EmailIndexer(token, uid).get_email_body(email["id"])
                            if body and full:
                                full["body"] = body
                                db.upsert_email(dict(full))
                if body:
                    safe_html = (
                        "<!DOCTYPE html><html lang='fr'><head>"
                        "<meta charset='UTF-8'><base target='_blank'>"
                        "<style>*{box-sizing:border-box}"
                        "body{margin:12px 16px;font-family:-apple-system,BlinkMacSystemFont,"
                        "'Segoe UI',sans-serif;font-size:14px;line-height:1.6;color:#1a1a1a}"
                        "img{max-width:100%;height:auto}"
                        "a{color:#0078d4}"
                        "blockquote{border-left:3px solid #ccc;margin:4px 0 4px 16px;padding-left:12px;color:#555}"
                        "table{border-collapse:collapse;max-width:100%}"
                        "td,th{padding:4px 8px;border:1px solid #ddd}"
                        "</style></head><body>" + body + "</body></html>"
                    )
                    components.html(safe_html, height=600, scrolling=True)
                else:
                    st.info("Corps non disponible. Ouvrez dans Outlook.")

    # Pagination
    if total_pages > 1:
        st.markdown("---")
        p1, p2, p3 = st.columns([1, 4, 1])
        with p1:
            if page > 0 and st.button("← Précédent", key=f"prev_{tab_key}"):
                st.session_state[page_key] -= 1
                st.rerun()
        with p2:
            st.markdown(f"<p style='text-align:center'>Page <b>{page+1}</b> / {total_pages}</p>",
                        unsafe_allow_html=True)
        with p3:
            if page < total_pages - 1 and st.button("Suivant →", key=f"next_{tab_key}"):
                st.session_state[page_key] += 1
                st.rerun()


# ══════════════════════════════════════════
# PAGE PRINCIPALE
# ══════════════════════════════════════════

@st.cache_data(ttl=60, show_spinner=False)
def _get_stats(db_path: str) -> dict:
    from database import Database as _DB
    db = _DB.__new__(_DB)
    db.db_path = __import__("pathlib").Path(db_path)
    db._conn = None
    db._init_db()
    return db.get_stats()

@st.cache_data(ttl=60, show_spinner=False)
def _get_folders(db_path: str) -> list:
    from database import Database as _DB
    db = _DB.__new__(_DB)
    db.db_path = __import__("pathlib").Path(db_path)
    db._conn = None
    db._init_db()
    return db.get_folders()


def page_main():
    user_info = st.session_state["user_info"]
    user_id   = user_info.get("id") or user_info.get("mail") or "unknown"
    db        = Database(user_id)
    db_path   = str(db.db_path)
    stats     = _get_stats(db_path)
    folders   = _get_folders(db_path)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### 👤 {user_info.get('displayName', 'Utilisateur')}")
        st.caption(user_info.get("mail") or user_info.get("userPrincipalName", ""))
        st.markdown("---")

        c1, c2 = st.columns(2)
        c1.metric("Emails",   f"{stats['total_emails']:,}")
        c2.metric("Dossiers", stats["total_folders"])
        if stats.get("last_sync"):
            try:
                dt = datetime.fromisoformat(stats["last_sync"])
                st.caption(f"🔄 Dernière sync : {dt.strftime('%d/%m/%Y %H:%M')}")
            except Exception:
                pass

        st.markdown("---")
        if st.button("⚡ Sync incrémentale", use_container_width=True, type="primary",
                     help="Traite uniquement les nouveaux emails depuis la dernière sync."):
            st.session_state.update({"show_sync": True, "sync_force_full": False})
            st.rerun()
        if st.button("🔁 Sync complète", use_container_width=True,
                     help="Re-télécharge toute la boîte. À utiliser si des emails manquent."):
            st.session_state.update({"show_sync": True, "sync_force_full": True})
            st.rerun()

        with st.expander("📊 Par dossier"):
            for row in stats["by_folder"][:15]:
                st.caption(f"**{row['folder_name']}** : {row['cnt']:,}")

        st.markdown("---")
        with st.expander("🗑️ Réinitialiser la base"):
            st.caption("Supprime TOUS les emails indexés.")
            if st.button("Vider la base", use_container_width=True, type="secondary"):
                db.reset_all()
                st.cache_data.clear()
                for k in ["show_sync", "sync_force_full", "search_kw",
                          "search_date_from", "search_date_to"]:
                    st.session_state.pop(k, None)
                st.success("Base vidée.")
                st.rerun()

        if st.button("🚪 Déconnexion", use_container_width=True):
            for k in ["user_info", "access_token", "refresh_token",
                      "app_authenticated", "device_flow", "show_sync",
                      "search_kw", "search_date_from", "search_date_to"]:
                st.session_state.pop(k, None)
            st.cache_data.clear()
            st.rerun()

    # ── Sync ──────────────────────────────────────────────────────────────────
    if st.session_state.get("show_sync"):
        token = get_access_token()
        if token:
            run_sync(token, user_id, st.session_state.get("sync_force_full", False))
        else:
            st.error("Session expirée, reconnectez-vous.")
            st.session_state.pop("show_sync", None)
        return

    # ── Formulaire de recherche ───────────────────────────────────────────────
    st.markdown("# 📧 Recherche d'emails")

    with st.form("search_form"):
        kw_input = st.text_input(
            "🔍 Mots-clés",
            placeholder="réunion, budget  ou  réunion; budget  ou  réunion  budget",
            help="Séparez par virgule, point-virgule ou double-espace. Logique ET.",
        )
        d1, d2 = st.columns(2)
        date_from = d1.date_input("📅 Du",  value=None, format="DD/MM/YYYY")
        date_to   = d2.date_input("📅 Au",  value=None, format="DD/MM/YYYY")
        submitted = st.form_submit_button("Rechercher", type="primary",
                                          use_container_width=True)

    keywords = parse_keywords(kw_input) if kw_input else []
    df_iso   = date_from.isoformat() if date_from else None
    dt_iso   = date_to.isoformat()   if date_to   else None

    if submitted:
        if not keywords:
            st.warning("Veuillez entrer au moins un mot-clé.")
        else:
            st.session_state.update({
                "search_kw": keywords,
                "search_date_from": df_iso,
                "search_date_to":   dt_iso,
            })
            # Remet la pagination à 0 pour une nouvelle recherche
            for key in ["page_all", "page_no_sent", "page_specific"]:
                st.session_state[key] = 0

    kw_active = st.session_state.get("search_kw", [])
    df_active = st.session_state.get("search_date_from")
    dt_active = st.session_state.get("search_date_to")

    if kw_active:
        tags     = "  ".join(f"`{k}`" for k in kw_active)
        date_str = ""
        if df_active and dt_active:
            date_str = f" | 📅 {_fr(df_active)} → {_fr(dt_active)}"
        elif df_active:
            date_str = f" | 📅 depuis {_fr(df_active)}"
        elif dt_active:
            date_str = f" | 📅 jusqu'au {_fr(dt_active)}"
        st.caption(f"🔍 {tags}{date_str} — logique ET")

    # ── Onglets ───────────────────────────────────────────────────────────────
    tab_all, tab_inbox, tab_folder = st.tabs([
        "📬 Toute la boîte",
        "📥 Hors Envoyés / Supprimés",
        "📁 Dossier spécifique",
    ])

    with tab_all:
        if kw_active:
            show_results(db, kw_active, "all", None, "all",
                         df_active, dt_active, user_id)
        else:
            st.info("Entrez des mots-clés pour lancer une recherche.")

    with tab_inbox:
        if kw_active:
            show_results(db, kw_active, "no_sent_deleted", None, "no_sent",
                         df_active, dt_active, user_id)
        else:
            st.info("Entrez des mots-clés pour lancer une recherche.")

    with tab_folder:
        if not folders:
            st.info("Aucun dossier indexé. Lancez une synchronisation.")
        else:
            folder_map = {f["display_path"]: f["id"]
                          for f in sorted(folders, key=lambda x: x["display_path"])}
            selected  = st.selectbox("Dossier", options=list(folder_map.keys()))
            folder_id = folder_map.get(selected)
            if kw_active and folder_id:
                show_results(db, kw_active, "specific", [folder_id], "specific",
                             df_active, dt_active, user_id)
            elif not kw_active:
                st.info("Entrez des mots-clés pour lancer une recherche.")


# ── Point d'entrée ────────────────────────────────────────────────────────────

def main():
    if not check_password():
        st.stop()
    if is_logged_in():
        page_main()
    else:
        page_login()

if __name__ == "__main__":
    main()