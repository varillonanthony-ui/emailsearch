"""
app.py – Email Search multi-comptes Office 365.
Supporte N boîtes mail sur des tenants Microsoft différents.
Auth : mot de passe applicatif + Device Code Flow (endpoint multi-tenant).
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
.account-card {
    border:1px solid #e0e0e0; border-radius:8px;
    padding:10px 12px; margin:4px 0; background:#fafafa;
}
</style>""", unsafe_allow_html=True)


# ── Helpers globaux ───────────────────────────────────────────────────────────

def _client_id() -> str:
    return st.secrets["AZURE_CLIENT_ID"]

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


# ══════════════════════════════════════════════════════════════════════════════
# GESTION MULTI-COMPTES
# Structure session_state["accounts"] :
#   { user_id: { "access_token", "refresh_token", "user_info" } }
# ══════════════════════════════════════════════════════════════════════════════

def get_accounts() -> dict:
    return st.session_state.get("accounts", {})

def has_accounts() -> bool:
    return bool(get_accounts())

def get_valid_token(user_id: str) -> str | None:
    """Retourne un token valide pour un compte, le rafraîchit si besoin."""
    acc = get_accounts().get(user_id)
    if not acc:
        return None
    ref = acc.get("refresh_token")
    if ref:
        new = refresh_access_token(_client_id(), ref)
        if new:
            st.session_state["accounts"][user_id]["access_token"] = new
            return new
    return acc.get("access_token")

def add_account(access_token: str, refresh_token: str | None, user_info: dict):
    """Ajoute ou met à jour un compte dans la session."""
    user_id = user_info.get("id") or user_info.get("mail") or "unknown"
    if "accounts" not in st.session_state:
        st.session_state["accounts"] = {}
    st.session_state["accounts"][user_id] = {
        "access_token":  access_token,
        "refresh_token": refresh_token,
        "user_info":     user_info,
    }
    return user_id

def remove_account(user_id: str):
    """Déconnecte un compte."""
    st.session_state.get("accounts", {}).pop(user_id, None)
    # Nettoie les clés de session liées à ce compte
    for k in list(st.session_state.keys()):
        if user_id in k:
            del st.session_state[k]


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
    """Page de connexion du premier compte."""
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## 📧 Email Search")
        st.markdown("---")
        st.markdown("### Connexion Microsoft Office 365")
        st.caption("Vous pouvez connecter plusieurs comptes de domaines différents.")
        st.markdown("<br>", unsafe_allow_html=True)
        _device_code_panel("first")


def _device_code_panel(panel_key: str):
    """
    Panneau générique de device code flow.
    panel_key : identifiant unique pour ne pas mélanger les états entre panneaux.
    """
    client_id   = _client_id()
    flow_key    = f"device_flow_{panel_key}"

    if flow_key not in st.session_state:
        if st.button("🔐 Se connecter avec Microsoft",
                     type="primary", use_container_width=True,
                     key=f"btn_connect_{panel_key}"):
            with st.spinner("Initialisation…"):
                try:
                    st.session_state[flow_key] = start_device_flow(client_id)
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
        st.caption("🔒 Compatible avec tous les domaines Microsoft 365.")
        return

    flow      = st.session_state[flow_key]
    user_code = flow.get("user_code", "")
    verify    = flow.get("verification_uri", "https://microsoft.com/devicelogin")

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
    st.markdown("**3.** Connectez-vous, puis revenez ici et cliquez le bouton.")
    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ J'ai validé le code", type="primary",
                     use_container_width=True, key=f"btn_ok_{panel_key}"):
            with st.spinner("Vérification…"):
                for _ in range(20):
                    tok, ref = poll_token(client_id, flow["device_code"])
                    if tok:
                        ui = get_user_info(tok)
                        if ui:
                            add_account(tok, ref, ui)
                            st.session_state.pop(flow_key, None)
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Impossible de récupérer le profil.")
                        break
                    time.sleep(3)
                else:
                    st.warning("Code non validé ou expiré. Réessayez.")
    with c2:
        if st.button("↩ Recommencer", use_container_width=True,
                     key=f"btn_restart_{panel_key}"):
            st.session_state.pop(flow_key, None)
            st.rerun()


# ══════════════════════════════════════════
# SYNCHRONISATION
# ══════════════════════════════════════════

def run_sync(user_id: str, force_full: bool = False):
    token = get_valid_token(user_id)
    if not token:
        st.error("Token invalide. Reconnectez ce compte.")
        return

    user_info = get_accounts()[user_id]["user_info"]
    label     = user_info.get("mail") or user_info.get("displayName", user_id)
    mode      = "complète" if force_full else "incrémentale"
    st.markdown(f"### 🔄 Sync {mode} — {label}")

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
        new_ph.metric("🆕 Nouveaux",   f"{r.emails_new:,}")
        upd_ph.metric("✏️ Mis à jour", f"{r.emails_updated:,}")
        done = r.folders_done + r.folders_skip
        fold_ph.metric("📁 Dossiers",  f"{done}/{r.total_folders}")
        warn_ph.metric("⚠️ Alertes",   str(len(r.warnings)))
        if r.total_folders:
            bar.progress(min(done / r.total_folders, 1.0))
        log.append(msg)
        log_ph.code("\n".join(log[-8:]))

    try:
        result = EmailIndexer(token, user_id).sync(
            force_full=force_full, on_status=on_status)
        bar.progress(1.0)
        log_ph.empty()

        parts = []
        if result.emails_new:     parts.append(f"**{result.emails_new:,}** nouveaux")
        if result.emails_updated: parts.append(f"**{result.emails_updated:,}** mis à jour")
        st.success(f"✅ {label} — {', '.join(parts) or 'aucun changement'}")

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
        if result.errors:
            with st.expander(f"🔴 {len(result.errors)} erreur(s)"):
                for e in result.errors:
                    st.caption(e)

        st.session_state.pop("pending_sync", None)
        st.cache_data.clear()
        st.rerun()

    except PermissionError as e:
        st.error(f"🔑 {e}")
        remove_account(user_id)
    except Exception as e:
        st.error(f"❌ {e}")
        st.caption("Les dossiers déjà terminés sont sauvegardés.")
        st.session_state.pop("pending_sync", None)


# ══════════════════════════════════════════
# RECHERCHE
# ══════════════════════════════════════════

@st.cache_data(ttl=120, show_spinner=False)
def _cached_search(db_path: str, keywords: tuple, folder_filter: str,
                   folder_ids_key: str, date_from, date_to, offset: int):
    from database import Database as _DB
    import pathlib
    db = _DB.__new__(_DB)
    db.db_path = pathlib.Path(db_path)
    db._conn   = None
    db._init_db()
    folder_ids = list(folder_ids_key.split(",")) if folder_ids_key else None
    return db.search_emails(
        keywords=list(keywords), folder_filter=folder_filter,
        folder_ids=folder_ids, date_from=date_from, date_to=date_to,
        limit=PAGE_SIZE, offset=offset)

@st.cache_data(ttl=60, show_spinner=False)
def _get_stats(db_path: str) -> dict:
    from database import Database as _DB
    import pathlib
    db = _DB.__new__(_DB)
    db.db_path = pathlib.Path(db_path)
    db._conn   = None
    db._init_db()
    return db.get_stats()

@st.cache_data(ttl=60, show_spinner=False)
def _get_folders(db_path: str) -> list:
    from database import Database as _DB
    import pathlib
    db = _DB.__new__(_DB)
    db.db_path = pathlib.Path(db_path)
    db._conn   = None
    db._init_db()
    return db.get_folders()


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
        st.warning("Aucun résultat — tous les mots-clés doivent être présents.")
        return

    total_pages = max(1, (total - 1) // PAGE_SIZE + 1)
    st.markdown(f"**{total:,} résultat(s)** | page **{page+1}/{total_pages}**")
    st.markdown("---")

    for email in results:
        badges = ""
        if not email["is_read"]:
            badges += '<span class="badge badge-unread">Non lu</span>'
        if email["importance"] == "high":
            badges += '<span class="badge badge-high">🔴 Urgent</span>'

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
                subj_hl = highlight(email["subject"] or "(Sans objet)", keywords)
                st.markdown(
                    f"{badges}<br><b>Objet :</b> {subj_hl}<br>"
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

            preview_hl = highlight(email["body_preview"] or "", keywords)
            st.markdown(f"<div style='color:#444;line-height:1.6;margin-top:8px'>"
                        f"{preview_hl}</div>", unsafe_allow_html=True)

            if st.button("📄 Contenu complet", key=f"full_{tab_key}_{email['id']}"):
                full = db.get_email_detail(email["id"])
                body = (full.get("body") or "") if full else ""
                if not body:
                    token = get_valid_token(user_id) if user_id else None
                    if token:
                        with st.spinner("Chargement…"):
                            body = EmailIndexer(token, user_id).get_email_body(email["id"])
                            if body and full:
                                full["body"] = body
                                db.upsert_email(dict(full))
                if body:
                    safe_html = (
                        "<!DOCTYPE html><html lang='fr'><head><meta charset='UTF-8'>"
                        "<base target='_blank'><style>"
                        "*{box-sizing:border-box}"
                        "body{margin:12px 16px;font-family:-apple-system,BlinkMacSystemFont,"
                        "'Segoe UI',sans-serif;font-size:14px;line-height:1.6;color:#1a1a1a}"
                        "img{max-width:100%;height:auto}a{color:#0078d4}"
                        "blockquote{border-left:3px solid #ccc;margin:4px 0 4px 16px;"
                        "padding-left:12px;color:#555}"
                        "table{border-collapse:collapse;max-width:100%}"
                        "td,th{padding:4px 8px;border:1px solid #ddd}"
                        "</style></head><body>" + body + "</body></html>"
                    )
                    components.html(safe_html, height=600, scrolling=True)
                else:
                    st.info("Corps non disponible. Ouvrez dans Outlook.")

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
# PAGE PRINCIPALE MULTI-COMPTES
# ══════════════════════════════════════════

def page_main():
    accounts = get_accounts()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 📧 Email Search")
        st.markdown("---")

        # ── Comptes connectés ─────────────────────────────────────────────────
        st.markdown("### 👥 Comptes connectés")
        for uid, acc in list(accounts.items()):
            ui    = acc["user_info"]
            name  = ui.get("displayName", "Utilisateur")
            mail  = ui.get("mail") or ui.get("userPrincipalName", "")
            db    = Database(uid)
            stats = _get_stats(str(db.db_path))

            with st.expander(f"**{name}**  \n{mail}", expanded=False):
                c1, c2 = st.columns(2)
                c1.metric("Emails",   f"{stats['total_emails']:,}")
                c2.metric("Dossiers", stats["total_folders"])
                if stats.get("last_sync"):
                    try:
                        dt = datetime.fromisoformat(stats["last_sync"])
                        st.caption(f"🔄 {dt.strftime('%d/%m/%Y %H:%M')}")
                    except Exception:
                        pass

                if st.button("⚡ Sync incrémentale", key=f"sync_inc_{uid}",
                             use_container_width=True, type="primary"):
                    st.session_state["pending_sync"] = {"uid": uid, "full": False}
                    st.rerun()
                if st.button("🔁 Sync complète", key=f"sync_full_{uid}",
                             use_container_width=True):
                    st.session_state["pending_sync"] = {"uid": uid, "full": True}
                    st.rerun()

                with st.expander("🗑️ Réinitialiser"):
                    if st.button("Vider la base", key=f"reset_{uid}",
                                 use_container_width=True, type="secondary"):
                        db.reset_all()
                        st.cache_data.clear()
                        st.success("Base vidée.")
                        st.rerun()

                if st.button("🚪 Déconnecter ce compte", key=f"logout_{uid}",
                             use_container_width=True):
                    remove_account(uid)
                    st.cache_data.clear()
                    st.rerun()

        st.markdown("---")

        # ── Ajouter un compte ─────────────────────────────────────────────────
        with st.expander("➕ Ajouter un compte"):
            st.caption("Connectez une autre boîte mail (autre domaine Microsoft 365).")
            _device_code_panel("add")

        st.markdown("---")
        if st.button("🚪 Tout déconnecter", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.cache_data.clear()
            st.rerun()

    # ── Sync en attente ───────────────────────────────────────────────────────
    pending = st.session_state.get("pending_sync")
    if pending:
        run_sync(pending["uid"], force_full=pending["full"])
        return

    # ── Sélecteur de compte(s) à chercher ─────────────────────────────────────
    st.markdown("# 📧 Recherche d'emails")

    account_options = {
        uid: (acc["user_info"].get("mail") or acc["user_info"].get("displayName", uid))
        for uid, acc in accounts.items()
    }
    # Option "Tous les comptes"
    ALL = "__all__"
    options_display = {ALL: "🔍 Tous les comptes", **account_options}
    selected_uid = st.selectbox(
        "Chercher dans :",
        options=list(options_display.keys()),
        format_func=lambda k: options_display[k],
    )

    # ── Formulaire de recherche ───────────────────────────────────────────────
    with st.form("search_form"):
        kw_input = st.text_input(
            "🔍 Mots-clés",
            placeholder="réunion, budget  ou  réunion; budget  ou  réunion  budget",
            help="Séparez par virgule, point-virgule ou double-espace. Logique ET.",
        )
        d1, d2 = st.columns(2)
        date_from = d1.date_input("📅 Du", value=None, format="DD/MM/YYYY")
        date_to   = d2.date_input("📅 Au", value=None, format="DD/MM/YYYY")
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
                "search_kw":        keywords,
                "search_date_from": df_iso,
                "search_date_to":   dt_iso,
            })
            for k in list(st.session_state.keys()):
                if k.startswith("page_"):
                    st.session_state[k] = 0

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

    # ── Comptes cibles ────────────────────────────────────────────────────────
    target_uids = list(accounts.keys()) if selected_uid == ALL else [selected_uid]

    # ── Onglets dossier-filtre ────────────────────────────────────────────────
    tab_all, tab_inbox, tab_folder = st.tabs([
        "📬 Toute la boîte",
        "📥 Hors Envoyés / Supprimés",
        "📁 Dossier spécifique",
    ])

    def _render_for_accounts(folder_filter, tab_suffix, folder_ids_per_uid=None):
        """Affiche les résultats pour un ou plusieurs comptes."""
        if not kw_active:
            st.info("Entrez des mots-clés pour lancer une recherche.")
            return
        for uid in target_uids:
            db = Database(uid)
            mail = accounts[uid]["user_info"].get("mail", uid)
            if len(target_uids) > 1:
                st.markdown(f"#### 📬 {mail}")
            fids = (folder_ids_per_uid or {}).get(uid)
            show_results(db, kw_active, folder_filter, fids,
                         f"{tab_suffix}_{uid}", df_active, dt_active, uid)

    with tab_all:
        _render_for_accounts("all", "all")

    with tab_inbox:
        _render_for_accounts("no_sent_deleted", "inbox")

    with tab_folder:
        if not kw_active:
            st.info("Entrez des mots-clés pour lancer une recherche.")
        else:
            folder_ids_per_uid: dict[str, list[str]] = {}
            for uid in target_uids:
                db      = Database(uid)
                folders = _get_folders(str(db.db_path))
                mail    = accounts[uid]["user_info"].get("mail", uid)
                if not folders:
                    st.info(f"Aucun dossier pour {mail}. Lancez une synchronisation.")
                    continue
                folder_map = {f["display_path"]: f["id"]
                              for f in sorted(folders, key=lambda x: x["display_path"])}
                label    = f"Dossier — {mail}" if len(target_uids) > 1 else "Dossier"
                selected = st.selectbox(label, options=list(folder_map.keys()),
                                        key=f"folder_sel_{uid}")
                fid = folder_map.get(selected)
                if fid:
                    folder_ids_per_uid[uid] = [fid]

            if folder_ids_per_uid:
                for uid, fids in folder_ids_per_uid.items():
                    db   = Database(uid)
                    mail = accounts[uid]["user_info"].get("mail", uid)
                    if len(target_uids) > 1:
                        st.markdown(f"#### 📬 {mail}")
                    show_results(db, kw_active, "specific", fids,
                                 f"specific_{uid}", df_active, dt_active, uid)


# ── Point d'entrée ────────────────────────────────────────────────────────────

def main():
    if not check_password():
        st.stop()
    if has_accounts():
        page_main()
    else:
        page_login()

if __name__ == "__main__":
    main()