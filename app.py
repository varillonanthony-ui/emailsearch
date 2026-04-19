"""
app.py – Application Streamlit principale.
• Auth 2 étapes : mot de passe applicatif + Device Code Flow Microsoft
  (aucun redirect URI, fonctionne dans les iframes Streamlit Cloud)
• Synchronisation complète + incrémentale via Graph API
• Recherche multi-mots-clés (logique ET) avec filtres par dossier
• Base SQLite isolée par utilisateur
"""

import streamlit as st
import time
from datetime import datetime

from auth import start_device_flow, poll_token, refresh_access_token, get_user_info
from database import Database
from email_indexer import EmailIndexer

PAGE_SIZE = 25   # emails par page dans les résultats de recherche

st.set_page_config(page_title="📧 Email Search", page_icon="📧",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.badge { display:inline-block; padding:2px 10px; border-radius:20px;
         font-size:.72em; font-weight:600; margin-right:4px; vertical-align:middle; }
.badge-unread     { background:#e3f2fd; color:#1565c0; }
.badge-attachment { background:#e8f5e9; color:#2e7d32; }
.badge-high       { background:#ffebee; color:#b71c1c; }
</style>""", unsafe_allow_html=True)


# helpers secrets
def _creds():
    return st.secrets["AZURE_TENANT_ID"], st.secrets["AZURE_CLIENT_ID"]

# helpers auth
def is_logged_in():
    return "user_info" in st.session_state and "access_token" in st.session_state

def get_access_token():
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


# ══════════════════════════════════════════
# ÉTAPE 1 – Mot de passe applicatif
# ══════════════════════════════════════════

def check_password():
    if st.session_state.get("app_authenticated"):
        return True
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## 📧 Email Search")
        st.markdown("---")
        st.markdown("### 🔒 Accès sécurisé")
        st.markdown("Cette application est réservée aux membres de l'organisation.")
        st.markdown("<br>", unsafe_allow_html=True)
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


# ══════════════════════════════════════════
# ÉTAPE 2 – Device Code Flow Microsoft
# ══════════════════════════════════════════

def page_login():
    """Connexion via Device Code Flow — aucun redirect URI requis."""
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## 📧 Email Search")
        st.markdown("---")
        st.markdown("### Connexion Microsoft Office 365")
        st.markdown("<br>", unsafe_allow_html=True)
        tenant_id, client_id = _creds()

        if "device_flow" not in st.session_state:
            if st.button("🔐  Se connecter avec Microsoft",
                         type="primary", use_container_width=True):
                with st.spinner("Initialisation…"):
                    try:
                        st.session_state["device_flow"] = start_device_flow(tenant_id, client_id)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")
            st.markdown("<br>", unsafe_allow_html=True)
            st.caption("🔒 Accès réservé au tenant Microsoft configuré. "
                       "Chaque utilisateur dispose de sa propre base de données.")
            return

        flow      = st.session_state["device_flow"]
        user_code = flow.get("user_code", "")
        verify    = flow.get("verification_uri", "https://microsoft.com/devicelogin")

        st.markdown("**Suivez ces 3 étapes :**")
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**Étape 1 —** Ouvrez ce lien dans un onglet :")
        st.markdown(
            f"<a href='{verify}' target='_blank' style='"
            "display:inline-block;padding:8px 18px;background:#0078d4;color:#fff;"
            "border-radius:6px;text-decoration:none;font-weight:600'>"
            f"🌐 {verify}</a>",
            unsafe_allow_html=True)

        st.markdown("<br>**Étape 2 —** Entrez ce code :", unsafe_allow_html=True)
        st.markdown(
            f"<div style='font-size:2.4em;font-weight:800;letter-spacing:8px;"
            "background:#f0f4ff;border:2px solid #0078d4;border-radius:10px;"
            f"padding:16px 0;text-align:center;color:#0078d4;margin:8px 0'>{user_code}</div>",
            unsafe_allow_html=True)

        st.markdown("**Étape 3 —** Connectez-vous avec votre compte Office 365, puis revenez ici.")
        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅  J'ai validé le code", type="primary", use_container_width=True):
                with st.spinner("Vérification…"):
                    ok = False
                    for _ in range(20):
                        tok, ref = poll_token(tenant_id, client_id, flow["device_code"])
                        if tok:
                            ui = get_user_info(tok)
                            if ui:
                                st.session_state.update({
                                    "access_token": tok,
                                    "refresh_token": ref,
                                    "user_info": ui,
                                })
                                st.session_state.pop("device_flow", None)
                                ok = True
                                st.rerun()
                            else:
                                st.error("Impossible de récupérer le profil.")
                            break
                        time.sleep(3)
                    if not ok:
                        st.warning("Code pas encore validé ou expiré. "
                                   "Vérifiez la connexion Microsoft et réessayez.")
        with c2:
            if st.button("↩  Recommencer", use_container_width=True):
                st.session_state.pop("device_flow", None)
                st.rerun()


# ── Formatage ─────────────────────────────────────────────────────────────────

def fmt_date(s):
    if not s:
        return "—"
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return s

def highlight_keywords(text, keywords):
    import html as _html, re
    safe = _html.escape(text)
    for kw in keywords:
        if kw:
            safe = re.compile(re.escape(kw), re.IGNORECASE).sub(
                lambda m: "<mark style='background:#fff176;border-radius:3px;padding:0 1px'>"
                          + m.group() + "</mark>", safe)
    return safe


# ── Synchronisation ───────────────────────────────────────────────────────────

def run_sync(access_token: str, user_id: str, force_full: bool = False):
    from email_indexer import SyncResult

    mode = "complète" if force_full else "incrémentale"
    st.markdown(f"### 🔄 Synchronisation {mode} en cours…")
    st.caption(
        "💡 Seuls les métadonnées + aperçu (255 car.) sont téléchargés. "
        "Le corps complet est chargé à la demande. Cela permet d'indexer "
        "10 000+ emails en quelques minutes."
    )

    status_ph  = st.empty()
    c1, c2, c3, c4 = st.columns(4)
    new_ph     = c1.empty()
    upd_ph     = c2.empty()
    fold_ph    = c3.empty()
    warn_ph    = c4.empty()
    progress   = st.progress(0.0)
    log_ph     = st.empty()
    log_lines: list[str] = []

    indexer = EmailIndexer(access_token, user_id)

    def on_status(msg: str, r: SyncResult):
        status_ph.info(f"📬 {msg}")
        new_ph.metric("🆕 Nouveaux",   f"{r.emails_new:,}")
        upd_ph.metric("✏️ Màj",        f"{r.emails_updated:,}")
        done = r.folders_done + r.folders_skip
        fold_ph.metric("📁 Dossiers",  f"{done}/{r.total_folders}")
        warn_ph.metric("⚠️ Alertes",   str(len(r.warnings)))
        if r.total_folders:
            progress.progress(min(done / r.total_folders, 1.0))
        log_lines.append(msg)
        log_ph.code("\n".join(log_lines[-10:]))

    try:
        result = indexer.sync(force_full=force_full, on_status=on_status)
        progress.progress(1.0)
        log_ph.empty()

        # ── Résumé ────────────────────────────────────────────────────────────
        parts = []
        if result.emails_new:
            parts.append(f"**{result.emails_new:,}** nouveaux")
        if result.emails_updated:
            parts.append(f"**{result.emails_updated:,}** mis à jour")
        if not parts:
            parts = ["aucun changement"]

        st.success(f"✅ Synchronisation {mode} terminée — {', '.join(parts)}")

        # ── Tableau de vérification par dossier ───────────────────────────────
        if result.folder_details:
            with st.expander(
                f"📊 Détail par dossier ({len(result.folder_details)} dossiers)",
                expanded=bool(result.warnings),
            ):
                for info in result.folder_details:
                    gap  = info.expected - info.indexed
                    icon = "✅" if gap <= 10 else "⚠️"
                    st.markdown(
                        f"{icon} **{info.name}** — "
                        f"{info.indexed:,} / {info.expected:,} indexés "
                        f"(+{info.new} nouveaux)"
                    )

        # ── Avertissements dossiers incomplets ────────────────────────────────
        if result.warnings:
            st.warning(
                "Certains dossiers semblent incomplets. "
                "Relancez une **Sync complète** pour corriger."
            )
            with st.expander(f"⚠️ {len(result.warnings)} avertissement(s)"):
                for w in result.warnings:
                    st.caption(w)

        if result.errors:
            with st.expander(f"🔴 {len(result.errors)} erreur(s)"):
                for e in result.errors:
                    st.caption(e)

        st.session_state.pop("show_sync", None)
        st.session_state.pop("sync_force_full", None)
        st.rerun()

    except PermissionError as e:
        st.error(f"🔑 {e}")
        for k in ["user_info", "access_token", "refresh_token"]:
            st.session_state.pop(k, None)
    except Exception as e:
        st.error(f"❌ {e}")
        st.caption(
            "Les dossiers déjà terminés sont sauvegardés. "
            "Relancez pour reprendre là où la sync s'est arrêtée."
        )
        st.session_state.pop("show_sync", None)
        st.session_state.pop("sync_force_full", None)


def show_results(db, keywords, folder_filter, folder_ids, tab_key, date_from=None, date_to=None, user_id=None):
    page_key = f"page_{tab_key}"
    if page_key not in st.session_state:
        st.session_state[page_key] = 0
    page = st.session_state[page_key]

    if not keywords:
        st.info("ℹ️ Entrez au moins un mot-clé pour lancer la recherche.")
        return

    with st.spinner("Recherche…"):
        results, total = db.search_emails(
            keywords=keywords,
            folder_filter=folder_filter,
            folder_ids=folder_ids,
            date_from=date_from,
            date_to=date_to,
            limit=PAGE_SIZE,
            offset=page * PAGE_SIZE,
        )
    if total == 0:
        st.warning(f"Aucun résultat pour : **{' + '.join(keywords)}**  "
                   "(tous les mots-clés doivent être présents)")
        return

    total_pages = max(1, (total - 1) // PAGE_SIZE + 1)
    st.markdown(
        f"**{total:,} résultat(s)** — "
        + " &nbsp;`ET`&nbsp; ".join(f"`{k}`" for k in keywords)
        + f" &nbsp;|&nbsp; page **{page+1}** / {total_pages}",
        unsafe_allow_html=True)
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
            f"— {fmt_date(email['received_datetime'])}"):

            c1, c2 = st.columns([4, 1])
            with c1:
                st.markdown(
                    f"{badges}<br>"
                    f"<b>Objet :</b> {subject_hl}<br>"
                    f"<b>De :</b> {email['sender_name']} &lt;{email['sender_email']}&gt;<br>"
                    f"<b>À :</b> {(email['recipients'] or '—')[:300]}<br>"
                    f"<b>Dossier :</b> 📁 {email['folder_name']}&nbsp;&nbsp;"
                    f"<b>Date :</b> {fmt_date(email['received_datetime'])}",
                    unsafe_allow_html=True)
            with c2:
                if email.get("web_link"):
                    st.markdown(
                        f'<a href="{email["web_link"]}" target="_blank">' +
                        "<button style='background:#0078d4;color:#fff;border:none;"
                        "padding:8px 14px;border-radius:6px;cursor:pointer;"
                        "width:100%;font-size:.9em'>📬 Ouvrir Outlook</button></a>",
                        unsafe_allow_html=True)

            st.markdown("---")
            st.markdown(f"<div style='color:#333;line-height:1.6'>{preview_hl}</div>",
                        unsafe_allow_html=True)

            if st.button("📄 Contenu complet", key=f"full_{tab_key}_{email['id']}"):
                full = db.get_email_detail(email["id"])
                body = full.get("body", "") if full else ""
                if not body:
                    # Corps non stocké lors de la sync bulk → chargement à la demande
                    token = get_access_token()
                    if token:
                        with st.spinner("Chargement du corps…"):
                            from email_indexer import EmailIndexer as _EI
                            body = _EI(token, user_id or st.session_state.get('user_info',{}).get('id','x')).get_email_body(email["id"])
                            if body and full:
                                # Mise en cache en base pour les prochaines fois
                                full["body"] = body
                                db.upsert_email(dict(full))
                if body:
                    st.text_area("Corps complet", body, height=350,
                                 key=f"body_{tab_key}_{email['id']}")
                else:
                    st.info("Corps non disponible. Ouvrez l'email dans Outlook.")

    if total_pages > 1:
        st.markdown("---")
        pc1, pc2, pc3 = st.columns([1, 4, 1])
        with pc1:
            if page > 0 and st.button("← Précédent", key=f"prev_{tab_key}"):
                st.session_state[page_key] -= 1
                st.rerun()
        with pc2:
            st.markdown(f"<p style='text-align:center'>Page <b>{page+1}</b> / {total_pages}</p>",
                        unsafe_allow_html=True)
        with pc3:
            if page < total_pages - 1 and st.button("Suivant →", key=f"next_{tab_key}"):
                st.session_state[page_key] += 1
                st.rerun()


# ── Page principale ───────────────────────────────────────────────────────────

def page_main():
    user_info = st.session_state["user_info"]
    user_id   = user_info.get("id") or user_info.get("mail") or "unknown"
    db        = Database(user_id)
    stats     = db.get_stats()
    folders   = db.get_folders()

    with st.sidebar:
        st.markdown(f"### 👤 {user_info.get('displayName', 'Utilisateur')}")
        st.caption(user_info.get("mail") or user_info.get("userPrincipalName", ""))
        st.markdown("---")
        col_a, col_b = st.columns(2)
        with col_a: st.metric("Emails", f"{stats['total_emails']:,}")
        with col_b: st.metric("Dossiers", stats["total_folders"])
        if stats.get("last_sync"):
            try:
                dt = datetime.fromisoformat(stats["last_sync"])
                st.caption(f"🔄 Dernière sync : {dt.strftime('%d/%m/%Y %H:%M')}")
            except Exception:
                pass
        st.markdown("---")
        if st.button("⚡ Sync incrémentale", use_container_width=True, type="primary",
                     help="Traite uniquement les emails nouveaux, modifiés ou supprimés "
                          "depuis la dernière synchronisation."):
            st.session_state["show_sync"]       = True
            st.session_state["sync_force_full"] = False
            st.rerun()
        if st.button("🔁 Sync complète (reset)", use_container_width=True,
                     help="Réinitialise tous les delta tokens et re-télécharge "
                          "l'intégralité de la boîte mail. À utiliser si vous soupçonnez "
                          "un désynchronisation ou après une longue absence."):
            st.session_state["show_sync"]       = True
            st.session_state["sync_force_full"] = True
            st.rerun()
        with st.expander("📊 Emails par dossier"):
            for row in stats["by_folder"][:15]:
                st.markdown(f"**{row['folder_name']}** : {row['cnt']:,}")
        st.markdown("---")
        st.markdown("---")
        with st.expander("🗑️ Réinitialiser la base"):
            st.caption(
                "Supprime TOUS les emails indexés et repart de zéro. "
                "Une sync complète sera nécessaire ensuite."
            )
            if st.button("🗑️ Vider la base de données", use_container_width=True,
                         type="secondary"):
                db.reset_all()
                for k in ["show_sync","sync_force_full","search_kw"]:
                    st.session_state.pop(k, None)
                st.success("Base vidée. Lancez une sync complète.")
                st.rerun()

        if st.button("🚪 Se déconnecter", use_container_width=True):
            for k in ["user_info","access_token","refresh_token",
                      "app_authenticated","device_flow","show_sync","search_kw"]:
                st.session_state.pop(k, None)
            st.rerun()

    if st.session_state.get("show_sync"):
        token = get_access_token()
        if token:
            force_full = st.session_state.get("sync_force_full", False)
            run_sync(token, user_id, force_full=force_full)
        else:
            st.error("Session expirée, reconnectez-vous.")
            st.session_state.pop("show_sync", None)
            st.session_state.pop("sync_force_full", None)
        return

    st.markdown("# 📧 Recherche d'emails")
    with st.form("search_form"):
        kw_input = st.text_input(
            "🔍 Mots-clés",
            placeholder="Ex: réunion budget   ou   réunion, budget   ou   réunion; budget",
            help=(
                "Séparez par virgule, point-virgule ou double-espace. "
                "Logique ET : TOUS les mots-clés doivent être présents."
            ),
        )
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            date_from = st.date_input(
                "📅 Date de début",
                value=None,
                format="DD/MM/YYYY",       # format français
                help="Laisser vide pour ne pas filtrer (JJ/MM/AAAA)",
            )
        with col_d2:
            date_to = st.date_input(
                "📅 Date de fin",
                value=None,
                format="DD/MM/YYYY",       # format français
                help="Laisser vide pour ne pas filtrer (JJ/MM/AAAA)",
            )
        submitted = st.form_submit_button("Rechercher", type="primary", use_container_width=True)

    # Découpe sur virgule, point-virgule OU espaces multiples
    import re as _re
    keywords = (
        [k for k in _re.split(r"[,;]+|\s{2,}", kw_input.strip()) if k.strip()]
        if kw_input else []
    )
    keywords = [k.strip() for k in keywords if k.strip()]

    # Convertit les dates en ISO string (None si non renseigné)
    # date_input renvoie un objet date Python → isoformat() donne "YYYY-MM-DD" pour la DB
    df_iso = date_from.isoformat() if date_from else None
    dt_iso = date_to.isoformat()   if date_to   else None

    if submitted and not keywords:
        st.warning("Veuillez entrer au moins un mot-clé.")

    tab_all, tab_no_sent, tab_specific = st.tabs([
        "📬 Toute la boîte mail",
        "📥 Hors Envoyés / Supprimés",
        "📁 Dossier spécifique",
    ])

    if submitted:
        st.session_state["search_date_from"] = df_iso
        st.session_state["search_date_to"]   = dt_iso

    kw_active  = keywords if submitted else st.session_state.get("search_kw", [])
    df_active  = df_iso   if submitted else st.session_state.get("search_date_from")
    dt_active  = dt_iso   if submitted else st.session_state.get("search_date_to")

    # Confirmation visuelle des filtres actifs
    if kw_active:
        tags     = "  ".join(f"`{k}`" for k in kw_active)
        date_str = ""
        def _fr(iso: str) -> str:
            """Convertit YYYY-MM-DD en JJ/MM/AAAA pour l'affichage."""
            try:
                y, m, d = iso[:10].split("-")
                return f"{d}/{m}/{y}"
            except Exception:
                return iso

        if df_active and dt_active:
            date_str = f" &nbsp;|&nbsp; 📅 du **{_fr(df_active)}** au **{_fr(dt_active)}**"
        elif df_active:
            date_str = f" &nbsp;|&nbsp; 📅 depuis le **{_fr(df_active)}**"
        elif dt_active:
            date_str = f" &nbsp;|&nbsp; 📅 jusqu'au **{_fr(dt_active)}**"
        st.markdown(
            f"🔍 {tags}{date_str} &nbsp;—&nbsp; logique **ET**",
            unsafe_allow_html=True,
        )

    with tab_all:
        if kw_active:
            show_results(db, kw_active, "all", None, "all", df_active, dt_active, user_id)
        else: st.info("ℹ️ Entrez des mots-clés ci-dessus pour rechercher.")

    with tab_no_sent:
        if kw_active:
            show_results(db, kw_active, "no_sent_deleted", None, "no_sent", df_active, dt_active, user_id)
        else: st.info("ℹ️ Entrez des mots-clés ci-dessus pour rechercher.")

    with tab_specific:
        if not folders:
            st.info("Aucun dossier indexé. Lancez une synchronisation d'abord.")
        else:
            folder_map = {f["display_path"]: f["id"]
                          for f in sorted(folders, key=lambda x: x["display_path"])}
            selected  = st.selectbox("Choisir un dossier", options=list(folder_map.keys()))
            folder_id = folder_map.get(selected)
            if kw_active and folder_id:
                show_results(db, kw_active, "specific", [folder_id], "specific", df_active, dt_active, user_id)
            elif not kw_active:
                st.info("ℹ️ Entrez des mots-clés ci-dessus pour rechercher.")

    if submitted and keywords:
        st.session_state["search_kw"] = keywords
        # dates already saved above


# ── Point d'entrée ───────────────────────────────────────────────────────────

def main():
    if not check_password():
        st.stop()
    if is_logged_in():
        page_main()
    else:
        page_login()

if __name__ == "__main__":
    main()