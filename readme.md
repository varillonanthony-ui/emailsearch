# 📧 Email Search — Office 365 + Streamlit

Application Streamlit permettant d'indexer et de rechercher dans tous vos emails Office 365,
avec authentification sécurisée par tenant Microsoft.

---

## 🏗️ Architecture

```
email-search/
├── app.py            # Interface Streamlit (UI, auth flow, recherche)
├── auth.py           # Authentification Microsoft OAuth 2.0 via MSAL
├── database.py       # SQLite par utilisateur + FTS5 plein texte
├── email_indexer.py  # Sync Graph API avec delta tokens
├── requirements.txt
├── .env.example      # Template de configuration
├── .gitignore
└── data/             # Créé automatiquement — bases SQLite (⚠️ ignoré par git)
    ├── user1-id.db
    └── user2-id.db
```

**Points clés :**
- Chaque utilisateur a sa propre base SQLite dans `data/`
- L'authentification est restreinte à votre tenant Azure AD
- La synchronisation utilise les **delta tokens** Graph API (seuls les nouveaux emails sont re-téléchargés lors des syncs suivantes)
- La recherche utilise **SQLite FTS5** avec logique ET sur tous les mots-clés

---

## ⚙️ Pré-requis

- Python 3.11+
- Un compte Office 365 (Microsoft 365)
- Accès au portail Azure pour créer une App Registration

---

## 🔐 Étape 1 — Créer l'App Registration Azure AD

### 1.1 Créer l'application

1. Connectez-vous à [portal.azure.com](https://portal.azure.com)
2. **Azure Active Directory** → **Inscriptions d'applications** → **Nouvelle inscription**
3. Remplissez :
   - **Nom** : `Email Search App` (ou ce que vous voulez)
   - **Types de comptes pris en charge** : `Comptes dans cet annuaire organisationnel uniquement (tenant unique)` ← **Important pour restreindre l'accès**
   - **URI de redirection** : sélectionnez `Web` et entrez `http://localhost:8501`
4. Cliquez **Inscrire**

### 1.2 Récupérer les identifiants

Sur la page de l'application, notez :
- **ID d'application (client)** → `AZURE_CLIENT_ID`
- **ID de l'annuaire (tenant)** → `AZURE_TENANT_ID`

### 1.3 Créer un secret client

1. **Certificats & secrets** → **Nouveau secret client**
2. Donnez une description, choisissez une expiration
3. **Copiez la valeur immédiatement** (elle ne sera plus visible) → `AZURE_CLIENT_SECRET`

### 1.4 Configurer les autorisations API

1. **Autorisations API** → **Ajouter une autorisation** → **Microsoft Graph** → **Autorisations déléguées**
2. Ajoutez :
   - `Mail.Read` — lecture des emails
   - `User.Read` — profil utilisateur
   - `offline_access` — refresh tokens
3. Cliquez **Accorder le consentement administrateur** (bouton en haut)

### 1.5 Configurer l'URI de redirection (si ce n'est pas fait)

1. **Authentification** → **Ajouter une plateforme** → **Web**
2. URI de redirection : `http://localhost:8501`
3. Cochez **Jetons d'accès** et **Jetons d'ID**

> **En production** (Streamlit Cloud, Azure App Service…), ajoutez l'URL publique en plus de localhost.

---

## 🚀 Étape 2 — Installation et configuration

```bash
# Cloner le dépôt
git clone https://github.com/votre-org/email-search.git
cd email-search

# Créer un environnement virtuel
python -m venv .venv
source .venv/bin/activate   # Linux/Mac
# ou .venv\Scripts\activate  # Windows

# Installer les dépendances
pip install -r requirements.txt

# Configurer l'environnement
cp .env.example .env
# Éditez .env et remplissez les 4 variables
```

---

## ▶️ Étape 3 — Lancer l'application

```bash
streamlit run app.py
```

L'application s'ouvre sur `http://localhost:8501`.

---

## 📖 Utilisation

### Première connexion

1. Cliquez **Se connecter avec Microsoft**
2. Authentifiez-vous avec votre compte Office 365
3. Autorisez l'accès à votre messagerie

### Synchronisation

1. Dans la sidebar, cliquez **🔄 Synchroniser les emails**
2. La première synchronisation peut prendre plusieurs minutes selon le volume
3. Les synchronisations suivantes sont **incrémentales** (delta sync) — beaucoup plus rapides

### Recherche

| Champ | Description |
|-------|-------------|
| Mots-clés | Séparés par des virgules. **Tous** doivent être présents (logique ET) |
| Toute la boîte | Recherche dans tous les dossiers |
| Hors Envoyés/Supprimés | Exclut : Éléments envoyés, Éléments supprimés, Brouillons, Courrier indésirable |
| Dossier spécifique | Sélecteur de dossier (tous les sous-dossiers disponibles) |

**Exemples :**
- `facture, 2024` → emails contenant "facture" ET "2024"
- `réunion, équipe, projet` → emails contenant les 3 termes

### Multi-utilisateurs

Chaque collègue du même tenant peut se connecter avec son propre compte.
Sa base de données (`data/<user-id>.db`) est **totalement isolée**.

---

## 🚢 Déploiement en production

### Streamlit Cloud

1. Pushez le code sur GitHub (sans `.env` ni `data/`)
2. Créez une app sur [share.streamlit.io](https://share.streamlit.io)
3. Dans **Settings → Secrets**, ajoutez :

```toml
AZURE_CLIENT_ID     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
AZURE_CLIENT_SECRET = "your_secret"
AZURE_TENANT_ID     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
REDIRECT_URI        = "https://your-app.streamlit.app"
```

4. Dans Azure AD, ajoutez `https://your-app.streamlit.app` comme URI de redirection supplémentaire

> ⚠️ Sur Streamlit Cloud, les fichiers `data/*.db` ne persistent pas entre les redémarrages.
> Pour un déploiement pérenne, utilisez Azure App Service + un volume persistant,
> ou remplacez SQLite par PostgreSQL.

### Azure App Service / Docker

Un `Dockerfile` minimal :

```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

---

## 🔧 Variables d'environnement

| Variable | Obligatoire | Description |
|----------|-------------|-------------|
| `AZURE_CLIENT_ID` | ✅ | ID de l'application Azure AD |
| `AZURE_CLIENT_SECRET` | ✅ | Secret client Azure AD |
| `AZURE_TENANT_ID` | ✅ | ID du tenant (restreint l'accès au tenant) |
| `REDIRECT_URI` | ✅ | URI de redirection après OAuth (défaut : `http://localhost:8501`) |

---

## 🛡️ Sécurité

- L'authentification via `AUTHORITY = login.microsoftonline.com/{TENANT_ID}` garantit que **seuls les comptes de votre tenant** peuvent se connecter.
- Les bases de données sont nommées par l'ID unique Microsoft de l'utilisateur — pas d'accès croisé possible.
- Les secrets ne sont jamais exposés côté client.
- Le dossier `data/` est dans `.gitignore` — les emails ne sont jamais commités.

---

## ❓ Dépannage

**"redirect_uri_mismatch"** → L'URI de redirection dans `.env` ne correspond pas à celle enregistrée dans Azure AD.

**"AADSTS50011"** → Même problème d'URI de redirection.

**"AADSTS700016" / application not found** → `AZURE_CLIENT_ID` ou `AZURE_TENANT_ID` incorrect.

**Token expiré pendant la sync** → Reconnectez-vous et relancez la synchronisation (le delta token est conservé, la sync reprendra là où elle s'est arrêtée).

**Recherche trop lente** → La première indexation doit être complète. Vérifiez le nombre d'emails dans la sidebar.