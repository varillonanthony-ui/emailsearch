import os

# Essayer Streamlit secrets (cloud) sinon .env (local)
try:
    import streamlit as st
    CLIENT_ID     = st.secrets["AZURE_CLIENT_ID"]
    CLIENT_SECRET = st.secrets["AZURE_CLIENT_SECRET"]
    TENANT_ID     = st.secrets["AZURE_TENANT_ID"]
    USER_EMAIL    = st.secrets["USER_EMAIL"]
except Exception:
    from dotenv import load_dotenv
    load_dotenv()
    CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    TENANT_ID     = os.getenv("AZURE_TENANT_ID")
    USER_EMAIL    = os.getenv("USER_EMAIL")

DB_PATH        = "data/emails.db"
INDEX_PATH     = "data/email_index"
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
SCOPES         = ["https://graph.microsoft.com/.default"]
