import requests
from msal import ConfidentialClientApplication
import config

# Test token
app = ConfidentialClientApplication(
    config.CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
    client_credential=config.CLIENT_SECRET
)
result = app.acquire_token_for_client(scopes=config.SCOPES)
print("Token OK:", "access_token" in result)
print("Erreur:", result.get("error_description", "aucune"))

# Test appel API
token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}
url = f"{config.GRAPH_ENDPOINT}/users/{config.USER_EMAIL}/messages?$top=5"
response = requests.get(url, headers=headers)

print("Status:", response.status_code)
print("Réponse:", response.json())
