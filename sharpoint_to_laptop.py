import os
import requests
import msal
from urllib.parse import quote
import sys
import io

# === Fix Unicode Output ===
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# === Get filename from CLI ===
if len(sys.argv) < 2:
    print("❌ Error: No filename provided.")
    sys.exit(1)

filename = sys.argv[1]

# === Azure App Credentials (Read from Environment Variables) ===
# It's crucial to retrieve these from environment variables for security.
# Using .get() with a default of None allows you to check if the variable was set.
client_id = os.environ.get("AZURE_CLIENT_ID") # Mapped from Jenkins credential ID
tenant_id = os.environ.get("AZURE_TENANT_ID") # Mapped from Jenkins credential ID

# Basic validation that environment variables are set
if not client_id:
    print("❌ Error: AZURE_CLIENT_ID environment variable not set.")
    sys.exit(1)
if not tenant_id:
    print("❌ Error: AZURE_TENANT_ID environment variable not set.")
    sys.exit(1)

authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["User.Read", "Files.ReadWrite.All"]

# === SharePoint File Info ===
file_path_on_sharepoint = quote(filename)

# === Save path: same directory as script ===
script_dir = os.path.dirname(os.path.abspath(__file__))
local_path = os.path.join(script_dir, filename)

# === SharePoint Site Info ===
site_domain = "iastsoftware20.sharepoint.com"
site_path = "sites/Testingversions"

# === Step 1: Authenticate ===
app = msal.PublicClientApplication(client_id=client_id, authority=authority)
result = app.acquire_token_interactive(scopes=scopes)

if "access_token" not in result:
    print("❌ Authentication failed.")
    sys.exit(1)

headers = {"Authorization": f"Bearer {result['access_token']}"}

# === Step 2: Get Site ID ===
site_info_url = f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/{site_path}"
site_info_response = requests.get(site_info_url, headers=headers)
if site_info_response.status_code != 200:
    print("❌ Failed to retrieve SharePoint site info.")
    sys.exit(1)

site_id = site_info_response.json()["id"]

# === Step 3: Get Drive ID ===
drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
drive_response = requests.get(drive_url, headers=headers)
if drive_response.status_code != 200:
    print("❌ Failed to get drive info.")
    sys.exit(1)

drive_id = drive_response.json()["id"]

# === Step 4: Download File ===
download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path_on_sharepoint}:/content"
download_response = requests.get(download_url, headers=headers)

if download_response.status_code == 200:
    with open(local_path, "wb") as file:
        file.write(download_response.content)
    print(f"✅ File downloaded successfully to: {local_path}")
    sys.exit(0)
elif download_response.status_code == 404:
    print(f"❌ File '{filename}' not found on SharePoint. Please check the name and try again.")
    sys.exit(1)
else:
    print(f"❌ Download failed. HTTP status: {download_response.status_code}")
    print(download_response.text)
    sys.exit(1)
