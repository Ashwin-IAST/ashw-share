import os
import requests
import msal
from urllib.parse import quote
import sys
import io

# === Fix Unicode Output ===
print("🔧 Setting up Unicode-compatible stdout...")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# === Get filename from CLI ===
print("📥 Checking command line arguments...")
if len(sys.argv) < 2:
    print("❌ Error: No filename provided.")
    sys.exit(1)

filename = sys.argv[1]
print(f"📄 Requested filename: {filename}")

# === Azure App Credentials (Read from Environment Variables) ===
print("🔐 Reading Azure credentials from environment variables...")
client_id = os.environ.get("AZURE_CLIENT_ID")
tenant_id = os.environ.get("AZURE_TENANT_ID")

if not client_id:
    print("❌ Error: AZURE_CLIENT_ID environment variable not set.")
    sys.exit(1)
if not tenant_id:
    print("❌ Error: AZURE_TENANT_ID environment variable not set.")
    sys.exit(1)

print("✅ Azure credentials loaded successfully.")

authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["User.Read", "Files.ReadWrite.All"]

# === SharePoint File Info ===
print("🔗 Encoding filename for SharePoint URL...")
file_path_on_sharepoint = quote(filename)

# === Save path: same directory as script ===
print("💾 Preparing local file path...")
script_dir = os.path.dirname(os.path.abspath(__file__))
local_path = os.path.join(script_dir, filename)
print(f"📁 Local path will be: {local_path}")

# === SharePoint Site Info ===
site_domain = "iastsoftware20.sharepoint.com"
site_path = "sites/Testingversions"
print(f"🌐 SharePoint site: {site_domain}/{site_path}")

# === Step 1: Authenticate ===
print("🔐 Starting authentication process...")
app = msal.PublicClientApplication(client_id=client_id, authority=authority)
print("💬 Prompting user for interactive sign-in...")
result = app.acquire_token_interactive(scopes=scopes)

if "access_token" not in result:
    print("❌ Authentication failed.")
    sys.exit(1)

print("✅ Authentication successful.")
headers = {"Authorization": f"Bearer {result['access_token']}"}

# === Step 2: Get Site ID ===
print("🔍 Retrieving SharePoint site ID...")
site_info_url = f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/{site_path}"
print(f"🌐 GET {site_info_url}")
site_info_response = requests.get(site_info_url, headers=headers)
if site_info_response.status_code != 200:
    print("❌ Failed to retrieve SharePoint site info.")
    print(f"🔴 Response code: {site_info_response.status_code}")
    print(site_info_response.text)
    sys.exit(1)

site_id = site_info_response.json()["id"]
print(f"✅ Site ID: {site_id}")

# === Step 3: Get Drive ID ===
print("📦 Retrieving document library (drive) ID...")
drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
print(f"🌐 GET {drive_url}")
drive_response = requests.get(drive_url, headers=headers)
if drive_response.status_code != 200:
    print("❌ Failed to get drive info.")
    print(f"🔴 Response code: {drive_response.status_code}")
    print(drive_response.text)
    sys.exit(1)

drive_id = drive_response.json()["id"]
print(f"✅ Drive ID: {drive_id}")

# === Step 4: Download File ===
print(f"⬇️ Attempting to download file: {filename}")
download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path_on_sharepoint}:/content"
print(f"🌐 GET {download_url}")
download_response = requests.get(download_url, headers=headers)

if download_response.status_code == 200:
    print("✅ File found. Downloading...")
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
