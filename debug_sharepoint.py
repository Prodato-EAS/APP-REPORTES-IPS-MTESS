import os
import requests
import msal
import json
from dotenv import load_dotenv

# Load Environment from .env
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
LIST_NAME = "Auditoria_General"

print("--- DIAGNOSTIC START ---")

# 1. Authenticate
try:
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        token = result["access_token"]
        print("✅ Auth Token Acquired")
    else:
        print(f"❌ Auth Failed: {result.get('error_description')}")
        exit()
except Exception as e:
    print(f"❌ Auth Exception: {e}")
    exit()

# 2. Get Site/List IDs
headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
site_name = "PruebaFlujo"
tenant_host = "hpauditorescontadores.sharepoint.com"

# Parse URL just in case, but using defaults from previous context if env var is simple
if SHAREPOINT_SITE_URL:
    try:
        parts = SHAREPOINT_SITE_URL.split("/sites/")
        if len(parts) > 1:
            site_name = parts[1].split("/")[0]
            tenant_host = parts[0].replace("https://", "")
    except: pass

print(f"Target: {tenant_host} / {site_name}")

try:
    # Get Site
    url_site = f"https://graph.microsoft.com/v1.0/sites/{tenant_host}:/sites/{site_name}"
    r_site = requests.get(url_site, headers=headers)
    if r_site.status_code != 200:
        print(f"❌ Get Site Failed: {r_site.text}")
        exit()
    site_id = r_site.json()["id"]
    print(f"✅ Site ID: {site_id}")

    # Get List
    url_list = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LIST_NAME}"
    r_list = requests.get(url_list, headers=headers)
    if r_list.status_code != 200:
        # Try getting by name
        print(f"⚠️ Direct list get failed, listing all...")
        url_lists = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
        r_lists = requests.get(url_lists, headers=headers)
        target_list = next((l for l in r_lists.json().get('value', []) if l['displayName'] == LIST_NAME), None)
        if not target_list:
            print(f"❌ List '{LIST_NAME}' not found.")
            exit()
        list_id = target_list['id']
    else:
        list_id = r_list.json()["id"]
    
    print(f"✅ List ID: {list_id}")

    # 3. Inspect Columns (Fields) of the first item
    print("\n--- INSPECTING ITEM FIELDS ---")
    url_items = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields&$top=1"
    r_items = requests.get(url_items, headers=headers)
    items = r_items.json().get("value", [])
    
    if not items:
        print("❌ No items found to inspect.")
    else:
        item = items[0]
        fields = item.get("fields", {})
        print(f"Item ID: {item.get('id')}")
        print("Available Fields (Keys):")
        # Print all keys that look like 'REVISADO'
        keys = list(fields.keys())
        potential_matches = [k for k in keys if "REVISADO" in k.upper() or "ESTADO" in k.upper()]
        print(f"Potential Matches for 'REVISADO': {potential_matches}")
        
        if "REVISADO" in fields:
            print(f"Current Value of REVISADO: '{fields['REVISADO']}'")
        else:
            print("❌ field 'REVISADO' NOT FOUND exactly.")

        # 4. Try PATCH
        print("\n--- ATTEMPTING PATCH ---")
        target_id = item.get("id")
        patch_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{target_id}/fields"
        
        # Try updating to 'REVISADO'
        body = {"REVISADO": "REVISADO"}
        print(f"Patching {target_id} with {body} ...")
        r_patch = requests.patch(patch_url, headers=headers, json=body)
        print(f"Status Code: {r_patch.status_code}")
        print(f"Response Body: {r_patch.text}")

except Exception as e:
    print(f"❌ Script Exception: {e}")
