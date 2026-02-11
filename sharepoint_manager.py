import os
import threading
import pandas as pd
import requests
import msal
import time
import json
from dotenv import load_dotenv

load_dotenv()

class SharePointManager:
    """
    Manages connections to SharePoint via Microsoft Graph API.
    """
    def __init__(self):
        self.site_url_full = os.environ.get("SHAREPOINT_SITE_URL", "")
        # Extract site name/host from URL or env. 
        # Example: https://hpauditorescontadores.sharepoint.com/sites/PruebaFlujo
        # Host: hpauditorescontadores.sharepoint.com
        # Site Path: /sites/PruebaFlujo (or just name 'PruebaFlujo' for graph)
        
        self.client_id = os.environ.get("CLIENT_ID")
        self.client_secret = os.environ.get("CLIENT_SECRET")
        self.tenant_id = os.environ.get("TENANT_ID")
        self.list_name_ips = "Auditoria_General"
        self.list_name_mtess = "Auditoria_MTESS_IPS"
        self.list_name_whitelist = "whitelist_ips_mtess"
        
        self.list_ids = {"IPS": None, "MTESS": None, "WHITELIST": None}
        
        self.dfs = {
            "IPS": pd.DataFrame(),
            "MTESS": pd.DataFrame()
        }
        
        self.webhook_url = os.environ.get("RPA_WEBHOOK_URL")

        # Site Name extraction for Graph query
        # Assuming URL format: https://<tenant>.sharepoint.com/sites/<sitename>
        try:
            parts = self.site_url_full.split("/sites/")
            if len(parts) > 1:
                self.site_name = parts[1].split("/")[0] # Get 'PruebaFlujo'
                self.tenant_host = parts[0].replace("https://", "") # Get 'hpauditorescontadores.sharepoint.com'
            else:
                self.site_name = "PruebaFlujo" # Fallback/Default
                self.tenant_host = "hpauditorescontadores.sharepoint.com"
        except:
             self.site_name = "PruebaFlujo"
             self.tenant_host = "hpauditorescontadores.sharepoint.com"

        self.site_id = None
        
        self._authenticate()
        
        # Pre-fetch IDs
        if self.access_token:
            self._get_site_id()
            self._get_list_ids()

        self.state_file = "version.json"
        self.last_change_times = self._read_state() # Returns dict {"IPS": ts, "MTESS": ts}
        self.last_local_update_times = {"IPS": 0, "MTESS": 0}

    def _save_state(self, times_dict):
        """Writes timestamps to version.json."""
        try:
            with open(self.state_file, 'w') as f:
                json.dump(times_dict, f)
        except Exception as e:
            print(f"Error saving state: {e}")

    def _read_state(self):
        """Reads timestamps from version.json."""
        default_state = {"IPS": 0, "MTESS": 0}
        try:
            if os.path.exists(self.state_file):
                with open(self.state_file, 'r') as f:
                    data = json.load(f)
                    # Backwards compatibility check
                    if "version" in data: 
                         # Old format, reset or migrate? Reset safe.
                         return default_state
                    return {**default_state, **data} # Merge defaults
        except Exception as e:
            print(f"Error reading state: {e}")
        return default_state

    def check_version(self):
        """Checks shared file versions. Fetches if newer."""
        # This returns the current state dict for the Server Loop to inspect.
        # But crucially, it should TRIGGER fetch if file is newer.
        
        file_versions = self._read_state()
        updated_sources = []
        
        for source in ["IPS", "MTESS"]:
            # Grace Period Logic per source
            if (time.time() - self.last_local_update_times[source]) < 5:
                continue

            # Check if File > Memory
            if file_versions[source] > (self.last_change_times[source] + 0.1):
                print(f"🔄 Syncing {source}: File ({file_versions[source]}) > Memory ({self.last_change_times[source]})")
                self.fetch_data(source)
                updated_sources.append(source)
        
        return self.last_change_times, updated_sources # Return tuple

    def _authenticate(self):
        """Authenticates using MSAL."""
        try:
            authority = f"https://login.microsoftonline.com/{self.tenant_id}"
            app = msal.ConfidentialClientApplication(
                self.client_id, 
                authority=authority, 
                client_credential=self.client_secret
            )
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                print("✅ Authentication successful (Graph API).")
            else:
                print(f"❌ Authentication failed: {result.get('error_description')}")
        except Exception as e:
            print(f"❌ Authentication error: {e}")

    def _get_site_id(self):
        """Fetches Site ID from Graph."""
        if not self.access_token: return
        try:
            # URL format: https://graph.microsoft.com/v1.0/sites/{host}:/sites/{name}
            url = f"https://graph.microsoft.com/v1.0/sites/{self.tenant_host}:/sites/{self.site_name}"
            headers = {"Authorization": f"Bearer {self.access_token}"}
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                self.site_id = response.json()["id"]
                print(f"✅ Site ID: {self.site_id}")
            else:
                print(f"❌ Failed to get Site ID: {response.text}")
        except Exception as e:
            print(f"Error getting Site ID: {e}")

    def _get_list_ids(self):
        """Fetches List IDs for both IPS and MTESS."""
        if not self.access_token or not self.site_id: return
        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists"
            headers = {"Authorization": f"Bearer {self.access_token}"}
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                lists = response.json().get("value", [])
                
                # Reset
                self.list_ids = {"IPS": None, "MTESS": None, "WHITELIST": None}
                
                for lst in lists:
                    if lst["displayName"] == self.list_name_ips:
                        self.list_ids["IPS"] = lst["id"]
                        print(f"✅ IPS List ID: {lst['id']}")
                    elif lst["displayName"] == self.list_name_mtess:
                        self.list_ids["MTESS"] = lst["id"]
                        print(f"✅ MTESS List ID: {lst['id']}")
                    elif lst["displayName"] == self.list_name_whitelist:
                        self.list_ids["WHITELIST"] = lst["id"]
                        print(f"✅ Whitelist List ID: {lst['id']}")
                
                if not self.list_ids["IPS"]: print(f"❌ List '{self.list_name_ips}' not found.")
                if not self.list_ids["MTESS"]: print(f"❌ List '{self.list_name_mtess}' not found.")
                if not self.list_ids["WHITELIST"]: print(f"❌ List '{self.list_name_whitelist}' not found.")

            else:
                 print(f"❌ Failed to get Lists: {response.text}")
        except Exception as e:
             print(f"Error getting List IDs: {e}")

    def fetch_data(self, source="IPS"):
        """
        Fetches items from the specified source ('IPS' or 'MTESS').
        """
        if not self.access_token or not self.site_id:
             self._authenticate()
             self._get_site_id()
             self._get_list_ids()

        target_list_id = self.list_ids.get(source)

        if not target_list_id:
            print(f"Missing List ID for {source}. Re-scanning...")
            self._get_list_ids()
            target_list_id = self.list_ids.get(source)
            if not target_list_id:
                 print(f"Still missing List ID for {source}. Aborting fetch.")
                 return self.dfs[source]

        try:
            # Expand fields
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{target_list_id}/items?expand=fields&$top=999"
            headers = {"Authorization": f"Bearer {self.access_token}"}
            
            all_items = []
            while url:
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    all_items.extend(data.get("value", []))
                    url = data.get("@odata.nextLink")
                else:
                    print(f"❌ Error fetching details items {source}: {response.text}")
                    break
            
            clean_data = []
            for item in all_items:
                fields = item.get("fields", {})
                
                # MAPPING LOGIC
                if source == "IPS":
                    # OLD LOGIC (field_x)
                    row = {
                        "ID": item.get("id"),
                        "Modified": item.get("lastModifiedDateTime", ""),
                        "Empresa": fields.get("Title", ""),
                        "Patronal_IPS": fields.get("field_1", ""),
                        "Patronal_REOP": fields.get("field_2", ""),
                        "Cedula": fields.get("field_3", ""),
                        "Nombre": fields.get("field_4", ""),
                        "Estado_IPS": fields.get("field_5", ""), 
                        "Estado_MTESS": fields.get("field_6", ""),
                        "AUD_ESTADO": fields.get("field_7", ""),
                        "Entrada_IPS": fields.get("field_8", ""),
                        "Entrada_MTESS": fields.get("field_9", ""),
                        "AUD_ENTRADA": fields.get("field_10", ""),
                        "Salida_IPS": fields.get("field_11", ""),
                        "Salida_MTESS": fields.get("field_12", ""),
                        "AUD_SALIDA": fields.get("field_13", ""),
                        "REVISADO": fields.get("REVISADO", ""),
                        "RUC": fields.get("RUC", ""),
                        "ModificadoPor": fields.get("ModificadoPor", item.get("lastModifiedBy", {}).get("user", {}).get("displayName", ""))
                    }
                else: # MTESS
                    # NEW LOGIC (Direct Names)
                    row = {
                        "ID": item.get("id"),
                        "Modified": item.get("lastModifiedDateTime", ""),
                        "Empresa": fields.get("EMPRESA", ""), # Title often aliases to user defined col if first? No, Title is Title.
                        # Wait, User said: "EMPRESA" is the column name. 
                        # In SharePoint 'Title' is usually mandatory. Creating a column 'EMPRESA' is separately.
                        # But if they renamed Title->EMPRESA, the internal name might still be Title.
                        # User said: "hay una pequeña diferencia que es en el nombre interno... EMPRESA, PATRONAL_IPS..."
                        # So I trust these are internal names.
                        "Empresa": fields.get("EMPRESA", fields.get("Title", "")), # Fallback Title just in case
                        "Patronal_IPS": fields.get("PATRONAL_IPS", ""),
                        "Patronal_REOP": fields.get("PATRONAL_REOP", ""),
                        "Cedula": fields.get("CEDULA", ""),
                        "Nombre": fields.get("NOMBRE", ""),
                        "Estado_IPS": fields.get("ESTADO_IPS", ""), 
                        "Estado_MTESS": fields.get("ESTADO_MTESS", ""),
                        "AUD_ESTADO": fields.get("AUD_ESTADO", ""),
                        "Entrada_IPS": fields.get("ENTRADA_IPS", ""),
                        "Entrada_MTESS": fields.get("ENTRADA_MTESS", ""),
                        "AUD_ENTRADA": fields.get("AUD_ENTRADA", ""),
                        "Salida_IPS": fields.get("SALIDA_IPS", ""),
                        "Salida_MTESS": fields.get("SALIDA_MTESS", ""),
                        "AUD_SALIDA": fields.get("AUD_SALIDA", ""),
                        "REVISADO": fields.get("REVISADO", ""),
                        "RUC": fields.get("RUC", ""),
                        "ModificadoPor": fields.get("ModificadoPor", item.get("lastModifiedBy", {}).get("user", {}).get("displayName", ""))
                    }

                clean_data.append(row)

            # GUARD: Race Condition Check
            if (time.time() - self.last_local_update_times[source]) < 5:
                print(f"⚠️ Discarding fetch {source}: User action is more recent.")
                return self.dfs[source]

            new_df = pd.DataFrame(clean_data)
            
            # Convert Modified
            if "Modified" in new_df.columns:
                new_df["Modified"] = pd.to_datetime(new_df["Modified"], errors="coerce")

            # Add Index
            if not new_df.empty:
                new_df.insert(0, "#", range(1, len(new_df) + 1))
            else:
                 new_df["#"] = []
                 # Ensure cols
                 cols = ["Empresa", "Patronal_IPS", "Patronal_REOP", "Cedula", "Nombre", 
                           "Estado_IPS", "Estado_MTESS", "AUD_ESTADO", 
                           "Entrada_IPS", "Entrada_MTESS", "AUD_ENTRADA",
                           "Salida_IPS", "Salida_MTESS", "AUD_SALIDA", "REVISADO", "ID", "ModificadoPor"]
                 for c in cols:
                     if c not in new_df.columns: new_df[c] = []

            self.dfs[source] = new_df
            
            self.last_change_times[source] = time.time()
            self._save_state(self.last_change_times)
            
            return self.dfs[source]
            
            return self.df
        except Exception as e:
            print(f"Error fetching data: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

    def get_inconsistencies(self, source="IPS"):
        df = self.dfs.get(source, pd.DataFrame())
        if df.empty: return df
        mask = (
            (self.df["AUD_ENTRADA"] != "COINCIDE") | 
            (self.df["AUD_ESTADO"] != "COINCIDE") | 
            (self.df["AUD_SALIDA"] != "COINCIDE")
        )
        return self.df[mask]

    def get_verified(self, source="IPS"):
        df = self.dfs.get(source, pd.DataFrame())
        if df.empty: return df
        # Check against "REVISADO" string
        mask = self.df["REVISADO"].astype(str).str.upper() == "REVISADO"
        return self.df[mask]

    def update_status_by_ids(self, ids, status, modifier=None, source="IPS"):
        """Updates 'REVISADO', 'Modified', and version timestamp."""
        try:
            df = self.dfs.get(source)
            if df is None: return

            # FIX: Ensure IDs are compared as Strings (Graph API IDs are strings)
            # Incoming 'ids' might be Integers from JSON, DF 'ID' is String.
            str_ids = [str(x) for x in ids]
            
            # Cast match column to string to be safe
            mask = df["ID"].astype(str).isin(str_ids)
            
            if not mask.any(): 
                print(f"⚠️ Warning: update_status_by_ids ({source}) matched 0 rows. IDs: {ids}")
                return

            self.dfs[source].loc[mask, "REVISADO"] = status
            # Update 'Modified' so the UI text updates too
            now_ts = pd.Timestamp.now(tz='UTC') 
            self.dfs[source].loc[mask, "Modified"] = now_ts
            
            # Update 'ModificadoPor' if provided
            if modifier:
                 self.dfs[source].loc[mask, "ModificadoPor"] = modifier
            
            self.last_change_times[source] = time.time()
            self.last_local_update_times[source] = self.last_change_times[source] # DEBOUNCE
            self._save_state(self.last_change_times) # Broadcast
            
        except Exception as e:
            print(f"Error updating local DF: {e}")

    def patch_sharepoint_background(self, item_ids, status="REVISADO", modifier_name=None, source="IPS"):
        thread = threading.Thread(target=self._patch_worker, args=(item_ids, status, modifier_name, source))
        thread.daemon = True
        thread.start()

    def _patch_worker(self, item_ids, status, modifier_name=None, source="IPS"):
        # Refresh Token for long-running processes
        self._authenticate()
        
        target_list_id = self.list_ids.get(source)

        if not self.access_token or not self.site_id or not target_list_id:
            print(f"❌ Cannot patch {source}: Missing connection details.")
            return

        # No lookup needed for custom string column "ModificadoPor"
        # We just write the name directly.

        print(f"Starting background patch ({source}) for {len(item_ids)} items to '{status}'...")
        url_base = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{target_list_id}/items"
        
        # NOTE: For Choice columns, Graph API expects the value as a string.
        # If the column is strictly defined, the value MUST match one of the choices exactly.
        # Ensure 'REVISADO' and 'PENDIENTE' are the exact values.
        
        headers = {
            "Authorization": f"Bearer {self.access_token}", 
            "Content-Type": "application/json"
        }
        
        for item_id in item_ids:
            try:
                # 1. Update STATUS first (CRITICAL)
                url = f"{url_base}/{item_id}/fields"
                body_status = {"REVISADO": status}
                
                resp_status = requests.patch(url, headers=headers, json=body_status)
                
                if resp_status.status_code == 200:
                    print(f"✅ Item {item_id} STATUS updated to {status}.")
                    
                if resp_status.status_code == 200:
                    print(f"✅ Item {item_id} STATUS updated to {status}.")
                    
                    # 2. Try updating Entity Custom Column if Name provided
                    if modifier_name:
                        try:
                             # User created custom column "ModificadoPor" (Text)
                             body_mod = {"ModificadoPor": modifier_name}
                             resp_mod = requests.patch(url, headers=headers, json=body_mod)
                             if resp_mod.status_code == 200:
                                 print(f"✅ Item {item_id} ModificadoPor updated to {modifier_name}.")
                             else:
                                 print(f"⚠️ item {item_id} ModificadoPor update failed: {resp_mod.text}")
                        except Exception as ex:
                             print(f"⚠️ ModificadoPor Patch Exception: {ex}")
                else:
                    print(f"❌ Failed to update {item_id}. Status: {resp_status.status_code}")
                    print(f"   Response: {resp_status.text}") 

            except Exception as e:
                print(f"Exception updating {item_id}: {e}")
        
        print("Background patch finished.")

    def _get_site_user_id(self, email):
        """Look up Site User ID (Integer) by email for Person fields."""
        if not self.access_token or not self.site_id: return None
        try:
             # Strategy: Query the hidden 'User Information List' to get the Integer ID.
             # This ID is required for 'EditorLookupId' or any Person 'LookupId'.
             # URL: /sites/{id}/lists('User Information List')/items?expand=fields&$filter=fields/EMail eq '{email}'
             
             # Note: 'User Information List' is the standard display name. 
             # Graph allows referring to lists by Display Name in some endpoints, 
             # but standard is /lists/{list-id}.
             # However, we can use the 'numerous' approach:
             # 1. Get List ID of 'User Information List'?
             # 2. Query items.
             
             # Let's try to find 'User Information List' ID first.
             # It might be cached or fetched.
             
             user_list_id = None
             url_lists = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists"
             headers = {"Authorization": f"Bearer {self.access_token}"}
             
             # Simple iteration to find it (it's system, might be hidden).
             # We might need to filter `?filter=displayName eq 'User Information List'`
             res = requests.get(f"{url_lists}?$filter=displayName eq 'User Information List'", headers=headers)
             if res.status_code == 200:
                 val = res.json().get('value', [])
                 if val:
                     user_list_id = val[0]['id']
             
             if not user_list_id:
                 # Fallback: Spanish name? 'Lista de información del usuario'?
                 # Or just try query all and find?
                 # Assuming English 'User Information List' is standard even in regions if created basic.
                 # If fails, we can't patch.
                 print("Could not find 'User Information List'.")
                 return None
                 
             # Query the item for the email
             url_items = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{user_list_id}/items"
             # Filter by EMail field in fields
             url_query = f"{url_items}?expand=fields&$filter=fields/EMail eq '{email}'"
             
             res = requests.get(url_query, headers=headers)
             if res.status_code == 200:
                 items = res.json().get('value', [])
                 if items:
                     # The ID of the item in User Info List IS the Lookup ID
                     return items[0]['id'] # Returns integer string e.g. "15"
                     
             print(f"User {email} not found in User Info List.")
             return None

        except Exception as e:
             print(f"Error lookup user: {e}")
             return None

    def trigger_rpa_background(self):
        if not self.webhook_url: return
        threading.Thread(target=self._rpa_worker, daemon=True).start()

    def _rpa_worker(self):
        try:
            requests.post(self.webhook_url, json={"action": "start_audit"})
            print("RPA Triggered.")
        except Exception as e:
            print(f"Error triggering RPA: {e}")

    # --- Whitelist Management ---

    def get_whitelist(self):
        """Fetches all emails from whitelist_ips_mtess."""
        if not self.access_token or not self.site_id: self._authenticate(); self._get_site_id(); self._get_list_ids()
        
        target_list_id = self.list_ids.get("WHITELIST")
        if not target_list_id:
             # Try to find it dynamically if not loaded
             self._get_list_ids()
             target_list_id = self.list_ids.get("WHITELIST")
             if not target_list_id:
                 print("❌ Whitelist list not found.")
                 return []

        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{target_list_id}/items?expand=fields"
            headers = {"Authorization": f"Bearer {self.access_token}"}
            res = requests.get(url, headers=headers)
            if res.status_code == 200:
                items = res.json().get("value", [])
                emails = []
                for item in items:
                    email = item.get("fields", {}).get("correo", "")
                    if email: emails.append({"id": item["id"], "email": email.lower()})
                return emails
            else:
                print(f"Error fetching whitelist: {res.text}")
                return []
        except Exception as e:
            print(f"Error getting whitelist: {e}")
            return []

    def add_to_whitelist(self, email):
        """Adds an email to whitelist_ips_mtess."""
        if not self.access_token or not self.site_id: self._authenticate()
        
        target_list_id = self.list_ids.get("WHITELIST")
        if not target_list_id: return False, "Lista no encontrada"

        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{target_list_id}/items"
            headers = {"Authorization": f"Bearer {self.access_token}", "Content-Type": "application/json"}
            
            body = {"fields": {"correo": email, "Title": email}} # Title is usually required
            res = requests.post(url, headers=headers, json=body)
            
            if res.status_code == 201:
                return True, "Agregado correctamente"
            else:
                return False, f"Error SharePoint: {res.text}"
        except Exception as e:
            return False, str(e)

    def remove_from_whitelist(self, item_id):
        """Removes an item from whitelist_ips_mtess by ID."""
        if not self.access_token or not self.site_id: self._authenticate()
        
        target_list_id = self.list_ids.get("WHITELIST")
        if not target_list_id: return False, "Lista no encontrada"

        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{target_list_id}/items/{item_id}"
            headers = {"Authorization": f"Bearer {self.access_token}"}
            res = requests.delete(url, headers=headers)
            
            if res.status_code == 204:
                return True, "Eliminado correctamente"
            else:
                return False, f"Error SharePoint: {res.text}"
        except Exception as e:
            return False, str(e)
