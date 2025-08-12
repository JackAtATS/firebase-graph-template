import os
import json
import requests
import msal
import dotenv
import time
from datetime import datetime, timedelta

ENV_FILE = os.path.join(os.path.dirname(__file__), "config.env")
dotenv.load_dotenv(ENV_FILE)

CLIENT_ID = os.getenv("MS_CLIENT_ID")
TENANT_ID = os.getenv("MS_TENANT_ID")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")  # optional; not used with /me/sendMail
CACHE_FILE = os.path.join(os.path.dirname(__file__), "msal_token_cache.bin")
MAX_RETRIES = 3 # Maximum number of retries for each API call

# Keep scopes in one place and use the same set for silent + interactive.
SCOPES = ["Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All", "Mail.Send"]

class GraphClient:
    def __init__(
        self,
        client_id=CLIENT_ID,
        tenant_id=TENANT_ID,
        cache_file=CACHE_FILE,
        sender_email=SENDER_EMAIL,
    ):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.cache_file = cache_file
        self.sender_email = sender_email
        self.cache = msal.SerializableTokenCache()
        self.max_retries = MAX_RETRIES


        # Load existing cache (if any)
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    content = f.read()
                    if content.strip():
                        self.cache.deserialize(content)
            except Exception as e:
                # Corrupt or unreadable cache shouldn't block auth
                print(f"⚠️ Could not read token cache: {e}")

        self.app = msal.PublicClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self.cache,
        )

        # Acquire a token (silent if possible, otherwise interactive)
        self._get_token_or_authenticate()

    def _save_cache_if_changed(self):
        if self.cache.has_state_changed:
            os.makedirs(os.path.dirname(self.cache_file), exist_ok=True)
            with open(self.cache_file, "w", encoding="utf-8") as f:
                f.write(self.cache.serialize())

    def _get_token_or_authenticate(self):
        accounts = self.app.get_accounts()
        token_response = None
        if accounts:
            token_response = self.app.acquire_token_silent(SCOPES, account=accounts[0])
        if not token_response or "access_token" not in token_response:
            token_response = self.app.acquire_token_interactive(scopes=SCOPES)

        if not token_response or "access_token" not in token_response:
            raise RuntimeError(
                f"Authentication failed: {token_response.get('error_description', 'No token returned') if token_response else 'No response'}"
            )

        self._save_cache_if_changed()
        self.token = token_response["access_token"]

    def _ensure_token(self):
        """
        Refresh the access token silently when possible.
        Call this before making Graph requests.
        """
        accounts = self.app.get_accounts()
        if not accounts:
            # No account loaded for some reason; fall back to full auth
            self._get_token_or_authenticate()
            return

        token_response = self.app.acquire_token_silent(SCOPES, account=accounts[0])
        if not token_response or "access_token" not in token_response:
            # Silent refresh failed—fall back to interactive
            token_response = self.app.acquire_token_interactive(scopes=SCOPES)
            if not token_response or "access_token" not in token_response:
                raise RuntimeError(
                    f"Authentication failed: {token_response.get('error_description', 'No token returned') if token_response else 'No response'}"
                )

        self._save_cache_if_changed()
        self.token = token_response["access_token"]

    def _auth_headers(self):
        self._ensure_token()
        return {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

    def get_excel_rows(self, excel_item_id, excel_worksheet) -> list:
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{excel_item_id}/workbook/worksheets/{excel_worksheet}/usedRange"
        response = requests.get(url, headers=headers)
        data = response.json() 
        if response.status_code == 429 and self.max_retries > 0:
            self.max_retries -= 1
            retry_after = int(response.headers.get('Retry-After', '30'))
            print(f"Throttled. Retrying after {retry_after} seconds...")
            time.sleep(retry_after)
            return self.get_excel_rows(excel_item_id, excel_worksheet)

        if response.status_code != 200:
            raise Exception(f"Failed to get Excel rows: {response.text}")
        rows = (data['values'])
        return rows
    
    def batch_update_excel_rows(self, excel_item_id, excel_worksheet, rows, end_column):
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        request_list = []
        for id, row in enumerate(rows):
            row_number = row[0]
            row_data = row[1]
            url = f"/me/drive/items/{excel_item_id}/workbook/worksheets('{excel_worksheet}')/range(address='A{row_number}:{end_column}{row_number}')"
            request_list.append({
                "id" : str(id + 1),
                "method": "PATCH",
                "url" : url,
                "body": {
                    "values": [row_data]
                },
                "headers" : {
                    "Content-Type": "application/json"
                }
            })
        url = f"https://graph.microsoft.com/v1.0/$batch"
        body = {
            "requests": request_list
        }
        response = requests.post(url, headers=headers, json=body)
        if response.status_code == 429 and self.max_retries > 0:
            self.max_retries -= 1
            retry_after = int(response.headers.get('Retry-After', '30'))
            print(f"Throttled. Retrying after {retry_after} seconds...")
            time.sleep(retry_after)
            return self.batch_update_excel_rows(excel_item_id, excel_worksheet, rows, end_column)
        if response.status_code != 200:
            raise Exception(f"Failed to batch update Excel: {response.text}")
        self.max_retries = MAX_RETRIES
        if "EditModeCannotAcquireLockTooManyRequests" in response.text and self.max_retries > 0:
            self.max_retries -= 1
            print("EditModeCannotAcquireLockTooManyRequests")
            time.sleep(10)
            return self.batch_update_excel_rows(excel_item_id, excel_worksheet, rows, end_column)
        self.max_retries = MAX_RETRIES
        return response.json()
    
    def reorder_excel_rows(self, excel_item_id, excel_worksheet, fields):
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{excel_item_id}/workbook/worksheets/{excel_worksheet}/usedRange/sort/apply"
        body = {
            "fields": fields
        }
        response = requests.post(url, headers=headers, json=body)
        if response.status_code == 429 and self.max_retries > 0:
            self.max_retries -= 1
            retry_after = int(response.headers.get('Retry-After', '30'))
            print(f"Throttled. Retrying after {retry_after} seconds...")
            time.sleep(retry_after)
            return self.reorder_excel_rows(excel_item_id, excel_worksheet, fields)
        if response.status_code != 200:
            raise Exception(f"Failed to reorder Excel rows: {response.text}")
        self.max_retries = MAX_RETRIES
        return response.json()
        
    def append_rows_to_table(self, excel_item_id, table_name, rows):
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{excel_item_id}/workbook/tables/{table_name}/rows/add"
        body = {
            "values": rows
        }
        response = requests.post(url, headers=headers, json=body)
        if response.status_code == 429 and self.max_retries > 0:
            self.max_retries -= 1
            retry_after = int(response.headers.get('Retry-After', '30'))
            print(f"Throttled. Retrying after {retry_after} seconds...")
            time.sleep(retry_after)
            return self.append_rows_to_table(excel_item_id, table_name, rows)
        if "EditModeCannotAcquireLockTooManyRequests" in response.text and self.max_retries > 0:
            self.max_retries -= 1
            print("EditModeCannotAcquireLockTooManyRequests")
            time.sleep(10)
            return self.append_rows_to_table(excel_item_id, table_name, rows)
        if response.status_code != 201:
            raise Exception(f"Failed to append rows: {response.text}")
        self.max_retries = MAX_RETRIES
        return response.json()    
    
    def list_tables(self, excel_item_id):
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{excel_item_id}/workbook/tables"
        response = requests.get(url, headers=headers)
        if response.status_code == 429 and self.max_retries > 0:
            self.max_retries -= 1
            retry_after = int(response.headers.get('Retry-After', '30'))
            print(f"Throttled. Retrying after {retry_after} seconds...")
            time.sleep(retry_after)
            return self.list_tables(excel_item_id)
        if response.status_code != 200:
            raise Exception(f"Failed to list tables: {response.text}")
        self.max_retries = MAX_RETRIES
        return response.json().get('value', [])
    
    def send_email(self, recipient_email: str, subject: str, body: str, save_to_sent=True):
        """
        Sends an email as the signed-in user.
        Note: do NOT set 'from' when using /me/sendMail.
        """
        headers = self._auth_headers()
        url = "https://graph.microsoft.com/v1.0/me/sendMail"
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": recipient_email}}],
            },
            "saveToSentItems": bool(save_to_sent),
        }
        resp = requests.post(url, headers=headers, json=payload)
        if resp.status_code not in (202, 200):
            raise Exception(f"Failed to send email: {resp.status_code} {resp.text}")
        return True