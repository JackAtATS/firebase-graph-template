#!/usr/bin/env python3
"""
Local authentication script to generate MSAL token cache.
Run this locally to authenticate and create the cache file.
"""

import os
import msal
import dotenv

# Load environment variables
import os
import msal
import dotenv

dotenv.load_dotenv("config.env")

CLIENT_ID = os.getenv('MS_CLIENT_ID')
TENANT_ID = os.getenv('MS_TENANT_ID')
CACHE_FILE = os.path.join(os.path.dirname(__file__), 'msal_token_cache.bin')

def authenticate():
    print(f"Authenticating with CLIENT_ID: {CLIENT_ID}")
    print(f"TENANT_ID: {TENANT_ID}")
    print(f"Cache file: {CACHE_FILE}")

    cache = msal.SerializableTokenCache()

    # ⬇️ Load existing cache from disk
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception as e:
            print(f"⚠️ Could not read token cache: {e}")

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f'https://login.microsoftonline.com/{TENANT_ID}',
        token_cache=cache
    )

    accounts = app.get_accounts()
    if accounts:
        print(f"Found {len(accounts)} account(s), trying to acquire token silently...")
        token_response = app.acquire_token_silent(
            ["Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All", "Mail.Send"],
            account=accounts[0]
        )
    else:
        print("No existing account found.")
        token_response = None

    if not token_response or 'access_token' not in token_response:
        print("No valid token found. Starting interactive authentication...")
        token_response = app.acquire_token_interactive(
            scopes=["Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All", "Mail.Send"]
        )

    if token_response and 'access_token' in token_response:
        print("✅ Authentication successful!")
        if cache.has_state_changed:
            os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
            with open(CACHE_FILE, "w", encoding="utf-8") as f:
                f.write(cache.serialize())
            print(f"✅ Token cache saved to: {CACHE_FILE}")
        else:
            print("ℹ️  Cache unchanged, no need to save.")
        return True

    print("❌ Authentication failed!")
    print(f"Error: {token_response.get('error_description', 'Unknown error') if token_response else 'No response'}")
    return False

if __name__ == "__main__":
    authenticate()