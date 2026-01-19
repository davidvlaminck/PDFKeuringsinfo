import json
import logging
import requests
from msal import ConfidentialClientApplication
import base64

class OneDriveClient:
    def __init__(self, client_id: str, client_secret: str, tenant_id: str, scopes=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scopes = scopes or ["https://graph.microsoft.com/.default"]
        self.app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )

    def get_access_token(self) -> str:
        result = self.app.acquire_token_for_client(scopes=self.scopes)
        if "access_token" in result:
            return result["access_token"]
        raise RuntimeError(
            f"Token acquisition failed: {result.get('error')} - {result.get('error_description')} "
            f"(correlation_id={result.get('correlation_id')})"
        )

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self.get_access_token()}"}


    def diagnose(self, test_user_upn: str | None = None):
        """
        - Acquires token and prints minimal MSAL result (sanitized)
        - Decodes JWT to show 'roles' claim (application permissions)
        - Lists drives visible to the app
        - Optionally probes a specific user's OneDrive
        """
        print("=== MSAL token acquisition ===")
        msal_result = self.app.acquire_token_for_client(scopes=self.scopes)
        print(json.dumps({
            k: v for k, v in msal_result.items()
            if k in ("token_type", "expires_in", "ext_expires_in", "correlation_id", "error", "error_description")
        }, indent=2))

        if "access_token" not in msal_result:
            print("❌ No access_token returned. See error fields above.")
            return

        access_token = msal_result["access_token"]

        # --- Decode JWT to inspect roles ---
        print("\n=== Decoded token roles ===")
        try:
            payload_b64 = access_token.split(".")[1] + "=="
            payload_json = base64.urlsafe_b64decode(payload_b64).decode("utf-8")
            payload = json.loads(payload_json)
            print("roles:", payload.get("roles", []))
        except Exception as e:
            print(f"⚠️ Could not decode token: {e}")

        headers = {"Authorization": f"Bearer {access_token}"}

        # --- List drives visible to the app ---
        print("\n=== Graph probe: /v1.0/drives ===")
        drives_resp = requests.get("https://graph.microsoft.com/v1.0/drives", headers=headers)
        print(f"HTTP {drives_resp.status_code}")
        try:
            drives_data = drives_resp.json()
            if drives_resp.status_code == 200:
                for d in drives_data.get("value", []):
                    print(f"📂 Drive: {d.get('name')} (id={d.get('id')}) type={d.get('driveType')}")
            else:
                print(json.dumps(drives_data, indent=2))
        except Exception:
            print(drives_resp.text)

        # --- Optional: probe a specific user's OneDrive ---
        if test_user_upn:
            print(f"\n=== Graph probe: /v1.0/users/{test_user_upn}/drive ===")
            drive_url = f"https://graph.microsoft.com/v1.0/users/{test_user_upn}/drive"
            drv_resp = requests.get(drive_url, headers=headers)
            print(f"HTTP {drv_resp.status_code}")
            try:
                print(json.dumps(drv_resp.json(), indent=2))
                if drv_resp.status_code in (401, 403):
                    print("⚠️ Likely missing admin consent for Files.ReadWrite.All (application) "
                          "or the user's OneDrive is not provisioned/licensed.")
            except Exception:
                print(drv_resp.text)


import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def load_credentials(path: Path) -> dict:
    import json
    with open(path, "r") as f:
        return json.load(f)

if __name__ == "__main__":
    credentials = load_credentials(Path("cred_onedrive.json"))
    client = OneDriveClient(
        client_id=credentials["client_id"],
        client_secret=credentials["client_secret"],
        tenant_id=credentials["tenant_id"],
    )

    # Optional: set a known user UPN to validate OneDrive access
    test_user_upn = "david.vlaminck@mow.vlaanderen.be" # e.g., "john.doe@yourtenant.com"
    client.diagnose(test_user_upn=test_user_upn)