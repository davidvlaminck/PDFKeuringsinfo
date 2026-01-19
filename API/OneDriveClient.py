import logging
import json
import time
import requests
from pathlib import Path

from msal import ConfidentialClientApplication

logging.basicConfig(
    level=logging.INFO,  # Toon INFO en hoger
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S")


# ---------- Token persistence ----------
def load_credentials(path: Path) -> dict:
    with path.open("r") as f:
        return json.load(f)

class OneDriveClient:
    def __init__(
        self,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        user_email: str,
        scopes=None,
    ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.user_email = user_email
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scopes = scopes or ["https://graph.microsoft.com/.default"]

        try:
            self.app = ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=self.authority
            )
        except Exception as e:
            print(e.args[0])
            raise RuntimeError(f"Fout bij initialisatie van MSAL-app: {e}") from e



    def get_access_token(self) -> str:
        result = self.app.acquire_token_for_client(scopes=self.scopes)
        if "access_token" in result:
            return result["access_token"]
        else:
            raise RuntimeError(f"Token acquisition failed: {result.get('error_description')}")

    # ---------- Graph API calls ----------
    def list_root_files(self):
        """List files/folders in the root of the user's OneDrive."""
        access_token = self.get_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = f"https://graph.microsoft.com/v1.0/users/{self.user_email}/drive/root/children"
        resp = requests.get(url, headers=headers)

        if resp.status_code == 200:
            items = resp.json().get("value", [])
            print("📂 Bestanden in OneDrive-root:")
            for item in items:
                soort = "Map" if "folder" in item else "Bestand"
                print(f"- {item['name']} ({soort})")
        else:
            print(f"Fout {resp.status_code}: {resp.text}")

    def download_file_by_name(self, filename: str, save_path: Path):
        """Zoekt bestand in root en downloadt het lokaal."""
        access_token = self.get_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}

        # Zoek het bestand in de root
        list_url = f"https://graph.microsoft.com/v1.0/users/{self.user_email}/drive/root/children"
        resp = requests.get(list_url, headers=headers)
        if resp.status_code != 200:
            logging.error(f"Fout bij ophalen lijst: {resp.status_code} - {resp.text}")
            return

        items = resp.json().get("value", [])
        match = next((i for i in items if i["name"].lower() == filename.lower()), None)
        if not match:
            logging.warning(f"Bestand '{filename}' niet gevonden in root")
            return

        file_id = match["id"]
        download_url = f"https://graph.microsoft.com/v1.0/users/{self.user_email}/drive/items/{file_id}/content"
        download_resp = requests.get(download_url, headers=headers, stream=True)

        if download_resp.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in download_resp.iter_content(chunk_size=8192):
                    f.write(chunk)
            logging.info(f"✅ Bestand '{filename}' gedownload naar {save_path}")
        else:
            logging.error(f"Fout bij downloaden: {download_resp.status_code} - {download_resp.text}")

    def upload_file(self, local_path: Path, onedrive_path: str = None):
        """
        Uploadt een bestand naar OneDrive.
        :param local_path: Lokaal pad naar bestand
        :param onedrive_path: Pad + bestandsnaam in OneDrive (default = root + bestandsnaam)
        """
        access_token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }

        local_path = Path(local_path)
        if not local_path.exists():
            logging.error(f"Lokaal bestand '{local_path}' bestaat niet")
            return

        if not onedrive_path:
            onedrive_path = local_path.name  # zelfde bestandsnaam

        url = f"https://graph.microsoft.com/v1.0/users/{self.user_email}/drive/root:/{onedrive_path}:/content"
        logging.info(f"⏫ Uploaden van '{local_path}' naar '{onedrive_path}' in OneDrive...")

        with open(local_path, "rb") as f:
            resp = requests.put(url, headers=headers, data=f)

        if resp.status_code in {200, 201}:
            logging.info(f"✅ Bestand succesvol geüpload als '{onedrive_path}'")
        else:
            logging.error(f"Fout bij upload: {resp.status_code} - {resp.text}")

# ---------- Gebruik ----------
if __name__ == "__main__":
    credentials = load_credentials(Path("cred_onedrive.json"))

    client = OneDriveClient(
        client_id=credentials["client_id"],
        client_secret=credentials["client_secret"],
        tenant_id=credentials["tenant_id"],
        user_email=credentials["user_email"]
    )
    client.list_root_files()

    exit()

    client.download_file_by_name('r_template.xlsx', Path('r_template.xlsx'))

    client.upload_file(Path('r_template.xlsx'))
