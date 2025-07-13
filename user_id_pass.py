import os
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

def download_excel_from_drive(file_id, credentials_json):
    # Load credentials from GitHub secret (as JSON string)
    creds_dict = json.loads(credentials_json)
    creds = service_account.Credentials.from_service_account_info(creds_dict)
    service = build('drive', 'v3', credentials=creds)

    # Request and download file
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

    # Save as local Excel file
    with open("users.xlsx", "wb") as f:
        f.write(fh.getbuffer())

# === DOWNLOAD FROM GOOGLE DRIVE ===
if "GDRIVE_CREDENTIALS_JSON" in os.environ:
    file_id = "1eXvMefOX2ps4voSiqf-YJlDVczzdFLxE"  # e.g. "1AbCDeFgHiJkLmNoPqR"
    download_excel_from_drive(file_id, os.environ["GDRIVE_CREDENTIALS_JSON"])
else:
    raise ValueError("❌ Missing GDRIVE_CREDENTIALS_JSON environment variable.")
