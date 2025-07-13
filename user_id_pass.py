import os
import io
import json
import random
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# === STEP 1: Download users.xlsx from Google Drive ===
def download_excel_from_drive(file_id, credentials_json):
    creds_dict = json.loads(credentials_json)
    creds = service_account.Credentials.from_service_account_info(creds_dict)
    service = build('drive', 'v3', credentials=creds)

    # Check file type
    file = service.files().get(fileId=file_id, fields='mimeType, name').execute()
    mime_type = file['mimeType']
    print(f"📄 Downloading file: {file['name']} ({mime_type})")

    # Download method based on type
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        request = service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        request = service.files().get_media(fileId=file_id)

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

    print("✅ Authenticated as:", creds.service_account_email)
    print("📁 Attempting to download file ID:", file_id)

    with open("users.xlsx", "wb") as f:
        f.write(fh.getbuffer())

# === DOWNLOAD: Triggered if credentials are found ===
if "GDRIVE_CREDENTIALS_JSON" in os.environ:
    # 🔧 REPLACE this with your actual FILE ID from Google Drive
    file_id = "1eXvMefOX2ps4voSiqf-YJlDVczzdFLxE"
    download_excel_from_drive(file_id, os.environ["GDRIVE_CREDENTIALS_JSON"])
else:
    raise ValueError("❌ Missing GDRIVE_CREDENTIALS_JSON environment variable.")


# === STEP 2: Generate user IDs and passwords from users.xlsx ===
excel_path = "users.xlsx"  # Downloaded file
sheet_name = "Employee_data"  # Sheet name inside Excel
output_path = "ID_pass.xlsx"  # Where credentials are saved

# Load latest employee data
df_new = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl')
df_new.columns = df_new.columns.str.strip()

# Create UniqueKey (based on email if available)
if 'Email Address' in df_new.columns:
    df_new["UniqueKey"] = df_new["Email Address"].str.strip().str.lower()
else:
    df_new["UniqueKey"] = (df_new["First Name"].str.strip().str.lower() + "_" +
                           df_new["Last Name"].str.strip().str.lower())

# Load existing ID/Pass file or create empty
if os.path.exists(output_path):
    df_existing = pd.read_excel(output_path, engine='openpyxl')
    df_existing.columns = df_existing.columns.str.strip()
    if "UniqueKey" not in df_existing.columns:
        if 'Email Address' in df_existing.columns:
            df_existing["UniqueKey"] = df_existing["Email Address"].str.strip().str.lower()
        else:
            df_existing["UniqueKey"] = (df_existing["First Name"].str.strip().str.lower() + "_" +
                                        df_existing["Last Name"].str.strip().str.lower())
else:
    df_existing = pd.DataFrame(columns=list(df_new.columns) + ["Username", "Password", "UniqueKey"])
    df_existing.index = pd.RangeIndex(start=0, stop=0)

# Check for duplicate UniqueKeys
if df_new["UniqueKey"].duplicated().any():
    raise ValueError("❌ Duplicate UniqueKeys in new employee data!")
if df_existing["UniqueKey"].duplicated().any():
    raise ValueError("❌ Duplicate UniqueKeys in existing ID/pass data!")

# Filter new users
new_users = df_new[~df_new["UniqueKey"].isin(df_existing["UniqueKey"])]

# Username and Password generators
def generate_password(first_name):
    special_chars = "!@#$"
    return "Smb" + first_name[:2].capitalize() + str(random.randint(0, 9)) + random.choice(special_chars)

def generate_username(first, last, index):
    return (first[:2] + last[:2]).lower() + f"{index:02}"

# Generate credentials
start_index = len(df_existing)
new_users = new_users.copy()
new_users["Username"] = [generate_username(row["First Name"], row["Last Name"], i + start_index)
                         for i, row in new_users.iterrows()]
new_users["Password"] = new_users["First Name"].apply(generate_password)

# Combine and save
df_existing.reset_index(drop=True, inplace=True)
new_users.reset_index(drop=True, inplace=True)
df_final = pd.concat([df_existing, new_users], ignore_index=True) if not df_existing.empty else new_users
df_final.drop(columns=["UniqueKey"], inplace=True)
df_final.to_excel(output_path, index=False)

print(f"✅ {len(new_users)} new user(s) added.")
print(f"✅ Usernames and passwords saved to: {output_path}")
