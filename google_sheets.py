"""
Integração com Google Sheets (leitura) e Google Drive (upload).
- Render/servidor: usa Service Account via variável GOOGLE_SERVICE_ACCOUNT_JSON
- Local: OAuth com browser (salva token.json)
"""
import io
import os
import json

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.file",
]

DRIVE_OUTPUT_FOLDER = "1FJd0MyZSXiD6gE70l7LVkNfuTmg7KEvy"

BASE_DIR = os.path.dirname(__file__)
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")


def _get_credentials():
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json:
        from google.oauth2.service_account import Credentials
        return Credentials.from_service_account_info(json.loads(sa_json), scopes=SCOPES)

    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=8090)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return creds


def read_sheet(sheet_id: str, aba: str = "Acompanhamento Geral") -> list:
    from googleapiclient.discovery import build
    creds = _get_credentials()
    service = build("sheets", "v4", credentials=creds)
    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=sheet_id,
            range=f"'{aba}'",
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )
    return result.get("values", [])


def upload_to_drive(buf: io.BytesIO, filename: str, folder_id: str = DRIVE_OUTPUT_FOLDER):
    """Faz upload de um buffer PPTX para o Google Drive."""
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload

    creds = _get_credentials()
    service = build("drive", "v3", credentials=creds)

    buf.seek(0)
    media = MediaIoBaseUpload(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        resumable=False,
    )

    file_metadata = {
        "name": filename,
        "parents": [folder_id],
    }

    service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True,
    ).execute()
