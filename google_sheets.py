"""
Lê dados do Google Sheets.
- Render/servidor: usa Service Account via variável de ambiente GOOGLE_SERVICE_ACCOUNT_JSON
- Local (primeira vez): abre navegador para autenticação OAuth e salva token.json
"""
import os
import json

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
BASE_DIR = os.path.dirname(__file__)
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")


def _get_service():
    from googleapiclient.discovery import build

    # Service Account (Render / produção)
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json:
        from google.oauth2.service_account import Credentials
        info = json.loads(sa_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return build("sheets", "v4", credentials=creds)

    # OAuth local
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
            creds = flow.run_local_server(port=8080)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("sheets", "v4", credentials=creds)


def read_sheet(sheet_id: str, aba: str = "Acompanhamento Geral") -> list:
    service = _get_service()
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
