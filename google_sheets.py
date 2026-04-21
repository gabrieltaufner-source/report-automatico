"""
Lê dados diretamente do Google Sheets via API.
Na primeira execução abre o navegador para autenticação.
O token é salvo em token.json para reutilização.
"""
import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
BASE_DIR = os.path.dirname(__file__)
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")


def _get_service():
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


def read_sheet(sheet_id: str, aba: str = "Acompanhamento Geral") -> list[list]:
    """
    Retorna todas as linhas da aba como lista de listas (valores formatados).
    """
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
