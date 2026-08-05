import json
import os

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

credenciais = json.loads(os.environ["GOOGLE_CREDENTIALS"])

credentials = Credentials.from_service_account_info(
    credenciais,
    scopes=SCOPES
)

gc = gspread.authorize(credentials)

planilha = gc.open_by_key(os.environ["GOOGLE_SHEET_ID"])

aba = planilha.worksheet("OFERTAS")

aba.append_row([
    1,
    "TESTE",
    "Automação funcionando",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "OK",
    ""
])

print("Conexão realizada com sucesso.")