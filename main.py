import json
import os
import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials

# ===== CONFIG =====
CSV_URL = os.environ["SHOPEE_FEED_URL"]

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

# ===== BAIXAR CSV =====
arquivo = "feed.csv"

r = requests.get(CSV_URL, timeout=60)
r.raise_for_status()

with open(arquivo, "wb") as f:
    f.write(r.content)

# ===== LER CSV =====
df = pd.read_csv(arquivo)

# ===== FILTROS =====
df = df[df["item_rating"] >= 4.8]
df = df[df["discount_percentage"] >= 10]

df = df.drop_duplicates(subset=["itemid"])
# ===== EXCLUIR PRODUTOS JÁ EXISTENTES NA PLANILHA =====
ids_existentes = aba.col_values(1)

ids_existentes = set(str(x).strip() for x in ids_existentes)

df["itemid"] = df["itemid"].astype(str)

df = df[~df["itemid"].isin(ids_existentes)]

# ===== ENVIAR PARA PLANILHA =====
linhas = []

for _, produto in df.iterrows():
    linhas.append([
        produto["itemid"],
        produto["title"],
        produto["price"],
        produto["sale_price"],
        produto["discount_percentage"],
        produto["item_rating"],
        produto["product_link"],
        produto["image_link"]
    ])

if linhas:
    aba.append_rows(linhas, value_input_option="USER_ENTERED")

print(f"{len(linhas)} produtos enviados.")
