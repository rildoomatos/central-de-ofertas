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

# ===== FUNÇÕES =====

def moeda(valor):
    try:
        valor = float(valor)
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)


def gerar_legenda(produto):

    titulo = str(produto["title"])
    preco_atual = moeda(produto["sale_price"])
    preco_antigo = moeda(produto["price"])

    try:
        desconto = int(float(produto["discount_percentage"]))
    except:
        desconto = produto["discount_percentage"]

    try:
        avaliacao = float(produto["item_rating"])
        avaliacao = f"{avaliacao:.1f}".replace(".", ",")
    except:
        avaliacao = produto["item_rating"]

    legenda = (
        "🔥 *OFERTA NA SHOPEE!*\n\n"
        f"🛍️ {titulo}\n\n"
        f"❌ De: {preco_antigo}\n"
        f"✅ Por: *{preco_atual}*\n"
        f"🔥 {desconto}% OFF\n"
        f"⭐ Avaliação: {avaliacao}\n\n"
        "🛒 *Compre aqui:*\n"
        "[COLE O LINK DE AFILIADO]\n\n"
        "⚠️ Preço e disponibilidade podem mudar a qualquer momento!"
    )

    return legenda


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

# ===== ORDENAR MELHORES OFERTAS =====

df = df.sort_values(
    by=["discount_percentage", "item_rating"],
    ascending=[False, False]
)

# ===== EXCLUIR PRODUTOS JÁ EXISTENTES =====

ids_existentes = aba.col_values(1)

ids_existentes = set(
    str(x).strip()
    for x in ids_existentes
)

df["itemid"] = df["itemid"].astype(str)

df = df[
    ~df["itemid"].isin(ids_existentes)
]

# ===== ENVIAR PARA PLANILHA =====

linhas = []

for _, produto in df.iterrows():

    legenda = gerar_legenda(produto)

    linhas.append([
        produto["itemid"],                 # ID
        "Shopee",                          # Marketplace
        produto["title"],                  # Produto
        produto["sale_price"],             # Preço Atual
        produto["price"],                  # Preço Anterior
        produto["discount_percentage"],    # Desconto
        produto["item_rating"],            # Avaliação
        "",                                # Vendas
        "",                                # Categoria
        produto["product_link"],           # Link Original
        "",                                # Link Afiliado
        legenda,                           # Legenda
        "AGUARDANDO LINK",                 # Status
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        )
    ])

if linhas:
    aba.append_rows(
        linhas,
        value_input_option="USER_ENTERED"
    )

print(f"{len(linhas)} produtos enviados.")
