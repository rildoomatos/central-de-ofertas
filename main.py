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
        return (
            f"R$ {valor:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except:
        return str(valor)


def gerar_legenda(produto):

    titulo = str(produto["title"])
    preco_atual = moeda(produto["sale_price"])
    preco_antigo = moeda(produto["price"])

    desconto = int(float(produto["discount_percentage"]))

    avaliacao = float(produto["item_rating"])
    avaliacao = f"{avaliacao:.1f}".replace(".", ",")

    return (
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


# ===== ATUALIZAR LINKS DE AFILIADO =====

dados_planilha = aba.get_all_values()
atualizacoes = []

for numero_linha, linha in enumerate(dados_planilha[1:], start=2):

    if len(linha) < 13:
        continue

    link_afiliado = linha[10].strip()
    legenda = linha[11]
    status = linha[12]

    if link_afiliado and status != "PRONTO":

        legenda_final = legenda.replace(
            "[COLE O LINK DE AFILIADO]",
            link_afiliado
        )

        atualizacoes.append({
            "range": f"L{numero_linha}:M{numero_linha}",
            "values": [[legenda_final, "PRONTO"]]
        })

if atualizacoes:
    aba.batch_update(
        atualizacoes,
        value_input_option="USER_ENTERED"
    )


# ===== BAIXAR CSV =====

arquivo = "feed.csv"

r = requests.get(CSV_URL, timeout=60)
r.raise_for_status()

with open(arquivo, "wb") as f:
    f.write(r.content)


# ===== LER CSV =====

df = pd.read_csv(arquivo)

df["item_rating"] = pd.to_numeric(
    df["item_rating"],
    errors="coerce"
)

df["discount_percentage"] = pd.to_numeric(
    df["discount_percentage"],
    errors="coerce"
)


# ===== CRIAR / ATUALIZAR ABA CONFIG =====

try:
    config = planilha.worksheet("CONFIG")

except gspread.WorksheetNotFound:

    config = planilha.add_worksheet(
        title="CONFIG",
        rows=200,
        cols=2
    )

    config.update(
        range_name="A1:B2",
        values=[
            ["CONFIGURAÇÃO", "VALOR"],
            ["Categoria", "TODAS"]
        ]
    )


categorias = sorted(
    df["global_category1"]
    .dropna()
    .astype(str)
    .unique()
)

config.update(
    range_name="A4:A4",
    values=[["CATEGORIAS DISPONÍVEIS"]]
)

if categorias:
    config.update(
        range_name=f"A5:A{4 + len(categorias)}",
        values=[[categoria] for categoria in categorias]
    )


# ===== CATEGORIA ESCOLHIDA =====

categoria_escolhida = config.acell("B2").value

if not categoria_escolhida:
    categoria_escolhida = "TODAS"

categoria_escolhida = categoria_escolhida.strip()


# ===== FILTROS =====

df = df[df["item_rating"] >= 4.8]

df = df[
    df["discount_percentage"] >= 10
]

df = df.drop_duplicates(
    subset=["itemid"]
)


# ===== FILTRAR CATEGORIA =====

if categoria_escolhida.upper() != "TODAS":

    df = df[
        df["global_category1"]
        .astype(str)
        .str.lower()
        ==
        categoria_escolhida.lower()
    ]


# ===== EXCLUIR PRODUTOS JÁ EXISTENTES =====

ids_existentes = set(
    str(x).strip()
    for x in aba.col_values(1)
)

df["itemid"] = df["itemid"].astype(str)

df = df[
    ~df["itemid"].isin(ids_existentes)
]


# ===== ORDENAR MELHORES OFERTAS =====

df = df.sort_values(
    by=[
        "discount_percentage",
        "item_rating"
    ],
    ascending=[
        False,
        False
    ]
)


# ===== SOMENTE AS 5 MELHORES =====

df = df.head(5)


# ===== ENVIAR PARA PLANILHA =====

linhas = []

for _, produto in df.iterrows():

    legenda = gerar_legenda(produto)

    linhas.append([
        produto["itemid"],
        "Shopee",
        produto["title"],
        produto["sale_price"],
        produto["price"],
        produto["discount_percentage"],
        produto["item_rating"],
        "",
        produto["global_category1"],
        produto["product_link"],
        "",
        legenda,
        "AGUARDANDO LINK",
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        )
    ])


if linhas:
    aba.append_rows(
        linhas,
        value_input_option="USER_ENTERED"
    )


print(
    f"Categoria: {categoria_escolhida}"
)

print(
    f"{len(linhas)} ofertas selecionadas."
)

print(
    f"{len(atualizacoes)} ofertas ficaram PRONTAS."
)
