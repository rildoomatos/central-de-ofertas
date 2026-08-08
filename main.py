import json
import os
import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials

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


def numero(valor):
    try:
        return round(float(str(valor).replace(",", ".")), 2)
    except:
        return 0


def gerar_legenda(produto, link_afiliado=""):

    preco_atual = moeda(produto["sale_price"])
    preco_antigo = moeda(produto["price"])

    desconto = int(float(produto["discount_percentage"]))

    avaliacao = (
        f"{float(produto['item_rating']):.1f}"
        .replace(".", ",")
    )

    link = (
        link_afiliado
        if link_afiliado
        else "[COLE O LINK DE AFILIADO]"
    )

    return (
        "🔥 *OFERTA NA SHOPEE!*\n\n"
        f"🛍️ {produto['title']}\n\n"
        f"❌ De: {preco_antigo}\n"
        f"✅ Por: *{preco_atual}*\n"
        f"🔥 {desconto}% OFF\n"
        f"⭐ Avaliação: {avaliacao}\n\n"
        "🛒 *Compre aqui:*\n"
        f"{link}\n\n"
        "⚠️ Preço e disponibilidade podem mudar a qualquer momento!"
    )


# ===== BAIXAR FEED =====

arquivo = "feed.csv"

r = requests.get(CSV_URL, timeout=60)
r.raise_for_status()

with open(arquivo, "wb") as f:
    f.write(r.content)

df = pd.read_csv(arquivo)

df["item_rating"] = pd.to_numeric(
    df["item_rating"],
    errors="coerce"
)

df["discount_percentage"] = pd.to_numeric(
    df["discount_percentage"],
    errors="coerce"
)

df["sale_price"] = pd.to_numeric(
    df["sale_price"],
    errors="coerce"
)

df["price"] = pd.to_numeric(
    df["price"],
    errors="coerce"
)

df["itemid"] = df["itemid"].astype(str)

df = df.drop_duplicates(subset=["itemid"])


# ===== CONFIG =====

try:
    config = planilha.worksheet("CONFIG")

except gspread.WorksheetNotFound:

    config = planilha.add_worksheet(
        title="CONFIG",
        rows=200,
        cols=2
    )

    config.update(
        range_name="A1:B3",
        values=[
            ["CONFIGURAÇÃO", "VALOR"],
            ["Categoria", "TODAS"],
            ["Categoria anterior", ""]
        ]
    )


categoria_atual = config.acell("B2").value or "TODAS"
categoria_atual = categoria_atual.strip()

categoria_anterior = config.acell("B3").value or ""
categoria_anterior = categoria_anterior.strip()


# Se ainda não existe categoria anterior,
# tenta descobrir pela planilha atual
if not categoria_anterior:

    dados_atuais = aba.get_all_values()

    if len(dados_atuais) > 1 and len(dados_atuais[1]) >= 9:
        categoria_anterior = dados_atuais[1][8].strip()


# ===== SE TROCOU CATEGORIA, LIMPAR OFERTAS =====

trocou_categoria = (
    categoria_anterior
    and categoria_atual.lower() != categoria_anterior.lower()
)

if trocou_categoria:

    if aba.row_count > 1:
        aba.batch_clear([
            f"A2:N{aba.row_count}"
        ])

    print(
        f"Categoria alterada: "
        f"{categoria_anterior} → {categoria_atual}"
    )


config.update(
    range_name="B3",
    values=[[categoria_atual]]
)


# ===== FILTRAR CATEGORIA =====

df_categoria = df.copy()

if categoria_atual.upper() != "TODAS":

    df_categoria = df_categoria[
        df_categoria["global_category1"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        categoria_atual.lower()
    ]


# ===== LER PRODUTOS QUE JÁ ESTÃO NA PLANILHA =====

dados = aba.get_all_values()

produtos_existentes = {}

for numero_linha, linha in enumerate(dados[1:], start=2):

    if not linha:
        continue

    item_id = str(linha[0]).strip()

    if item_id:
        produtos_existentes[item_id] = {
            "linha": numero_linha,
            "dados": linha
        }


# ===== ATUALIZAR PRODUTOS COM PREÇO ALTERADO =====

atualizacoes = []

feed_por_id = {
    str(produto["itemid"]): produto
    for _, produto in df_categoria.iterrows()
}

for item_id, existente in produtos_existentes.items():

    if item_id not in feed_por_id:
        continue

    produto = feed_por_id[item_id]

    linha_antiga = existente["dados"]

    preco_antigo_planilha = (
        numero(linha_antiga[3])
        if len(linha_antiga) > 3
        else 0
    )

    preco_novo = numero(
        produto["sale_price"]
    )

    # Se preço não mudou, não faz nada
    if preco_antigo_planilha == preco_novo:
        continue

    link_afiliado = (
        linha_antiga[10].strip()
        if len(linha_antiga) > 10
        else ""
    )

    status = (
        linha_antiga[12].strip()
        if len(linha_antiga) > 12
        else "AGUARDANDO LINK"
    )

    legenda = gerar_legenda(
        produto,
        link_afiliado
    )

    nova_linha = [[
        item_id,
        "Shopee",
        produto["title"],
        produto["sale_price"],
        produto["price"],
        produto["discount_percentage"],
        produto["item_rating"],
        "",
        produto["global_category1"],
        produto["product_link"],
        link_afiliado,
        legenda,
        status,
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        )
    ]]

    atualizacoes.append({
        "range":
            f"A{existente['linha']}:N{existente['linha']}",
        "values": nova_linha
    })


if atualizacoes:

    aba.batch_update(
        atualizacoes,
        value_input_option="USER_ENTERED"
    )


# ===== SELECIONAR NOVAS OFERTAS =====

novos = df_categoria[
    (df_categoria["item_rating"] >= 4.8)
    &
    (df_categoria["discount_percentage"] >= 10)
].copy()

ids_existentes = set(
    produtos_existentes.keys()
)

novos = novos[
    ~novos["itemid"].isin(ids_existentes)
]

novos = novos.sort_values(
    by=[
        "discount_percentage",
        "item_rating"
    ],
    ascending=[
        False,
        False
    ]
)

# Até 5 novas ofertas por execução
novos = novos.head(5)


# ===== ADICIONAR NOVOS PRODUTOS =====

linhas_novas = []

for _, produto in novos.iterrows():

    linhas_novas.append([
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
        gerar_legenda(produto),
        "AGUARDANDO LINK",
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        )
    ])


if linhas_novas:

    aba.append_rows(
        linhas_novas,
        value_input_option="USER_ENTERED"
    )


print(f"Categoria: {categoria_atual}")
print(f"Produtos atualizados: {len(atualizacoes)}")
print(f"Novos produtos adicionados: {len(linhas_novas)}")
