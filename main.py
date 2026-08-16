import json
import os
import time
import hashlib
import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# CONFIGURAÇÕES
# =========================================================

APP_ID = os.environ["SHOPEE_APP_ID"]
SECRET = os.environ["SHOPEE_API_PASSWORD"]
CSV_URL = os.environ["SHOPEE_FEED_URL"]

API_URL = "https://open-api.affiliate.shopee.com.br/graphql"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

credenciais = json.loads(
    os.environ["GOOGLE_CREDENTIALS"]
)

credentials = Credentials.from_service_account_info(
    credenciais,
    scopes=SCOPES
)

gc = gspread.authorize(credentials)

planilha = gc.open_by_key(
    os.environ["GOOGLE_SHEET_ID"]
)

aba = planilha.worksheet("OFERTAS")


# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def numero(valor):
    try:
        return float(
            str(valor)
            .replace("R$", "")
            .replace(".", "")
            .replace(",", ".")
            .strip()
        )
    except:
        return 0


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


def calcular_preco_original(preco_atual, desconto):

    preco_atual = float(preco_atual)
    desconto = float(desconto)

    if desconto <= 0 or desconto >= 100:
        return preco_atual

    return round(
        preco_atual / (1 - desconto / 100),
        2
    )


# =========================================================
# API SHOPEE
# =========================================================

def chamar_api(query):

    payload = json.dumps(
        {"query": query},
        separators=(",", ":")
    )

    timestamp = str(
        int(time.time())
    )

    fator = (
        APP_ID
        + timestamp
        + payload
        + SECRET
    )

    signature = hashlib.sha256(
        fator.encode("utf-8")
    ).hexdigest()

    headers = {
        "Authorization": (
            f"SHA256 Credential={APP_ID},"
            f"Timestamp={timestamp},"
            f"Signature={signature}"
        ),
        "Content-Type": "application/json"
    }

    resposta = requests.post(
        API_URL,
        data=payload,
        headers=headers,
        timeout=60
    )

    resposta.raise_for_status()

    resultado = resposta.json()

    if resultado.get("errors"):
        raise Exception(
            f"Erro API Shopee: {resultado['errors']}"
        )

    return resultado.get(
        "data",
        {}
    )


# =========================================================
# BUSCAR PRODUTOS
# =========================================================

CAMPOS_PRODUTO = """
itemId
productName
sales
ratingStar
priceMin
priceMax
priceDiscountRate
commissionRate
commission
productLink
imageUrl
productCatIds
"""


def buscar_produtos(
    categoria_id=None,
    pagina=1,
    limite=50
):

    filtro_categoria = ""

    if categoria_id:
        filtro_categoria = (
            f"productCatId:{int(categoria_id)}"
        )

    query = f"""
    {{
      productOfferV2(
        listType:0
        {filtro_categoria}
        sortType:2
        page:{pagina}
        limit:{limite}
      ) {{
        nodes {{
          {CAMPOS_PRODUTO}
        }}
        pageInfo {{
          page
          limit
          hasNextPage
        }}
      }}
    }}
    """

    dados = chamar_api(query)

    return dados.get(
        "productOfferV2",
        {}
    )


def buscar_produto_por_id(item_id):

    query = f"""
    {{
      productOfferV2(
        itemId:{int(item_id)}
        page:1
        limit:1
      ) {{
        nodes {{
          {CAMPOS_PRODUTO}
        }}
      }}
    }}
    """

    dados = chamar_api(query)

    resultado = dados.get(
        "productOfferV2",
        {}
    )

    produtos = resultado.get(
        "nodes",
        []
    )

    if produtos:
        return produtos[0]

    return None


# =========================================================
# GERAR LINK DE AFILIADO AUTOMATICAMENTE
# =========================================================

def gerar_link_afiliado(url_original):

    url_segura = json.dumps(
        str(url_original)
    )

    query = f"""
    mutation {{
      generateShortLink(
        input: {{
          originUrl: {url_segura}
        }}
      ) {{
        shortLink
      }}
    }}
    """

    dados = chamar_api(query)

    resultado = dados.get(
        "generateShortLink",
        {}
    )

    return resultado.get(
        "shortLink",
        ""
    )


# =========================================================
# LEGENDA
# =========================================================

def gerar_legenda(produto, link_afiliado):

    preco_atual = float(
        produto.get("priceMin") or 0
    )

    desconto = float(
        produto.get("priceDiscountRate") or 0
    )

    preco_anterior = calcular_preco_original(
        preco_atual,
        desconto
    )

    avaliacao = float(
        produto.get("ratingStar") or 0
    )

    vendas = int(
        produto.get("sales") or 0
    )

    return (
        "🔥 *OFERTA NA SHOPEE!*\n\n"
        f"🛍️ {produto.get('productName', '')}\n\n"
        f"❌ De: {moeda(preco_anterior)}\n"
        f"✅ Por: *{moeda(preco_atual)}*\n"
        f"🔥 {int(desconto)}% OFF\n"
        f"⭐ Avaliação: {avaliacao:.1f}\n"
        f"🛒 +{vendas} vendidos\n\n"
        "👉 *Compre aqui:*\n"
        f"{link_afiliado}\n\n"
        "⚠️ Preço e disponibilidade podem mudar a qualquer momento!"
    )


# =========================================================
# CONFIG
# =========================================================

try:

    config = planilha.worksheet(
        "CONFIG"
    )

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


categoria_atual = (
    config.acell("B2").value
    or "TODAS"
).strip()

categoria_anterior = (
    config.acell("B3").value
    or ""
).strip()


# =========================================================
# FEED SOMENTE PARA CATEGORIAS
# =========================================================

feed = requests.get(
    CSV_URL,
    timeout=60
)

feed.raise_for_status()

with open(
    "categorias.csv",
    "wb"
) as arquivo:

    arquivo.write(
        feed.content
    )


df_categorias = pd.read_csv(
    "categorias.csv",
    usecols=[
        "global_category1",
        "global_catid1"
    ]
)

df_categorias = (
    df_categorias
    .dropna()
    .drop_duplicates()
)


mapa_categorias = {}

for _, linha in df_categorias.iterrows():

    nome = str(
        linha["global_category1"]
    ).strip()

    categoria_id = int(
        linha["global_catid1"]
    )

    if nome not in mapa_categorias:
        mapa_categorias[nome] = categoria_id


categorias = sorted(
    mapa_categorias.keys()
)


# Atualizar lista visível na CONFIG

config.update(
    range_name="A4",
    values=[
        ["CATEGORIAS DISPONÍVEIS"]
    ]
)

if categorias:

    config.update(
        range_name=f"A5:A{4 + len(categorias)}",
        values=[
            [categoria]
            for categoria in categorias
        ]
    )


# =========================================================
# DESCOBRIR ID DA CATEGORIA
# =========================================================

categoria_id = None

if categoria_atual.upper() != "TODAS":

    categoria_id = mapa_categorias.get(
        categoria_atual
    )

    if categoria_id is None:

        raise Exception(
            f"Categoria '{categoria_atual}' "
            "não encontrada no feed."
        )


# =========================================================
# DETECTAR TROCA DE CATEGORIA
# =========================================================

if not categoria_anterior:

    dados_atuais = aba.get_all_values()

    if (
        len(dados_atuais) > 1
        and
        len(dados_atuais[1]) >= 9
    ):

        categoria_anterior = (
            dados_atuais[1][8]
            .strip()
        )


trocou_categoria = (
    categoria_anterior
    and
    categoria_atual.lower()
    != categoria_anterior.lower()
)


if trocou_categoria:

    if aba.row_count > 1:

        aba.batch_clear([
            f"A2:N{aba.row_count}"
        ])

    print(
        f"Categoria alterada: "
        f"{categoria_anterior} → "
        f"{categoria_atual}"
    )


config.update(
    range_name="B3",
    values=[
        [categoria_atual]
    ]
)


# =========================================================
# LER PRODUTOS EXISTENTES
# =========================================================

dados_planilha = (
    aba.get_all_values()
)

produtos_existentes = {}


for numero_linha, linha in enumerate(
    dados_planilha[1:],
    start=2
):

    if not linha:
        continue

    item_id = str(
        linha[0]
    ).strip()

    if item_id:

        produtos_existentes[
            item_id
        ] = {
            "linha": numero_linha,
            "dados": linha
        }


# =========================================================
# ATUALIZAR PRODUTOS EXISTENTES
# =========================================================

atualizacoes = []


for item_id, existente in produtos_existentes.items():

    produto = buscar_produto_por_id(
        item_id
    )

    if not produto:
        continue

    linha_antiga = (
        existente["dados"]
    )

    preco_planilha = (
        numero(linha_antiga[3])
        if len(linha_antiga) > 3
        else 0
    )

    preco_api = float(
        produto.get("priceMin")
        or 0
    )

    link_afiliado = (
        linha_antiga[10].strip()
        if len(linha_antiga) > 10
        else ""
    )


    # Se já tem link e preço não mudou,
    # não faz nada

    if (
        round(preco_planilha, 2)
        ==
        round(preco_api, 2)
        and
        link_afiliado
    ):

        continue


    # Se não tem link, gerar automaticamente

    if not link_afiliado:

        link_afiliado = (
            gerar_link_afiliado(
                produto["productLink"]
            )
        )


    desconto = float(
        produto.get(
            "priceDiscountRate"
        ) or 0
    )

    preco_anterior = (
        calcular_preco_original(
            preco_api,
            desconto
        )
    )


    legenda = gerar_legenda(
        produto,
        link_afiliado
    )


    nova_linha = [[

        str(produto["itemId"]),
        "Shopee",
        produto["productName"],
        preco_api,
        preco_anterior,
        desconto,
        float(
            produto.get(
                "ratingStar"
            ) or 0
        ),
        int(
            produto.get(
                "sales"
            ) or 0
        ),
        categoria_atual,
        produto["productLink"],
        link_afiliado,
        legenda,
        "PRONTO",
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        )

    ]]


    atualizacoes.append({

        "range":
            f"A{existente['linha']}:"
            f"N{existente['linha']}",

        "values":
            nova_linha

    })


    time.sleep(0.15)


if atualizacoes:

    aba.batch_update(
        atualizacoes,
        value_input_option="USER_ENTERED"
    )


# =========================================================
# BUSCAR 5 NOVAS MELHORES OFERTAS
# =========================================================

ids_existentes = set(
    produtos_existentes.keys()
)

novas_ofertas = []

pagina = 1

MAX_PAGINAS = 10


while (
    len(novas_ofertas) < 5
    and
    pagina <= MAX_PAGINAS
):

    resultado = buscar_produtos(
        categoria_id=categoria_id,
        pagina=pagina,
        limite=50
    )

    produtos = resultado.get(
        "nodes",
        []
    )


    if not produtos:
        break


    for produto in produtos:

        item_id = str(
            produto.get("itemId")
        )

        if item_id in ids_existentes:
            continue


        vendas = int(
            produto.get("sales")
            or 0
        )

        avaliacao = float(
            produto.get("ratingStar")
            or 0
        )

        desconto = float(
            produto.get(
                "priceDiscountRate"
            )
            or 0
        )

        comissao = float(
            produto.get(
                "commissionRate"
            )
            or 0
        )


        # ===== FILTROS =====

        if comissao < 0.05:
            continue

        if avaliacao < 4.8:
            continue

        if vendas <= 500:
            continue

        if desconto < 10:
            continue


        novas_ofertas.append(
            produto
        )

        ids_existentes.add(
            item_id
        )


        if len(novas_ofertas) >= 5:
            break


    page_info = resultado.get(
        "pageInfo",
        {}
    )

    if not page_info.get(
        "hasNextPage"
    ):
        break

    pagina += 1


# =========================================================
# ADICIONAR NOVAS OFERTAS
# =========================================================

linhas_novas = []


for produto in novas_ofertas:

    preco_atual = float(
        produto.get(
            "priceMin"
        ) or 0
    )

    desconto = float(
        produto.get(
            "priceDiscountRate"
        ) or 0
    )

    preco_anterior = (
        calcular_preco_original(
            preco_atual,
            desconto
        )
    )


    link_afiliado = (
        gerar_link_afiliado(
            produto["productLink"]
        )
    )


    legenda = gerar_legenda(
        produto,
        link_afiliado
    )


    linhas_novas.append([

        str(produto["itemId"]),
        "Shopee",
        produto["productName"],
        preco_atual,
        preco_anterior,
        desconto,
        float(
            produto.get(
                "ratingStar"
            ) or 0
        ),
        int(
            produto.get(
                "sales"
            ) or 0
        ),
        categoria_atual,
        produto["productLink"],
        link_afiliado,
        legenda,
        "PRONTO",
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        )

    ])


    time.sleep(0.15)


if linhas_novas:

    aba.append_rows(
        linhas_novas,
        value_input_option="USER_ENTERED"
    )


# =========================================================
# RESULTADO
# =========================================================

print(
    f"Categoria: {categoria_atual}"
)

print(
    f"Produtos atualizados: "
    f"{len(atualizacoes)}"
)

print(
    f"Novas ofertas: "
    f"{len(linhas_novas)}"
)

print(
    "Automação concluída."
)
