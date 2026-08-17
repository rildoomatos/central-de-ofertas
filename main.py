import json
import os
import time
import hashlib
from urllib.parse import quote

import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# CONFIGURAÇÕES GERAIS
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
# OFERTAS - GARANTIR 21 COLUNAS
# =========================================================

if aba.col_count < 21:
    aba.resize(cols=21)


cabecalhos = [[
    "ID",
    "Marketplace",
    "Produto",
    "Preço Atual",
    "Preço Anterior",
    "Desconto",
    "Avaliação",
    "Vendas",
    "Categoria",
    "Link Original",
    "Link Afiliado",
    "Legenda",
    "Status",
    "Data",
    "Imagem",
    "Comissão %",
    "Comissão estimada",
    "Pontuação",
    "Enviar WhatsApp",
    "Texto Status",
    "Abrir Imagem"
]]

aba.update(
    range_name="A1:U1",
    values=cabecalhos
)


# =========================================================
# TRADUÇÃO DAS CATEGORIAS
# =========================================================

TRADUCOES_CATEGORIAS = {

    "women clothes": "Roupas Femininas",
    "women's clothes": "Roupas Femininas",
    "women apparel": "Roupas Femininas",
    "women's apparel": "Roupas Femininas",

    "men clothes": "Roupas Masculinas",
    "men's clothes": "Roupas Masculinas",
    "men apparel": "Roupas Masculinas",
    "men's apparel": "Roupas Masculinas",

    "women shoes": "Calçados Femininos",
    "women's shoes": "Calçados Femininos",

    "men shoes": "Calçados Masculinos",
    "men's shoes": "Calçados Masculinos",

    "women bags": "Bolsas Femininas",
    "women's bags": "Bolsas Femininas",

    "men bags": "Bolsas Masculinas",
    "men's bags": "Bolsas Masculinas",

    "fashion accessories": "Acessórios de Moda",

    "watches": "Relógios",

    "mobile & gadgets": "Celulares e Acessórios",
    "mobile and gadgets": "Celulares e Acessórios",
    "phones & accessories": "Celulares e Acessórios",

    "computers & accessories": "Computadores e Acessórios",
    "computers and accessories": "Computadores e Acessórios",

    "home appliances": "Eletrodomésticos",

    "home & living": "Casa e Decoração",
    "home and living": "Casa e Decoração",

    "home & garden": "Casa e Jardim",
    "home and garden": "Casa e Jardim",

    "kitchenware": "Cozinha",

    "beauty & personal care": "Beleza e Cuidados Pessoais",
    "beauty and personal care": "Beleza e Cuidados Pessoais",
    "beauty": "Beleza",

    "health": "Saúde",
    "health & personal care": "Saúde e Cuidados Pessoais",

    "mom & baby": "Mamãe e Bebê",
    "mom and baby": "Mamãe e Bebê",

    "babies & kids": "Bebês e Crianças",
    "babies and kids": "Bebês e Crianças",
    "baby & kids": "Bebês e Crianças",

    "toys": "Brinquedos",
    "toys & games": "Brinquedos e Jogos",
    "toys and games": "Brinquedos e Jogos",

    "sports & outdoors": "Esportes e Lazer",
    "sports and outdoors": "Esportes e Lazer",
    "sports": "Esportes",

    "hobbies & collections": "Hobbies e Colecionáveis",
    "hobbies and collections": "Hobbies e Colecionáveis",

    "gaming & consoles": "Games e Consoles",
    "gaming and consoles": "Games e Consoles",
    "games & consoles": "Games e Consoles",

    "cameras & drones": "Câmeras e Drones",
    "cameras and drones": "Câmeras e Drones",

    "audio": "Áudio",

    "food & beverages": "Alimentos e Bebidas",
    "food and beverages": "Alimentos e Bebidas",

    "groceries": "Mercado",

    "pets": "Produtos para Pets",
    "pet supplies": "Produtos para Pets",

    "automotive": "Automotivo",
    "automobiles": "Automotivo",

    "motorcycles": "Motos",
    "motorcycle": "Motos",

    "books & magazines": "Livros e Revistas",
    "books and magazines": "Livros e Revistas",
    "books": "Livros",

    "stationery": "Papelaria",

    "office & stationery": "Escritório e Papelaria",
    "office and stationery": "Escritório e Papelaria",

    "travel & luggage": "Viagem e Bagagem",
    "travel and luggage": "Viagem e Bagagem",

    "tickets & vouchers": "Ingressos e Vouchers",
    "tickets and vouchers": "Ingressos e Vouchers",

    "electronics": "Eletrônicos",
    "consumer electronics": "Eletrônicos",

    "tools & home improvement": "Ferramentas e Construção",
    "tools and home improvement": "Ferramentas e Construção"
}


def traduzir_categoria(nome):

    nome = str(nome).strip()

    return TRADUCOES_CATEGORIAS.get(
        nome.lower(),
        nome
    )


# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def numero(valor):

    try:

        texto = (
            str(valor)
            .replace("R$", "")
            .replace("%", "")
            .strip()
        )

        if "," in texto:

            texto = (
                texto
                .replace(".", "")
                .replace(",", ".")
            )

        return float(texto)

    except:

        return 0.0


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


def calcular_preco_original(
    preco_atual,
    desconto
):

    preco_atual = float(preco_atual)
    desconto = float(desconto)

    if desconto <= 0 or desconto >= 100:
        return preco_atual

    return round(
        preco_atual / (1 - desconto / 100),
        2
    )


def calcular_score(produto):

    vendas = int(
        produto.get("sales")
        or 0
    )

    avaliacao = float(
        produto.get("ratingStar")
        or 0
    )

    desconto = float(
        produto.get("priceDiscountRate")
        or 0
    )

    comissao = float(
        produto.get("commissionRate")
        or 0
    )

    comissao_percentual = (
        comissao * 100
    )

    score = (
        (comissao_percentual * 3)
        +
        (min(vendas, 10000) / 100)
        +
        desconto
        +
        (avaliacao * 5)
    )

    return round(score, 2)


def calcular_comissao(produto):

    try:

        comissao = float(
            produto.get("commission")
            or 0
        )

        if comissao > 0:
            return round(comissao, 2)

    except:
        pass

    preco = float(
        produto.get("priceMin")
        or 0
    )

    taxa = float(
        produto.get("commissionRate")
        or 0
    )

    return round(
        preco * taxa,
        2
    )


def formula_imagem(url):

    if not url:
        return ""

    url = str(url).replace('"', "")

    return f'=IMAGE("{url}")'


def formula_abrir_imagem(url):

    if not url:
        return ""

    url = str(url).replace('"', "")

    return (
        f'=HYPERLINK("{url}";"ABRIR IMAGEM")'
    )


def formula_whatsapp(legenda):

    if not legenda:
        return ""

    texto_codificado = quote(
        str(legenda),
        safe="",
        encoding="utf-8"
    )

    url = (
        "https://api.whatsapp.com/send/"
        f"?text={texto_codificado}"
    )

    return (
        f'=HYPERLINK("{url}";"ENVIAR")'
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
            f"Erro API Shopee: "
            f"{resultado['errors']}"
        )

    return resultado.get(
        "data",
        {}
    )


# =========================================================
# CAMPOS DA API
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


# =========================================================
# BUSCAR PRODUTOS
# =========================================================

def buscar_produtos(
    categoria_id=None,
    palavra_chave="",
    marca="",
    pagina=1,
    limite=50
):

    filtros = [
        "listType:0"
    ]

    if categoria_id:

        filtros.append(
            f"productCatId:{int(categoria_id)}"
        )


    termos = []

    if palavra_chave.strip():
        termos.append(
            palavra_chave.strip()
        )

    if marca.strip():
        termos.append(
            marca.strip()
        )


    termo_busca = " ".join(
        termos
    ).strip()


    if termo_busca:

        termo_graphql = json.dumps(
            termo_busca,
            ensure_ascii=False
        )

        filtros.append(
            f"keyword:{termo_graphql}"
        )

        # Relevância quando existe palavra-chave
        filtros.append(
            "sortType:1"
        )

    else:

        # Mais vendidos quando não existe palavra-chave
        filtros.append(
            "sortType:2"
        )


    filtros.append(
        f"page:{pagina}"
    )

    filtros.append(
        f"limit:{limite}"
    )


    argumentos = "\n".join(
        filtros
    )


    query = f"""
    {{
      productOfferV2(
        {argumentos}
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


# =========================================================
# BUSCAR PRODUTO PELO ID
# =========================================================

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
# GERAR LINK AFILIADO
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
# LEGENDA PARA GRUPO
# =========================================================

def gerar_legenda(
    produto,
    link_afiliado
):

    preco_atual = float(
        produto.get("priceMin")
        or 0
    )

    desconto = float(
        produto.get("priceDiscountRate")
        or 0
    )

    preco_anterior = (
        calcular_preco_original(
            preco_atual,
            desconto
        )
    )

    avaliacao = float(
        produto.get("ratingStar")
        or 0
    )

    vendas = int(
        produto.get("sales")
        or 0
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
        "⚠️ Preço e disponibilidade podem "
        "mudar a qualquer momento!"
    )


# =========================================================
# TEXTO PARA STATUS
# =========================================================

def gerar_texto_status(
    produto,
    link_afiliado
):

    preco_atual = float(
        produto.get("priceMin")
        or 0
    )

    desconto = float(
        produto.get("priceDiscountRate")
        or 0
    )

    preco_anterior = (
        calcular_preco_original(
            preco_atual,
            desconto
        )
    )

    nome_produto = str(
        produto.get(
            "productName",
            ""
        )
    ).strip()

    return (
        "🔥 *OFERTA SHOPEE*\n\n"
        f"🛍️ *{nome_produto}*\n\n"
        f"❌ De: {moeda(preco_anterior)}\n"
        f"✅ Por: *{moeda(preco_atual)}*\n"
        f"🔥 {int(desconto)}% OFF\n\n"
        "👇 *Confira aqui:*\n"
        f"{link_afiliado}"
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
        rows=300,
        cols=5
    )


# Garantir colunas A até E
if config.col_count < 5:
    config.resize(cols=5)


# =========================================================
# PRESERVAR BUSCA EXISTENTE
# =========================================================

categoria_existente = (
    config.acell("B2").value
    or "TODAS"
).strip()


palavra_existente = (
    config.acell("B3").value
    or ""
).strip()


marca_existente = (
    config.acell("B4").value
    or ""
).strip()


busca_anterior_existente = (
    config.acell("B5").value
    or ""
).strip()


# =========================================================
# PRESERVAR FILTROS OU CRIAR PADRÕES
# =========================================================

comissao_existente = (
    config.acell("E2").value
    or "5"
)

avaliacao_existente = (
    config.acell("E3").value
    or "4,8"
)

vendas_existente = (
    config.acell("E4").value
    or "500"
)

desconto_existente = (
    config.acell("E5").value
    or "10"
)

preco_min_existente = (
    config.acell("E6").value
    or "0"
)

preco_max_existente = (
    config.acell("E7").value
    or "0"
)

quantidade_existente = (
    config.acell("E8").value
    or "5"
)


# =========================================================
# MONTAR PAINEL CONFIG
# =========================================================

config.update(
    range_name="A1:B5",
    values=[
        [
            "CONFIGURAÇÃO",
            "VALOR"
        ],
        [
            "Categoria",
            categoria_existente
        ],
        [
            "Palavra-chave",
            palavra_existente
        ],
        [
            "Marca",
            marca_existente
        ],
        [
            "Busca anterior",
            busca_anterior_existente
        ]
    ]
)


config.update(
    range_name="D1:E8",
    values=[
        [
            "FILTROS",
            "VALOR"
        ],
        [
            "Comissão mínima (%)",
            comissao_existente
        ],
        [
            "Avaliação mínima",
            avaliacao_existente
        ],
        [
            "Vendas mínimas",
            vendas_existente
        ],
        [
            "Desconto mínimo (%)",
            desconto_existente
        ],
        [
            "Preço mínimo (R$)",
            preco_min_existente
        ],
        [
            "Preço máximo (R$)",
            preco_max_existente
        ],
        [
            "Quantidade de ofertas",
            quantidade_existente
        ]
    ]
)


# =========================================================
# LER PAINEL
# =========================================================

categoria_atual = (
    config.acell("B2").value
    or "TODAS"
).strip()

palavra_chave = (
    config.acell("B3").value
    or ""
).strip()

marca = (
    config.acell("B4").value
    or ""
).strip()

busca_anterior = (
    config.acell("B5").value
    or ""
).strip()


COMISSAO_MINIMA = numero(
    config.acell("E2").value
)

AVALIACAO_MINIMA = numero(
    config.acell("E3").value
)

VENDAS_MINIMAS = int(
    numero(
        config.acell("E4").value
    )
)

DESCONTO_MINIMO = numero(
    config.acell("E5").value
)

PRECO_MINIMO = numero(
    config.acell("E6").value
)

PRECO_MAXIMO = numero(
    config.acell("E7").value
)

QUANTIDADE_OFERTAS = int(
    numero(
        config.acell("E8").value
    )
)


# Segurança
if QUANTIDADE_OFERTAS < 1:
    QUANTIDADE_OFERTAS = 5

if QUANTIDADE_OFERTAS > 50:
    QUANTIDADE_OFERTAS = 50


if not categoria_atual:
    categoria_atual = "TODAS"


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

mapa_original_para_portugues = {}


for _, linha in (
    df_categorias.iterrows()
):

    nome_original = str(
        linha["global_category1"]
    ).strip()

    categoria_id_feed = int(
        linha["global_catid1"]
    )

    nome_portugues = traduzir_categoria(
        nome_original
    )


    mapa_original_para_portugues[
        nome_original.lower()
    ] = nome_portugues


    if nome_portugues not in mapa_categorias:

        mapa_categorias[
            nome_portugues
        ] = categoria_id_feed


# =========================================================
# CONVERTER CATEGORIA ANTIGA EM INGLÊS
# =========================================================

if categoria_atual.upper() != "TODAS":

    traducao_existente = (
        mapa_original_para_portugues.get(
            categoria_atual.lower()
        )
    )

    if traducao_existente:

        categoria_atual = (
            traducao_existente
        )

        config.update(
            range_name="B2",
            values=[
                [categoria_atual]
            ]
        )


# =========================================================
# LISTA DE CATEGORIAS
# =========================================================

categorias = sorted(
    mapa_categorias.keys()
)


# A lista continua em A8 para manter
# o menu suspenso que você já criou.
if config.row_count >= 7:

    config.batch_clear([
        f"A7:B{config.row_count}"
    ])


config.update(
    range_name="A7",
    values=[
        ["CATEGORIAS DISPONÍVEIS"]
    ]
)


lista_config = [
    ["TODAS"]
]

lista_config += [
    [categoria]
    for categoria
    in categorias
]


config.update(
    range_name=(
        f"A8:"
        f"A{7 + len(lista_config)}"
    ),
    values=lista_config
)


# =========================================================
# ID DA CATEGORIA
# =========================================================

categoria_id = None


if categoria_atual.upper() != "TODAS":

    categoria_id = (
        mapa_categorias.get(
            categoria_atual
        )
    )

    if categoria_id is None:

        raise Exception(
            f"Categoria "
            f"'{categoria_atual}' "
            f"não encontrada."
        )


# =========================================================
# IDENTIFICAR A BUSCA
# =========================================================

assinatura_busca = (
    f"{categoria_atual.lower()}|"
    f"{palavra_chave.lower()}|"
    f"{marca.lower()}|"
    f"{COMISSAO_MINIMA}|"
    f"{AVALIACAO_MINIMA}|"
    f"{VENDAS_MINIMAS}|"
    f"{DESCONTO_MINIMO}|"
    f"{PRECO_MINIMO}|"
    f"{PRECO_MAXIMO}|"
    f"{QUANTIDADE_OFERTAS}"
)


# =========================================================
# SE ALTERAR BUSCA OU FILTROS, LIMPAR OFERTAS
# =========================================================

trocou_busca = (
    busca_anterior
    and
    assinatura_busca
    != busca_anterior
)


if trocou_busca:

    if aba.row_count > 1:

        aba.batch_clear([
            f"A2:U{aba.row_count}"
        ])

    print(
        "Busca ou filtros alterados. "
        "Ofertas anteriores removidas."
    )


config.update(
    range_name="B5",
    values=[
        [assinatura_busca]
    ]
)


# =========================================================
# PRODUTOS EXISTENTES
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


for item_id, existente in (
    produtos_existentes.items()
):

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


    status_atual = (
        linha_antiga[12].strip()
        if len(linha_antiga) > 12
        else ""
    )


    if status_atual.upper() == "ENVIADO":
        status_final = "ENVIADO"
    else:
        status_final = "PRONTO"


    link_estava_ausente = (
        not link_afiliado
    )


    if link_estava_ausente:

        link_afiliado = gerar_link_afiliado(
            produto["productLink"]
        )


    desconto = float(
        produto.get("priceDiscountRate")
        or 0
    )


    preco_anterior = (
        calcular_preco_original(
            preco_api,
            desconto
        )
    )


    comissao_percentual = round(
        float(
            produto.get("commissionRate")
            or 0
        ) * 100,
        2
    )


    comissao_estimada = (
        calcular_comissao(
            produto
        )
    )


    score = calcular_score(
        produto
    )


    image_url = produto.get(
        "imageUrl",
        ""
    )


    imagem = formula_imagem(
        image_url
    )


    abrir_imagem = (
        formula_abrir_imagem(
            image_url
        )
    )


    texto_status = (
        gerar_texto_status(
            produto,
            link_afiliado
        )
    )


    preco_mudou = (
        round(preco_planilha, 2)
        !=
        round(preco_api, 2)
    )


    if (
        preco_mudou
        or
        link_estava_ausente
    ):

        legenda = gerar_legenda(
            produto,
            link_afiliado
        )

        enviar_whatsapp = (
            formula_whatsapp(
                legenda
            )
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
                )
                or 0
            ),
            int(
                produto.get(
                    "sales"
                )
                or 0
            ),
            categoria_atual,
            produto["productLink"],
            link_afiliado,
            legenda,
            status_final,
            pd.Timestamp.now().strftime(
                "%d/%m/%Y %H:%M"
            ),
            imagem,
            comissao_percentual,
            comissao_estimada,
            score,
            enviar_whatsapp,
            texto_status,
            abrir_imagem
        ]]


        atualizacoes.append({
            "range":
                f"A{existente['linha']}:"
                f"U{existente['linha']}",
            "values":
                nova_linha
        })


    else:

        legenda_existente = (
            linha_antiga[11]
            if len(linha_antiga) > 11
            else ""
        )


        if not legenda_existente:

            legenda_existente = (
                gerar_legenda(
                    produto,
                    link_afiliado
                )
            )


        enviar_whatsapp = (
            formula_whatsapp(
                legenda_existente
            )
        )


        atualizacoes.append({
            "range":
                f"O{existente['linha']}:"
                f"U{existente['linha']}",
            "values": [[
                imagem,
                comissao_percentual,
                comissao_estimada,
                score,
                enviar_whatsapp,
                texto_status,
                abrir_imagem
            ]]
        })


    time.sleep(0.15)


if atualizacoes:

    aba.batch_update(
        atualizacoes,
        value_input_option=
            "USER_ENTERED"
    )


# =========================================================
# BUSCAR NOVAS OFERTAS
# =========================================================

ids_existentes = set(
    produtos_existentes.keys()
)

candidatos = []

pagina = 1

MAX_PAGINAS = 10


while pagina <= MAX_PAGINAS:

    resultado = buscar_produtos(
        categoria_id=categoria_id,
        palavra_chave=palavra_chave,
        marca=marca,
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


        comissao_percentual = (
            comissao * 100
        )


        preco_atual = float(
            produto.get("priceMin")
            or 0
        )


        # =================================================
        # FILTROS CONTROLADOS PELA PLANILHA
        # =================================================

        if (
            comissao_percentual
            < COMISSAO_MINIMA
        ):
            continue


        if (
            avaliacao
            < AVALIACAO_MINIMA
        ):
            continue


        if (
            vendas
            < VENDAS_MINIMAS
        ):
            continue


        if (
            desconto
            < DESCONTO_MINIMO
        ):
            continue


        if (
            PRECO_MINIMO > 0
            and
            preco_atual < PRECO_MINIMO
        ):
            continue


        if (
            PRECO_MAXIMO > 0
            and
            preco_atual > PRECO_MAXIMO
        ):
            continue


        produto["_score"] = (
            calcular_score(
                produto
            )
        )


        candidatos.append(
            produto
        )


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
# RANKING
# =========================================================

candidatos = sorted(
    candidatos,
    key=lambda produto:
        produto["_score"],
    reverse=True
)


novas_ofertas = (
    candidatos[
        :QUANTIDADE_OFERTAS
    ]
)


# =========================================================
# ADICIONAR NOVAS OFERTAS
# =========================================================

linhas_novas = []


for produto in novas_ofertas:

    preco_atual = float(
        produto.get("priceMin")
        or 0
    )


    desconto = float(
        produto.get(
            "priceDiscountRate"
        )
        or 0
    )


    preco_anterior = (
        calcular_preco_original(
            preco_atual,
            desconto
        )
    )


    link_afiliado = gerar_link_afiliado(
        produto["productLink"]
    )


    legenda = gerar_legenda(
        produto,
        link_afiliado
    )


    texto_status = gerar_texto_status(
        produto,
        link_afiliado
    )


    comissao_percentual = round(
        float(
            produto.get(
                "commissionRate"
            )
            or 0
        ) * 100,
        2
    )


    comissao_estimada = (
        calcular_comissao(
            produto
        )
    )


    score = calcular_score(
        produto
    )


    image_url = produto.get(
        "imageUrl",
        ""
    )


    imagem = formula_imagem(
        image_url
    )


    abrir_imagem = (
        formula_abrir_imagem(
            image_url
        )
    )


    enviar_whatsapp = (
        formula_whatsapp(
            legenda
        )
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
            )
            or 0
        ),
        int(
            produto.get(
                "sales"
            )
            or 0
        ),
        categoria_atual,
        produto["productLink"],
        link_afiliado,
        legenda,
        "PRONTO",
        pd.Timestamp.now().strftime(
            "%d/%m/%Y %H:%M"
        ),
        imagem,
        comissao_percentual,
        comissao_estimada,
        score,
        enviar_whatsapp,
        texto_status,
        abrir_imagem
    ])


    time.sleep(0.15)


if linhas_novas:

    aba.append_rows(
        linhas_novas,
        value_input_option=
            "USER_ENTERED"
    )


# =========================================================
# RESULTADO
# =========================================================

print(
    f"Categoria: "
    f"{categoria_atual}"
)

print(
    f"Palavra-chave: "
    f"{palavra_chave or 'Nenhuma'}"
)

print(
    f"Marca: "
    f"{marca or 'Nenhuma'}"
)

print(
    f"Comissão mínima: "
    f"{COMISSAO_MINIMA}%"
)

print(
    f"Avaliação mínima: "
    f"{AVALIACAO_MINIMA}"
)

print(
    f"Vendas mínimas: "
    f"{VENDAS_MINIMAS}"
)

print(
    f"Desconto mínimo: "
    f"{DESCONTO_MINIMO}%"
)

print(
    f"Preço mínimo: "
    f"{PRECO_MINIMO}"
)

print(
    f"Preço máximo: "
    f"{PRECO_MAXIMO or 'Sem limite'}"
)

print(
    f"Quantidade solicitada: "
    f"{QUANTIDADE_OFERTAS}"
)

print(
    f"Produtos aprovados nos filtros: "
    f"{len(candidatos)}"
)

print(
    f"Novas ofertas adicionadas: "
    f"{len(linhas_novas)}"
)

print(
    "Automação concluída com sucesso."
)
