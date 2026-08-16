import os
import time
import json
import hashlib
import requests

APP_ID = os.environ["SHOPEE_APP_ID"]
SECRET = os.environ["SHOPEE_API_PASSWORD"]

API_URL = "https://open-api.affiliate.shopee.com.br/graphql"


def chamar_api(query):
    payload = json.dumps(
        {"query": query},
        separators=(",", ":")
    )

    timestamp = str(int(time.time()))

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

    print("HTTP:", resposta.status_code)
    print(resposta.text)


# ===== TESTE =====

query = """
{
  productOfferV2(
    listType: 0
    sortType: 2
    page: 1
    limit: 5
  ) {
    nodes {
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
    }
    pageInfo {
      page
      limit
      hasNextPage
    }
  }
}
"""

chamar_api(query)
