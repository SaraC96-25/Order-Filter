import json
import re
from collections import defaultdict
from datetime import datetime, time
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st


API_VERSION = "2026-04"


def get_secret(name: str, default: str = "") -> str:
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


SHOPIFY_SHOP_DOMAIN = get_secret("SHOPIFY_SHOP_DOMAIN")
SHOPIFY_ACCESS_TOKEN = get_secret("SHOPIFY_ACCESS_TOKEN")


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_key(value):
    return normalize_text(value).casefold().replace(":", "").replace("_", " ").strip()


def gql_request(query: str, variables: dict | None = None) -> dict:
    if not SHOPIFY_SHOP_DOMAIN or not SHOPIFY_ACCESS_TOKEN:
        raise RuntimeError(
            "Configura SHOPIFY_SHOP_DOMAIN e SHOPIFY_ACCESS_TOKEN nei secrets di Streamlit."
        )

    domain = SHOPIFY_SHOP_DOMAIN.replace("https://", "").replace("http://", "").strip("/")
    url = f"https://{domain}/admin/api/{API_VERSION}/graphql.json"

    response = requests.post(
        url,
        headers={
            "Content-Type": "application/json",
            "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        },
        json={"query": query, "variables": variables or {}},
        timeout=60,
    )

    if response.status_code != 200:
        raise RuntimeError(f"Errore HTTP Shopify {response.status_code}: {response.text}")

    payload = response.json()
    if payload.get("errors"):
        raise RuntimeError(json.dumps(payload["errors"], indent=2, ensure_ascii=False))

    return payload["data"]


ORDERS_QUERY = """
query OrdersForReport($first: Int!, $after: String, $query: String!) {
  orders(first: $first, after: $after, query: $query, sortKey: CREATED_AT) {
    pageInfo {
      hasNextPage
      endCursor
    }
    edges {
      node {
        id
        name
        createdAt
        cancelledAt
        displayFinancialStatus
        displayFulfillmentStatus
        lineItems(first: 250) {
          pageInfo {
            hasNextPage
            endCursor
          }
          edges {
            node {
              id
              name
              title
              quantity
              currentQuantity
              sku
              variantTitle
              customAttributes {
                key
                value
              }
              product {
                id
                title
                handle
              }
              variant {
                id
                title
                sku
                selectedOptions {
                  name
                  value
                }
                product {
                  id
                  title
                  handle
                }
              }
            }
          }
        }
      }
    }
  }
}
"""


ORDER_LINE_ITEMS_QUERY = """
query OrderLineItems($orderId: ID!, $first: Int!, $after: String) {
  node(id: $orderId) {
    ... on Order {
      lineItems(first: $first, after: $after) {
        pageInfo {
          hasNextPage
          endCursor
        }
        edges {
          node {
            id
            name
            title
            quantity
            currentQuantity
            sku
            variantTitle
            customAttributes {
              key
              value
            }
            product {
              id
              title
              handle
            }
            variant {
              id
              title
              sku
              selectedOptions {
                name
                value
              }
              product {
                id
                title
                handle
              }
            }
          }
        }
      }
    }
  }
}
"""


def fetch_more_line_items(order_id: str, after: str) -> list[dict]:
    all_items = []
    cursor = after

    while cursor:
        data = gql_request(
            ORDER_LINE_ITEMS_QUERY,
            {"orderId": order_id, "first": 250, "after": cursor},
        )
        line_items = data["node"]["lineItems"]
        all_items.extend([edge["node"] for edge in line_items["edges"]])
        if not line_items["pageInfo"]["hasNextPage"]:
            break
        cursor = line_items["pageInfo"]["endCursor"]

    return all_items


def fetch_orders(shopify_query: str) -> list[dict]:
    orders = []
    cursor = None

    progress = st.progress(0, text="Lettura ordini da Shopify...")
    pages = 0

    while True:
        data = gql_request(
            ORDERS_QUERY,
            {"first": 100, "after": cursor, "query": shopify_query},
        )

        conn = data["orders"]
        for edge in conn["edges"]:
            order = edge["node"]
            line_items_conn = order["lineItems"]
            line_items = [li_edge["node"] for li_edge in line_items_conn["edges"]]

            if line_items_conn["pageInfo"]["hasNextPage"]:
                extra_items = fetch_more_line_items(
                    order["id"], line_items_conn["pageInfo"]["endCursor"]
                )
                line_items.extend(extra_items)

            order["lineItems_flat"] = line_items
            orders.append(order)

        pages += 1
        progress.progress(
            min(95, pages * 10),
            text=f"Lettura ordini da Shopify... pagine lette: {pages}, ordini: {len(orders)}",
        )

        if not conn["pageInfo"]["hasNextPage"]:
            break
        cursor = conn["pageInfo"]["endCursor"]

    progress.progress(100, text=f"Completato: {len(orders)} ordini letti.")
    return orders


def build_shopify_date_query(start_date, end_date, timezone_name: str, exclude_cancelled: bool) -> str:
    tz = ZoneInfo(timezone_name)
    start_local = datetime.combine(start_date, time.min).replace(tzinfo=tz)
    end_local = datetime.combine(end_date, time.max).replace(tzinfo=tz)

    start_utc = start_local.astimezone(ZoneInfo("UTC")).isoformat().replace("+00:00", "Z")
    end_utc = end_local.astimezone(ZoneInfo("UTC")).isoformat().replace("+00:00", "Z")

    parts = [f"created_at:>={start_utc}", f"created_at:<={end_utc}"]
    if exclude_cancelled:
        parts.append("-status:cancelled")
    return " ".join(parts)


def product_title_from_line_item(item: dict) -> str:
    product = item.get("product") or {}
    variant = item.get("variant") or {}
    variant_product = variant.get("product") or {}

    return (
        normalize_text(product.get("title"))
        or normalize_text(variant_product.get("title"))
        or normalize_text(item.get("title"))
        or normalize_text(item.get("name"))
    )


def product_matches(product_title: str, wanted_products: list[str], mode: str) -> bool:
    if not wanted_products:
        return True

    product_cf = product_title.casefold()
    wanted_cf = [p.casefold() for p in wanted_products if p.strip()]

    if mode == "Esatto":
        return product_cf in wanted_cf

    return any(w in product_cf for w in wanted_cf)


def parse_attribute_json(value: str) -> dict:
    value = normalize_text(value)
    if not value:
        return {}
    try:
        parsed = json.loads(value)
    except Exception:
        return {}

    if isinstance(parsed, dict):
        return parsed

    return {}


def extract_color(item: dict, color_keys: list[str]) -> tuple[str, str]:
    """
    Restituisce (colore, sorgente).
    Cerca in:
    1. customAttributes / line item properties, tipico per opzioni personalizzate VOPO
    2. selectedOptions della variante Shopify
    3. variantTitle come fallback, se contiene pattern tipo Colore: Nero
    """

    normalized_color_keys = {normalize_key(k) for k in color_keys if normalize_key(k)}

    # 1) Custom attributes / line item properties
    for attr in item.get("customAttributes") or []:
        key = normalize_text(attr.get("key"))
        value = normalize_text(attr.get("value"))

        if normalize_key(key) in normalized_color_keys and value:
            return value, f"customAttributes.{key}"

        # Alcune app salvano un JSON in una property.
        nested = parse_attribute_json(value)
        for nested_key, nested_value in nested.items():
            if normalize_key(nested_key) in normalized_color_keys and normalize_text(nested_value):
                return normalize_text(nested_value), f"customAttributes.{key}.{nested_key}"

    # 2) Selected options della variante Shopify
    variant = item.get("variant") or {}
    for opt in variant.get("selectedOptions") or []:
        name = normalize_text(opt.get("name"))
        value = normalize_text(opt.get("value"))
        if normalize_key(name) in normalized_color_keys and value:
            return value, f"variant.selectedOptions.{name}"

    # 3) Fallback su variantTitle: "Colore: Nero", "Taglia M / Nero", ecc.
    variant_title = normalize_text(item.get("variantTitle") or variant.get("title"))
    if variant_title and variant_title.casefold() != "default title":
        for color_key in color_keys:
            pattern = rf"{re.escape(color_key)}\s*[:=-]\s*([^,/|]+)"
            match = re.search(pattern, variant_title, flags=re.IGNORECASE)
            if match:
                return match.group(1).strip(), "variantTitle pattern"

    return "", ""


def build_report(
    orders: list[dict],
    wanted_products: list[str],
    product_match_mode: str,
    color_keys: list[str],
    quantity_field: str,
    include_without_color: bool,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    summary = defaultdict(int)
    detail_rows = []

    for order in orders:
        for item in order.get("lineItems_flat", []):
            product_title = product_title_from_line_item(item)
            if not product_matches(product_title, wanted_products, product_match_mode):
                continue

            color, color_source = extract_color(item, color_keys)
            if not color and not include_without_color:
                continue

            if not color:
                color = "Senza colore"

            quantity = int(item.get(quantity_field) or 0)

            key = (product_title, color)
            summary[key] += quantity

            detail_rows.append(
                {
                    "Ordine": order.get("name"),
                    "Data ordine": order.get("createdAt"),
                    "Prodotto": product_title,
                    "SKU": item.get("sku") or ((item.get("variant") or {}).get("sku")),
                    "Nome riga ordine": item.get("name"),
                    "Quantità conteggiata": quantity,
                    "Colore": color,
                    "Sorgente colore": color_source,
                    "Financial status": order.get("displayFinancialStatus"),
                    "Fulfillment status": order.get("displayFulfillmentStatus"),
                }
            )

    summary_rows = [
        {"Prodotto": product, "Colore": color, "Quantità totale": qty}
        for (product, color), qty in summary.items()
    ]

    summary_df = pd.DataFrame(summary_rows).sort_values(
        by=["Prodotto", "Colore"], kind="stable"
    ) if summary_rows else pd.DataFrame(columns=["Prodotto", "Colore", "Quantità totale"])

    detail_df = pd.DataFrame(detail_rows)
    return summary_df, detail_df


def dataframe_to_excel(summary_df: pd.DataFrame, detail_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Riepilogo")
        detail_df.to_excel(writer, index=False, sheet_name="Dettaglio ordini")

        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    return output.getvalue()


st.set_page_config(page_title="Report Shopify colori", page_icon="🎨", layout="wide")

st.title("🎨 Report Shopify: quantità ordinate per colore")
st.caption(
    "Filtra gli ordini Shopify per data e prodotto, poi raggruppa le quantità per variante/opzione colore."
)

with st.sidebar:
    st.header("Impostazioni")

    timezone_name = st.text_input("Timezone negozio", value="Europe/Rome")

    st.subheader("Date ordine")
    start_date = st.date_input("Da")
    end_date = st.date_input("A")

    exclude_cancelled = st.checkbox("Escludi ordini cancellati", value=True)

    st.subheader("Prodotti")
    product_text = st.text_area(
        "Titoli prodotto da includere, uno per riga",
        placeholder="20 T-shirt Economy\nFelpa Premium",
    )
    product_match_mode = st.radio(
        "Match prodotto",
        ["Contiene", "Esatto"],
        index=0,
        horizontal=True,
    )

    st.subheader("Campo colore")
    color_keys_text = st.text_input(
        "Nomi possibili del campo colore",
        value="Colore, Color, colour, colore",
        help="Se VOPO salva il campo con un nome diverso, aggiungilo qui.",
    )

    quantity_label = st.radio(
        "Quantità da conteggiare",
        ["currentQuantity (al netto di rimborsi/rimozioni)", "quantity (quantità originale)"],
        index=0,
    )

    include_without_color = st.checkbox("Mostra anche righe senza colore", value=False)

    run = st.button("Genera report", type="primary")


wanted_products = [line.strip() for line in product_text.splitlines() if line.strip()]
color_keys = [x.strip() for x in color_keys_text.split(",") if x.strip()]
quantity_field = "currentQuantity" if quantity_label.startswith("currentQuantity") else "quantity"


with st.expander("Come impostare i secrets su Streamlit Cloud"):
    st.code(
        """
SHOPIFY_SHOP_DOMAIN = "tuo-negozio.myshopify.com"
SHOPIFY_ACCESS_TOKEN = "shpat_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
        """.strip(),
        language="toml",
    )
    st.write("Non mettere mai il token dentro GitHub. Usa i secrets di Streamlit Cloud.")


if run:
    if start_date > end_date:
        st.error("La data iniziale non può essere successiva alla data finale.")
        st.stop()

    if not wanted_products:
        st.warning("Non hai indicato prodotti: il report includerà tutti i prodotti trovati negli ordini.")

    try:
        shopify_query = build_shopify_date_query(
            start_date=start_date,
            end_date=end_date,
            timezone_name=timezone_name,
            exclude_cancelled=exclude_cancelled,
        )

        st.info(f"Query Shopify: `{shopify_query}`")

        orders = fetch_orders(shopify_query)

        summary_df, detail_df = build_report(
            orders=orders,
            wanted_products=wanted_products,
            product_match_mode=product_match_mode,
            color_keys=color_keys,
            quantity_field=quantity_field,
            include_without_color=include_without_color,
        )

        st.subheader("Riepilogo")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.subheader("Dettaglio ordini")
        st.dataframe(detail_df, use_container_width=True, hide_index=True)

        excel_bytes = dataframe_to_excel(summary_df, detail_df)
        st.download_button(
            "Scarica Excel",
            data=excel_bytes,
            file_name=f"report_shopify_colori_{start_date}_{end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        csv_bytes = summary_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Scarica CSV riepilogo",
            data=csv_bytes,
            file_name=f"report_shopify_colori_{start_date}_{end_date}.csv",
            mime="text/csv",
        )

        with st.expander("Debug: campi colore trovati nei line item"):
            debug_rows = []
            for order in orders[:25]:
                for item in order.get("lineItems_flat", []):
                    product_title = product_title_from_line_item(item)
                    if not product_matches(product_title, wanted_products, product_match_mode):
                        continue
                    debug_rows.append(
                        {
                            "Ordine": order.get("name"),
                            "Prodotto": product_title,
                            "Nome riga ordine": item.get("name"),
                            "Variant title": item.get("variantTitle"),
                            "Selected options": json.dumps(
                                ((item.get("variant") or {}).get("selectedOptions") or []),
                                ensure_ascii=False,
                            ),
                            "Custom attributes": json.dumps(
                                item.get("customAttributes") or [],
                                ensure_ascii=False,
                            ),
                        }
                    )
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)

    except Exception as exc:
        st.error(str(exc))
        st.stop()
