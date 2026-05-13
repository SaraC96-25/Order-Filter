import json
import re
from collections import defaultdict
from datetime import datetime, time
from io import BytesIO
from zoneinfo import ZoneInfo
from pathlib import Path
from tempfile import NamedTemporaryFile

import cairosvg
import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


API_VERSION = "2026-04"
BASE_DIR = Path(__file__).resolve().parent
LOGO_SVG_PATH = BASE_DIR / "assets" / "wowstampa_logo.svg"
BRAND_GREEN = "3AAA35"
BRAND_DARK = "1D1D1B"
LIGHT_GREEN = "EAF6E9"
LIGHT_GREY = "F3F5F7"


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
              }            }
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
            }          }
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
    # Versione senza read_products:
    # usa solo i campi del line item disponibili con read_orders.
    return normalize_text(item.get("title")) or normalize_text(item.get("name"))


def product_matches(product_title: str, wanted_products: list[str], mode: str) -> bool:
    if not wanted_products:
        return True

    product_cf = product_title.casefold()
    wanted_cf = [p.casefold() for p in wanted_products if p.strip()]

    if mode == "Esatto":
        return product_cf in wanted_cf

    return any(w in product_cf for w in wanted_cf)



def extract_units_per_item(product_title: str, line_item_name: str = "") -> tuple[int, str]:
    """
    Estrae il moltiplicatore dal nome prodotto/riga ordine.

    Esempi:
    - "10 T-shirt Economy" -> 10
    - "20 T-shirt Economy" -> 20
    - "100 Shopper Cotone" -> 100

    Se non trova un numero iniziale, ritorna 1.
    """
    candidates = [normalize_text(product_title), normalize_text(line_item_name)]

    for candidate in candidates:
        if not candidate:
            continue

        match = re.match(r"^\s*(\d+)\s*(?:x|×|-)?\s+", candidate, flags=re.IGNORECASE)
        if match:
            value = int(match.group(1))
            if value > 0:
                return value, "numero iniziale nel nome prodotto"

    return 1, "fallback 1"


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

    Versione senza read_products:
    1. cerca in customAttributes / line item properties, tipico per opzioni personalizzate VOPO;
    2. usa variantTitle come fallback, se contiene pattern tipo "Colore: Nero".
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

    # 2) Fallback su variantTitle: "Colore: Nero", "Taglia M / Nero", ecc.
    variant_title = normalize_text(item.get("variantTitle"))
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
    multiply_by_pack_size: bool,
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

            ordered_line_quantity = int(item.get(quantity_field) or 0)
            units_per_item, units_source = extract_units_per_item(
                product_title=product_title,
                line_item_name=item.get("name"),
            )

            counted_quantity = (
                ordered_line_quantity * units_per_item
                if multiply_by_pack_size
                else ordered_line_quantity
            )

            key = color
            summary[key] += counted_quantity

            detail_rows.append(
                {
                    "Ordine": order.get("name"),
                    "Data ordine": order.get("createdAt"),
                    "Prodotto": product_title,
                    "SKU": item.get("sku"),
                    "Nome riga ordine": item.get("name"),
                    "Quantità righe ordine": ordered_line_quantity,
                    "Pezzi per prodotto": units_per_item if multiply_by_pack_size else 1,
                    "Sorgente moltiplicatore": units_source if multiply_by_pack_size else "disattivato",
                    "Quantità totale calcolata": counted_quantity,
                    "Colore": color,
                    "Sorgente colore": color_source,
                    "Financial status": order.get("displayFinancialStatus"),
                    "Fulfillment status": order.get("displayFulfillmentStatus"),
                }
            )

    summary_rows = [
        {"Colore": color, "Quantità totale": qty}
        for color, qty in summary.items()
    ]

    summary_df = pd.DataFrame(summary_rows).sort_values(
        by=["Colore"], kind="stable"
    ) if summary_rows else pd.DataFrame(columns=["Colore", "Quantità totale"])

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
    "Filtra gli ordini Shopify per data e prodotto, calcola i pezzi reali e unifica il risultato finale per colore, senza distinguere per prodotto."
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
        "Quantità ordini da conteggiare",
        ["currentQuantity (al netto di rimborsi/rimozioni)", "quantity (quantità originale)"],
        index=0,
    )

    multiply_by_pack_size = st.checkbox(
        "Moltiplica per il numero iniziale nel nome prodotto",
        value=True,
        help="Esempio: se il prodotto è '10 T-shirt' e la quantità ordinata è 3, il totale diventa 10 × 3 = 30.",
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
            multiply_by_pack_size=multiply_by_pack_size,
            include_without_color=include_without_color,
        )

        st.subheader("Riepilogo unificato per colore")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.subheader("Dettaglio ordini")
        st.dataframe(detail_df, use_container_width=True, hide_index=True)

        st.info("I file di export includono un layout istituzionale Wowstampa con logo, periodo analizzato, parametri del report e riepilogo finale per colore.")

        branded_excel_bytes = create_branded_excel(
            summary_df=summary_df,
            detail_df=detail_df,
            start_date=start_date,
            end_date=end_date,
            selected_products=wanted_products,
            quantity_field=quantity_field,
            multiply_by_pack_size=multiply_by_pack_size,
            include_without_color=include_without_color,
        )
        branded_pdf_bytes = create_branded_pdf(
            summary_df=summary_df,
            start_date=start_date,
            end_date=end_date,
            selected_products=wanted_products,
            quantity_field=quantity_field,
            multiply_by_pack_size=multiply_by_pack_size,
            include_without_color=include_without_color,
        )

        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            st.download_button(
                "Scarica Excel istituzionale",
                data=branded_excel_bytes,
                file_name=f"wowstampa_report_colori_{start_date}_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_dl2:
            st.download_button(
                "Scarica PDF istituzionale",
                data=branded_pdf_bytes,
                file_name=f"wowstampa_report_colori_{start_date}_{end_date}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with col_dl3:
            csv_bytes = summary_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Scarica CSV riepilogo",
                data=csv_bytes,
                file_name=f"report_shopify_colori_{start_date}_{end_date}.csv",
                mime="text/csv",
                use_container_width=True,
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
                            "Selected options": "Non lette: manca scope read_products",
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
