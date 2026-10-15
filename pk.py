from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import requests
import streamlit as st


# ============================================================
# INSTELLINGEN
# ============================================================

EXCEL_FILE = Path(__file__).parent / "Map1.xlsx"

BASE_URL = "https://api.poketrace.com/v1"
TIMEOUT = 30


# API-key komt uit Streamlit Secrets
API_KEY = st.secrets["POKETRACE_API_KEY"]

HEADERS = {
    "X-API-Key": API_KEY,
    "Accept": "application/json",
    "User-Agent": "Pokemon-Boosterbox-Excel-Tracker/1.0",
}


# ============================================================
# BESTAANDE FUNCTIES
# ============================================================

def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    return " ".join(
        str(value).lower().replace("-", " ").split()
    )


def api_get(endpoint: str, params: dict | None = None) -> dict:
    url = f"{BASE_URL}{endpoint}"

    response = requests.get(
        url,
        headers=HEADERS,
        params=params,
        timeout=TIMEOUT
    )

    if response.status_code == 401:
        raise RuntimeError(
            "API-key geweigerd. Controleer je PokeTrace API-key."
        )

    if response.status_code == 403:
        raise RuntimeError(
            "Geen toegang tot deze data."
        )

    if response.status_code == 429:
        raise RuntimeError(
            "De API-limiet werd bereikt."
        )

    if not response.ok:
        raise RuntimeError(
            f"PokeTrace fout {response.status_code}: "
            f"{response.text[:500]}"
        )

    return response.json()


def find_booster_box(product_name: str) -> dict:
    params = {
        "search": product_name,
        "market": "EU",
        "game": "pokemon",
        "product_type": "sealed",
        "product_family": "booster_box",
        "limit": 20,
    }

    payload = api_get("/cards", params=params)
    products = payload.get("data", [])

    if not products:
        raise LookupError(
            f"Geen boosterbox gevonden voor '{product_name}'."
        )

    wanted = normalize_text(product_name)

    for product in products:
        candidate = normalize_text(product.get("name"))

        if candidate == wanted:
            return product

    wanted_words = set(wanted.split())

    for product in products:
        candidate = normalize_text(product.get("name"))
        candidate_words = set(candidate.split())

        if wanted_words.issubset(candidate_words):
            return product

    available_names = [
        product.get("name", "<naam ontbreekt>")
        for product in products[:10]
    ]

    raise LookupError(
        f"Geen eenduidige match voor '{product_name}'. "
        f"Gevonden kandidaten: {available_names}"
    )


def get_product_details(product_id: str) -> dict:
    payload = api_get(
        f"/cards/{product_id}",
        params={"market": "EU"},
    )

    return payload.get("data", payload)


def collect_english_price_groups(
    value: Any,
    path: tuple[str, ...] = (),
) -> list:

    results = []

    if isinstance(value, dict):
        for key, child in value.items():
            new_path = path + (str(key),)

            if str(key).upper() == "EN" and isinstance(child, dict):
                avg = child.get("avg")
                count = child.get("saleCount")

                if avg is not None:
                    results.append({
                        "path": " > ".join(new_path),
                        "avg": float(avg),
                        "count": (
                            int(count)
                            if count is not None
                            else None
                        ),
                    })

            results.extend(
                collect_english_price_groups(child, new_path)
            )

    elif isinstance(value, list):
        for position, child in enumerate(value):
            results.extend(
                collect_english_price_groups(
                    child,
                    path + (str(position),),
                )
            )

    return results


def calculate_english_average(details: dict):
    prices = details.get("prices", {})
    unsold = prices.get("cardmarket_unsold", {})

    groups = collect_english_price_groups(unsold)

    if not groups:
        return None, None, 0

    groups_with_count = [
        group for group in groups
        if group["count"] is not None and group["count"] > 0
    ]

    if groups_with_count:
        total_count = sum(
            group["count"] for group in groups_with_count
        )

        weighted_sum = sum(
            group["avg"] * group["count"]
            for group in groups_with_count
        )

        return (
            weighted_sum / total_count,
            total_count,
            len(groups)
        )

    simple_average = (
        sum(group["avg"] for group in groups) / len(groups)
    )

    return simple_average, None, len(groups)


# ============================================================
# PRIJZEN UPDATEN
# ============================================================

def update_prices():

    df = pd.read_excel(
        EXCEL_FILE,
        engine="openpyxl"
    )

    if "Product_Name" not in df.columns:
        raise ValueError(
            "De kolom 'Product_Name' ontbreekt in Map1.xlsx."
        )

    today = datetime.now().strftime("%Y-%m-%d")

    price_column = f"CM_EN_Avg_{today}"
    status_column = f"Status_{today}"

    df[price_column] = pd.NA
    df[status_column] = pd.NA

    # Permanente metadata
    if "PokeTrace_ID" not in df.columns:
        df["PokeTrace_ID"] = pd.NA

    if "Cardmarket_ID" not in df.columns:
        df["Cardmarket_ID"] = pd.NA

    if "Matched_Product" not in df.columns:
        df["Matched_Product"] = pd.NA

    progress = st.progress(0)

    total = len(df)

    for position, (index, row) in enumerate(df.iterrows()):

        product_name = row["Product_Name"]

        try:

            existing_id = row.get("PokeTrace_ID")

            if pd.notna(existing_id):
                product_id = str(existing_id)

            else:
                result = find_booster_box(product_name)

                product_id = str(result["id"])

                df.loc[index, "PokeTrace_ID"] = product_id

            details = get_product_details(product_id)

            df.loc[index, "Matched_Product"] = details.get(
                "name",
                product_name
            )

            refs = details.get("refs") or {}

            df.loc[index, "Cardmarket_ID"] = refs.get(
                "cardmarketId"
            )

            english_average, listing_count, group_count = (
                calculate_english_average(details)
            )

            if english_average is None:

                df.loc[index, status_column] = (
                    "Geen Engelse prijzen gevonden"
                )

            else:

                df.loc[index, price_column] = round(
                    english_average,
                    2
                )

                df.loc[index, status_column] = "OK"

        except Exception as error:

            df.loc[index, status_column] = (
                f"FOUT: {error}"
            )

        progress.progress(
            (position + 1) / total
        )

    df.to_excel(
        EXCEL_FILE,
        index=False,
        engine="openpyxl"
    )

    return df


# ============================================================
# STREAMLIT INTERFACE
# ============================================================

st.set_page_config(
    page_title="Pokémon Booster Box Tracker",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Pokémon Booster Box Tracker")

st.write(
    "Cardmarket-prijzen voor Engelse Pokémon boosterboxen."
)


# ------------------------------------------------------------
# HUIDIGE DATA
# ------------------------------------------------------------

if not EXCEL_FILE.exists():

    st.error(
        "Map1.xlsx werd niet gevonden."
    )

    st.stop()


df_current = pd.read_excel(
    EXCEL_FILE,
    engine="openpyxl"
)


st.subheader("Boosterboxen")

st.dataframe(
    df_current,
    use_container_width=True,
    hide_index=True
)


# ------------------------------------------------------------
# UPDATEKNOP
# ------------------------------------------------------------
if st.button("Test PokeTrace API-key"):
    api_key = st.secrets["POKETRACE_API_KEY"]

    # Toon NIET de key zelf, alleen wat veilige diagnostiek
    st.write("Lengte API-key:", len(api_key))
    st.write("Begint met:", api_key[:3] + "...")

    response = requests.get(
        "https://api.poketrace.com/v1/auth/info",
        headers={
            "X-API-Key": api_key,
            "Accept": "application/json",
        },
        timeout=30,
    )

    st.write("HTTP status:", response.status_code)
    st.code(response.text)
    
st.divider()

if st.button(
    "🔄 Prijzen bijwerken",
    type="primary",
    use_container_width=True
):

    with st.spinner(
        "Cardmarket-prijzen worden opgehaald..."
    ):

        try:

            df_updated = update_prices()

            st.success(
                "Prijzen succesvol bijgewerkt."
            )

            st.dataframe(
                df_updated,
                use_container_width=True,
                hide_index=True
            )

        except Exception as error:

            st.error(
                f"Bijwerken mislukt: {error}"
            )


# ------------------------------------------------------------
# DOWNLOAD
# ------------------------------------------------------------

st.divider()

with open(EXCEL_FILE, "rb") as file:

    st.download_button(
        label="⬇️ Excel downloaden",
        data=file,
        file_name="Map1.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )
