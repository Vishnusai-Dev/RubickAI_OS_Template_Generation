import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO
import os

# ───────────────────────── FILE PATHS ─────────────────────────
DEFAULT_TEMPLATE = "sku-template (4).xlsx"
FALLBACK_UPLOADED_TEMPLATE = "/mnt/data/output_template (62).xlsx"

TEMPLATE_PATH = (
    FALLBACK_UPLOADED_TEMPLATE
    if os.path.exists(FALLBACK_UPLOADED_TEMPLATE)
    else DEFAULT_TEMPLATE
)

# ╭───────────────── NORMALISERS & HELPERS ─────────────────╮
def norm(s) -> str:
    if pd.isna(s):
        return ""
    return "".join(str(s).split()).lower()


def clean_header(header) -> str:
    if pd.isna(header):
        return ""
    header_str = str(header)
    header_str = re.sub(r"[^0-9A-Za-z ]+", " ", header_str)
    header_str = re.sub(r"\s+", " ", header_str).strip()
    return header_str


IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {
    "image", "img", "picture", "photo",
    "thumbnail", "thumb", "hero",
    "front", "back", "url"
}


def is_image_column(col_header_norm: str, series: pd.Series) -> bool:
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0.0
    return header_hit or ratio >= 0.30
# ╰───────────────────────────────────────────────────────────╯


# Marketplace → (productId header, variantId header)
MARKETPLACE_ID_MAP = {
    "Amazon": ("Seller SKU", "Parent SKU"),
    "Myntra": ("styleId", "styleGroupId"),
    "Ajio": ("*Item SKU", "*Style Code"),
    "Flipkart": ("Seller SKU ID", "Style Code"),
    "TataCliq": ("Seller Article SKU", "*Style Code"),
    "Zivame": ("Style Code", "SKU Code"),
    "Celio": ("Style Code", "SKU Code"),
}


def find_column_by_name_like(src_df: pd.DataFrame, name: str):
    if not name:
        return None

    name = str(name).strip()

    for c in src_df.columns:
        if str(c).strip() == name:
            return c

    nname = norm(name)
    for c in src_df.columns:
        if norm(c) == nname:
            return c

    for c in src_df.columns:
        if nname in norm(c):
            return c

    return None


def read_input_to_df(
    input_file,
    marketplace,
    header_row=1,
    data_row=2,
    sheet_name=None,
):
    """
    Read uploaded Excel into a DataFrame using marketplace config
    or supplied header/data rows (1-indexed).
    """
    marketplace_configs = {
        "Amazon": {"sheet": "Template", "header_row": 2, "data_row": 7, "sheet_index": None},
        "Flipkart": {"sheet": None, "header_row": 1, "data_row": 5, "sheet_index": 2},
        "Myntra": {"sheet": None, "header_row": 3, "data_row": 4, "sheet_index": 1},
        "Ajio": {"sheet": None, "header_row": 2, "data_row": 3, "sheet_index": 2},
        "TataCliq": {"sheet": None, "header_row": 4, "data_row": 6, "sheet_index": 0},
        "General": {"sheet": None, "header_row": header_row, "data_row": data_row, "sheet_index": 0},
    }

    config = marketplace_configs.get(marketplace, marketplace_configs["General"])

    if marketplace == "General" and sheet_name:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            sheet_name,
            header=header_row - 1,
            skiprows=data_row - header_row - 1,
        )

    elif marketplace == "Flipkart":
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse(xl.sheet_names[config["sheet_index"]], header=None)
        headers = temp_df.iloc[config["header_row"] - 1].tolist()
        src_df = temp_df.iloc[config["data_row"] - 1:].copy()
        src_df.columns = headers

    elif config["sheet"] is not None:
        src_df = pd.read_excel(
            input_file,
            sheet_name=config["sheet"],
            header=config["header_row"] - 1,
            skiprows=config["data_row"] - config["header_row"] - 1,
        )

    else:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            xl.sheet_names[config["sheet_index"]],
            header=config["header_row"] - 1,
            skiprows=config["data_row"] - config["header_row"] - 1,
        )

    src_df.dropna(axis=1, how="all", inplace=True)
    return src_df


# ───────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("Rubick OS Template Conversion")

if os.path.exists(TEMPLATE_PATH):
    st.info(f"Using template: {os.path.basename(TEMPLATE_PATH)}")

    try:
        with open(TEMPLATE_PATH, "rb") as f:
            st.download_button(
                "Download current template (for reference)",
                data=f.read(),
                file_name=os.path.basename(TEMPLATE_PATH),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception:
        pass


marketplace_options = [
    "General", "Amazon", "Flipkart",
    "Myntra", "Ajio", "TataCliq",
    "Zivame", "Celio",
]

marketplace_type = st.selectbox("Select Template Type", marketplace_options)

general_header_row = 1
general_data_row = 2

if marketplace_type == "General":
    st.info(
        "If header/data rows are left as default, "
        "Header row = 1 and Data row = 2."
    )
    general_header_row = st.number_input(
        "Header row (1-indexed)", min_value=1, value=1
    )
    general_data_row = st.number_input(
        "Data row (1-indexed)", min_value=1, value=2
    )

input_file = st.file_uploader(
    "Upload Input Excel File",
    type=["xlsx", "xls", "xlsm"],
)

if not input_file:
    st.info("Upload a file to enable processing.")
    st.stop()

selected_sheet = None
if marketplace_type == "General":
    xl = pd.ExcelFile(input_file)
    selected_sheet = st.selectbox("Select sheet", xl.sheet_names)

src_df = read_input_to_df(
    input_file,
    marketplace_type,
    header_row=general_header_row,
    data_row=general_data_row,
    sheet_name=selected_sheet,
)

st.subheader("Preview")
st.dataframe(src_df.head(5))

if st.button("Generate Output"):
    with st.spinner("Processing…"):
        result = process_file(
            input_file,
            marketplace_type,
            selected_variant_col=None,
            selected_product_col=None,
            general_header_row=general_header_row,
            general_data_row=general_data_row,
            general_sheet_name=selected_sheet,
        )

    if result:
        st.success("✅ Output Generated!")
        st.download_button(
            "📥 Download Output",
            data=result,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.caption("Built for Rubick.ai | By Vishnu Sai")
