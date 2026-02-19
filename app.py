import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO
import os

# ───────────────────────── FILE PATHS ─────────────────────────
DEFAULT_TEMPLATE = "sku-template (4).xlsx"
FALLBACK_UPLOADED_TEMPLATE = "/mnt/data/output_template (62).xlsx"

if os.path.exists(FALLBACK_UPLOADED_TEMPLATE):
    TEMPLATE_PATH = FALLBACK_UPLOADED_TEMPLATE
else:
    TEMPLATE_PATH = DEFAULT_TEMPLATE

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
IMAGE_KEYWORDS = {"image", "img", "picture", "photo", "thumbnail", "thumb", "hero", "front", "back", "url"}

def is_image_column(col_header_norm: str, series: pd.Series) -> bool:
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0.0
    return header_hit or ratio >= 0.30
# ╰───────────────────────────────────────────────────────────╯

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

# ───────────────────────── INPUT READER ─────────────────────────
def read_input_to_df(input_file, marketplace, header_row=1, data_row=2, sheet_name=None):
    marketplace_configs = {
        "Amazon": {"sheet": "Template", "header_row": 4, "data_row": 7, "sheet_index": None},
        "Flipkart": {"sheet": None, "header_row": 1, "data_row": 5, "sheet_index": 2},
        "Myntra": {"sheet": None, "header_row": 3, "data_row": 4, "sheet_index": 1},
        "Ajio": {"sheet": None, "header_row": 2, "data_row": 3, "sheet_index": 2},
        "TataCliq": {"sheet": None, "header_row": 4, "data_row": 6, "sheet_index": 0},
        "General": {"sheet": None, "header_row": header_row, "data_row": data_row, "sheet_index": 0}
    }
    config = marketplace_configs.get(marketplace, marketplace_configs["General"])

    # ---------- AMAZON (ROBUST FIX) ----------
    if marketplace == "Amazon":
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse("Template", header=None)

        header_idx = config["header_row"] - 1   # Excel row 4
        data_idx = config["data_row"] - 1       # Excel row 7

        headers = temp_df.iloc[header_idx].tolist()
        src_df = temp_df.iloc[data_idx:].copy()
        src_df.columns = headers

    # ---------- FLIPKART ----------
    elif marketplace == "Flipkart":
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse(xl.sheet_names[config["sheet_index"]], header=None)

        header_idx = config["header_row"] - 1
        data_idx = config["data_row"] - 1

        headers = temp_df.iloc[header_idx].tolist()
        src_df = temp_df.iloc[data_idx:].copy()
        src_df.columns = headers

    # ---------- GENERAL WITH SELECTED SHEET ----------
    elif marketplace == "General" and sheet_name:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            sheet_name,
            header=config["header_row"] - 1,
            skiprows=config["data_row"] - config["header_row"] - 1
        )

    # ---------- OTHERS ----------
    else:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            xl.sheet_names[config["sheet_index"]],
            header=config["header_row"] - 1,
            skiprows=config["data_row"] - config["header_row"] - 1
        )

    src_df.dropna(axis=1, how="all", inplace=True)
    return src_df

# ------------------------- PROCESS FILE -------------------------
def process_file(
    input_file,
    marketplace: str,
    selected_variant_col: str | None = None,
    selected_product_col: str | None = None,
    general_header_row: int = 1,
    general_data_row: int = 2,
    general_sheet_name: str | None = None,
):
    src_df = read_input_to_df(
        input_file,
        marketplace,
        header_row=general_header_row,
        data_row=general_data_row,
        sheet_name=general_sheet_name
    )

    columns_meta = []
    for col in src_df.columns:
        dtype = "imageurlarray" if is_image_column(norm(col), src_df[col]) else "string"
        columns_meta.append({"src": col, "out": col, "row3": "mandatory", "row4": dtype})

    color_cols = [c for c in src_df.columns if "color" in norm(c) or "colour" in norm(c)]
    size_cols = [c for c in src_df.columns if "size" in norm(c)]

    option1_data = pd.Series([""] * len(src_df), dtype=str)
    option2_data = pd.Series([""] * len(src_df), dtype=str)

    if size_cols:
        option1_data = src_df[size_cols[0]].fillna("").astype(str)
        if color_cols and color_cols[0] != size_cols[0]:
            option2_data = src_df[color_cols[0]].fillna("").astype(str)
    elif color_cols:
        option2_data = src_df[color_cols[0]].fillna("").astype(str)

    unique_opt1 = option1_data.replace("", pd.NA).dropna().unique().tolist()
    unique_opt2 = option2_data.replace("", pd.NA).dropna().unique().tolist()

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals = wb["Values"]
    ws_types = wb["Types"]

    def first_empty_col(ws, header_rows=(1,)):
        for col in range(1, 201):
            if all(ws.cell(row=r, column=col).value in (None, "") for r in header_rows):
                return col
        return ws.max_column + 1

    vals_start_col = first_empty_col(ws_vals, (1,))
    types_start_col = first_empty_col(ws_types, (1, 2, 3, 4))

    for idx, meta in enumerate(columns_meta):
        vcol = vals_start_col + idx
        tcol = types_start_col + idx
        header = clean_header(meta["out"])

        ws_vals.cell(row=1, column=vcol, value=header)
        for r, val in enumerate(src_df[meta["src"]], start=2):
            cell = ws_vals.cell(row=r, column=vcol)
            cell.value = None if pd.isna(val) else str(val)
            cell.number_format = "@"

        ws_types.cell(row=1, column=tcol, value=header)
        ws_types.cell(row=2, column=tcol, value=header)
        ws_types.cell(row=3, column=tcol, value=meta["row3"])
        ws_types.cell(row=4, column=tcol, value=meta["row4"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ───────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("Rubick OS Template Conversion")

marketplace_type = st.selectbox(
    "Select Template Type",
    ["General", "Amazon", "Flipkart", "Myntra", "Ajio", "TataCliq", "Zivame", "Celio"]
)

input_file = st.file_uploader("Upload Input Excel File", type=["xlsx", "xls", "xlsm"])

if input_file:
    src_df = read_input_to_df(input_file, marketplace_type)
    st.subheader("Preview (first 5 rows)")
    st.dataframe(src_df.head(5))

    if st.button("Generate Output"):
        result = process_file(input_file, marketplace_type)
        st.download_button(
            "Download Output",
            data=result,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption("Built for Rubick.ai | By Vishnu Sai")
