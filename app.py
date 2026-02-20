import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO
import os

# ───────────────────────── FILE PATHS ─────────────────────────
DEFAULT_TEMPLATE = "sku-template (4).xlsx"
FALLBACK_UPLOADED_TEMPLATE = "/mnt/data/output_template (62).xlsx"
TEMPLATE_PATH = FALLBACK_UPLOADED_TEMPLATE if os.path.exists(FALLBACK_UPLOADED_TEMPLATE) else DEFAULT_TEMPLATE

# ╭───────────────── HELPERS ─────────────────╮
def norm(s):
    if pd.isna(s):
        return ""
    return "".join(str(s).split()).lower()

def clean_header(h):
    if pd.isna(h):
        return ""
    h = re.sub(r"[^0-9A-Za-z ]+", " ", str(h))
    return re.sub(r"\s+", " ", h).strip()

def make_unique_headers(headers):
    seen = {}
    out = []
    for h in headers:
        h = clean_header(h)
        if h == "":
            h = "Unnamed"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        out.append(h)
    return out

IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {"image", "img", "picture", "photo", "thumbnail", "thumb", "hero", "front", "back", "url"}

def is_image_column(col_header_norm, series):
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0
    return header_hit or ratio >= 0.30
# ╰──────────────────────────────────────────╯

MARKETPLACE_ID_MAP = {
    "Amazon": ("Seller SKU", "Parent SKU"),
    "Myntra": ("styleId", "styleGroupId"),
    "Ajio": ("*Item SKU", "*Style Code"),
    "Flipkart": ("Seller SKU ID", "Style Code"),
    "TataCliq": ("Seller Article SKU", "*Style Code"),
    "Zivame": ("Style Code", "SKU Code"),
    "Celio": ("Style Code", "SKU Code"),
}

def find_column_by_name_like(df, name):
    if not name:
        return None
    name = str(name).strip()
    for c in df.columns:
        if str(c).strip() == name:
            return c
    n = norm(name)
    for c in df.columns:
        if norm(c) == n or n in norm(c):
            return c
    return None

# ───────────────────────── INPUT READER ─────────────────────────
def read_input_to_df(input_file, marketplace, header_row=1, data_row=2, sheet_name=None):
    configs = {
        "Amazon": {"sheet": "Template", "header_row": 4, "data_row": 7},
        "Flipkart": {"sheet_index": 2, "header_row": 1, "data_row": 5},
        "Myntra": {"sheet_index": 1, "header_row": 3, "data_row": 4},
        "Ajio": {"sheet_index": 2, "header_row": 2, "data_row": 3},
        "TataCliq": {"sheet_index": 0, "header_row": 4, "data_row": 6},
        "General": {"sheet_index": 0, "header_row": header_row, "data_row": data_row},
    }
    cfg = configs[marketplace]

    # ===== AMAZON (ROBUST, FIXED) =====
    if marketplace == "Amazon":
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse(cfg["sheet"], header=None)

        header_idx = cfg["header_row"] - 1   # Excel row 4
        data_idx = cfg["data_row"] - 1       # Excel row 7

        headers = make_unique_headers(temp_df.iloc[header_idx].tolist())
        src_df = temp_df.iloc[data_idx:].copy()
        src_df.columns = headers

    # ===== FLIPKART =====
    elif marketplace == "Flipkart":
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse(xl.sheet_names[cfg["sheet_index"]], header=None)
        headers = make_unique_headers(temp_df.iloc[cfg["header_row"] - 1].tolist())
        src_df = temp_df.iloc[cfg["data_row"] - 1:].copy()
        src_df.columns = headers

    # ===== GENERAL =====
    elif marketplace == "General" and sheet_name:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            sheet_name,
            header=cfg["header_row"] - 1,
            skiprows=cfg["data_row"] - cfg["header_row"] - 1
        )

    # ===== OTHERS =====
    else:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            xl.sheet_names[cfg["sheet_index"]],
            header=cfg["header_row"] - 1,
            skiprows=cfg["data_row"] - cfg["header_row"] - 1
        )

    src_df.dropna(axis=1, how="all", inplace=True)
    return src_df

# ───────────────────────── PROCESS FILE ─────────────────────────
def process_file(
    input_file,
    marketplace,
    selected_variant_col=None,
    selected_product_col=None,
    general_header_row=1,
    general_data_row=2,
    general_sheet_name=None,
):
    src_df = read_input_to_df(
        input_file,
        marketplace,
        general_header_row,
        general_data_row,
        general_sheet_name,
    )

    columns_meta = []
    for col in src_df.columns:
        dtype = "imageurlarray" if is_image_column(norm(col), src_df[col]) else "string"
        columns_meta.append({"src": col, "row3": "mandatory", "row4": dtype})

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals = wb["Values"]
    ws_types = wb["Types"]

    def first_empty_col(ws, rows):
        for c in range(1, 201):
            if all(ws.cell(row=r, column=c).value in (None, "") for r in rows):
                return c
        return ws.max_column + 1

    v_start = first_empty_col(ws_vals, (1,))
    t_start = first_empty_col(ws_types, (1, 2, 3, 4))

    for i, meta in enumerate(columns_meta):
        vcol = v_start + i
        tcol = t_start + i
        col = meta["src"]

        ws_vals.cell(row=1, column=vcol, value=col)
        for r, val in enumerate(src_df[col], start=2):
            cell = ws_vals.cell(row=r, column=vcol, value=None if pd.isna(val) else str(val))
            cell.number_format = "@"

        ws_types.cell(row=1, column=tcol, value=col)
        ws_types.cell(row=2, column=tcol, value=col)
        ws_types.cell(row=3, column=tcol, value=meta["row3"])
        ws_types.cell(row=4, column=tcol, value=meta["row4"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ───────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("Rubick OS Template Conversion")

marketplace = st.selectbox(
    "Select Template Type",
    ["General", "Amazon", "Flipkart", "Myntra", "Ajio", "TataCliq", "Zivame", "Celio"]
)

input_file = st.file_uploader("Upload Input Excel File", type=["xlsx", "xls", "xlsm"])

if input_file:
    src_df = read_input_to_df(input_file, marketplace)
    st.subheader("Preview (first 5 rows)")
    st.dataframe(src_df.head(5))

    if st.button("Generate Output"):
        result = process_file(input_file, marketplace)
        st.download_button(
            "📥 Download Output",
            data=result,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.caption("Built for Rubick.ai | By Vishnu Sai")
