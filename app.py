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

# ───────────────────────── HELPERS ─────────────────────────
def norm(s):
    if pd.isna(s):
        return ""
    return "".join(str(s).split()).lower()

def clean_header(header):
    if pd.isna(header):
        return ""
    header_str = str(header)
    header_str = re.sub(r"[^0-9A-Za-z ]+", " ", header_str)
    header_str = re.sub(r"\s+", " ", header_str).strip()
    return header_str

IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {"image", "img", "picture", "photo", "thumbnail", "thumb", "hero", "front", "back", "url"}

def is_image_column(col_header_norm, series):
    header_hit = any(k in col_header_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0.0
    return header_hit or ratio >= 0.30

# Marketplace → (productId, variantId)
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
    name = norm(name)
    for c in df.columns:
        if norm(c) == name or name in norm(c):
            return c
    return None

# ───────────────────────── INPUT READER ─────────────────────────
def read_input_to_df(input_file, marketplace, header_row=1, data_row=2, sheet_name=None):

    marketplace_configs = {
        # ✅ FIXED AMAZON CONFIG
        "Amazon": {"sheet": "Template", "header_row": 4, "data_row": 7, "sheet_index": None},

        "Flipkart": {"sheet": None, "header_row": 1, "data_row": 5, "sheet_index": 2},
        "Myntra": {"sheet": None, "header_row": 3, "data_row": 4, "sheet_index": 1},
        "Ajio": {"sheet": None, "header_row": 2, "data_row": 3, "sheet_index": 2},
        "TataCliq": {"sheet": None, "header_row": 4, "data_row": 6, "sheet_index": 0},
        "General": {"sheet": None, "header_row": header_row, "data_row": data_row, "sheet_index": 0}
    }

    config = marketplace_configs.get(marketplace, marketplace_configs["General"])

    if marketplace == "General" and sheet_name:
        xl = pd.ExcelFile(input_file)
        df = xl.parse(sheet_name, header=header_row - 1,
                      skiprows=data_row - header_row - 1)
    elif marketplace == "Flipkart":
        xl = pd.ExcelFile(input_file)
        temp = xl.parse(xl.sheet_names[config["sheet_index"]], header=None)
        headers = temp.iloc[config["header_row"] - 1]
        df = temp.iloc[config["data_row"] - 1:]
        df.columns = headers
    else:
        df = pd.read_excel(
            input_file,
            sheet_name=config["sheet"],
            header=config["header_row"] - 1,
            skiprows=config["data_row"] - config["header_row"] - 1
        )

    # 🔒 Drop junk Amazon columns
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    df.dropna(axis=1, how="all", inplace=True)

    return df

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
        header_row=general_header_row,
        data_row=general_data_row,
        sheet_name=general_sheet_name,
    )

    columns_meta = []
    for col in src_df.columns:
        dtype = "imageurlarray" if is_image_column(norm(col), src_df[col]) else "string"
        columns_meta.append({"src": col, "out": col, "row3": "mandatory", "row4": dtype})

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_vals = wb["Values"]
    ws_types = wb["Types"]

    def first_empty_col(ws, rows):
        for c in range(1, 300):
            if all(ws.cell(r, c).value in (None, "") for r in rows):
                return c
        return ws.max_column + 1

    vals_start = first_empty_col(ws_vals, (1,))
    types_start = first_empty_col(ws_types, (1, 2, 3, 4))

    for i, meta in enumerate(columns_meta):
        vcol = vals_start + i
        tcol = types_start + i

        header = clean_header(meta["out"])

        ws_vals.cell(1, vcol, header)
        for r, v in enumerate(src_df[meta["src"]], start=2):
            ws_vals.cell(r, vcol, None if pd.isna(v) else str(v)).number_format = "@"

        ws_types.cell(1, tcol, header)
        ws_types.cell(2, tcol, header)
        ws_types.cell(3, tcol, meta["row3"])
        ws_types.cell(4, tcol, meta["row4"])

    # ───────── Option columns ─────────
    size_cols = [c for c in src_df.columns if "size" in norm(c)]
    color_cols = [c for c in src_df.columns if "color" in norm(c)]

    opt1 = src_df[size_cols[0]] if size_cols else ""
    opt2 = src_df[color_cols[0]] if color_cols else ""

    for idx, (name, data) in enumerate([("Option 1", opt1), ("Option 2", opt2)]):
        v = vals_start + len(columns_meta) + idx
        t = types_start + len(columns_meta) + idx

        ws_vals.cell(1, v, name)
        ws_types.cell(1, t, name)
        ws_types.cell(2, t, name)
        ws_types.cell(3, t, "non mandatory")
        ws_types.cell(4, t, "select")

        if isinstance(data, pd.Series):
            for r, val in enumerate(data.fillna("").astype(str), start=2):
                ws_vals.cell(r, v, val if val else None)

    # ───────── variantId / productId ─────────
    if marketplace != "General":
        prod_name, var_name = MARKETPLACE_ID_MAP.get(marketplace, (None, None))
        prod_col = find_column_by_name_like(src_df, prod_name)
        var_col = find_column_by_name_like(src_df, var_name)

        start = vals_start + len(columns_meta) + 2

        for name, col in [("variantId", var_col), ("productId", prod_col)]:
            if col:
                ws_vals.cell(1, start, name)
                ws_types.cell(1, start, name)
                ws_types.cell(2, start, name)
                ws_types.cell(3, start, "mandatory")
                ws_types.cell(4, start, "string")

                for r, v in enumerate(src_df[col].fillna("").astype(str), start=2):
                    ws_vals.cell(r, start, v if v else None).number_format = "@"

                start += 1

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ───────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config("SKU Template Automation", layout="wide")
st.title("Rubick OS Template Conversion")

marketplace = st.selectbox(
    "Select Template Type",
    ["General", "Amazon", "Flipkart", "Myntra", "Ajio", "TataCliq", "Zivame", "Celio"]
)

input_file = st.file_uploader("Upload Input Excel File", ["xlsx", "xlsm", "xls"])

if input_file:
    if marketplace == "General":
        xl = pd.ExcelFile(input_file)
        sheet = st.selectbox("Select Sheet", xl.sheet_names)
    else:
        sheet = None

    df = read_input_to_df(input_file, marketplace, sheet_name=sheet)
    st.subheader("Preview (first 5 rows)")
    st.dataframe(df.head())

    if st.button("Generate Output"):
        out = process_file(input_file, marketplace, general_sheet_name=sheet)
        st.download_button(
            "📥 Download Output",
            data=out,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
