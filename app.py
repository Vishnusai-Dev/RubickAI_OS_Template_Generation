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
    header = str(header)
    header = re.sub(r"[^0-9A-Za-z ]+", " ", header)
    header = re.sub(r"\s+", " ", header).strip()
    return header

IMAGE_EXT_RE = re.compile(r"(?i)\.(jpe?g|png|gif|bmp|webp|tiff?)$")
IMAGE_KEYWORDS = {"image", "img", "picture", "photo", "thumbnail", "thumb", "hero", "front", "back", "url"}

def is_image_column(col_norm, series):
    header_hit = any(k in col_norm for k in IMAGE_KEYWORDS)
    sample = series.dropna().astype(str).head(20)
    ratio = sample.str.contains(IMAGE_EXT_RE).mean() if not sample.empty else 0
    return header_hit or ratio >= 0.3

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
    n = norm(name)
    for c in df.columns:
        if n == norm(c) or n in norm(c):
            return c
    return None

# ───────────────────────── INPUT READER ─────────────────────────
def read_input_to_df(input_file, marketplace, header_row=1, data_row=2, sheet_name=None):

    if marketplace == "Amazon":
        # ✅ SAFE AMAZON PARSER (NO skiprows, NO header guessing)
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse("Template", header=None)

        header_idx = 3   # Excel row 4
        data_idx = 6     # Excel row 7

        headers = temp_df.iloc[header_idx].tolist()
        df = temp_df.iloc[data_idx:].copy()
        df.columns = headers

    elif marketplace == "Flipkart":
        xl = pd.ExcelFile(input_file)
        temp_df = xl.parse(xl.sheet_names[2], header=None)
        headers = temp_df.iloc[0].tolist()
        df = temp_df.iloc[4:].copy()
        df.columns = headers

    else:
        marketplace_configs = {
            "Myntra": (1, 3, 4),
            "Ajio": (2, 2, 3),
            "TataCliq": (0, 4, 6),
            "General": (0, header_row, data_row),
        }

        sheet_idx, hdr, data = marketplace_configs.get(marketplace, marketplace_configs["General"])
        xl = pd.ExcelFile(input_file)
        df = xl.parse(
            xl.sheet_names[sheet_idx],
            header=hdr - 1,
            skiprows=data - hdr - 1
        )

    # 🔒 Clean junk columns
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

    v_start = first_empty_col(ws_vals, (1,))
    t_start = first_empty_col(ws_types, (1, 2, 3, 4))

    for i, meta in enumerate(columns_meta):
        vcol, tcol = v_start + i, t_start + i
        header = clean_header(meta["out"])

        ws_vals.cell(1, vcol, header)
        for r, v in enumerate(src_df[meta["src"]], start=2):
            ws_vals.cell(r, vcol, None if pd.isna(v) else str(v)).number_format = "@"

        ws_types.cell(1, tcol, header)
        ws_types.cell(2, tcol, header)
        ws_types.cell(3, tcol, meta["row3"])
        ws_types.cell(4, tcol, meta["row4"])

    # ───────── Options ─────────
    size_col = next((c for c in src_df.columns if "size" in norm(c)), None)
    color_col = next((c for c in src_df.columns if "color" in norm(c)), None)

    for idx, (name, col) in enumerate([("Option 1", size_col), ("Option 2", color_col)]):
        v, t = v_start + len(columns_meta) + idx, t_start + len(columns_meta) + idx
        ws_vals.cell(1, v, name)
        ws_types.cell(1, t, name)
        ws_types.cell(2, t, name)
        ws_types.cell(3, t, "non mandatory")
        ws_types.cell(4, t, "select")

        if col:
            for r, val in enumerate(src_df[col].fillna("").astype(str), start=2):
                ws_vals.cell(r, v, val if val else None)

    # ───────── variantId / productId ─────────
    if marketplace != "General":
        prod_name, var_name = MARKETPLACE_ID_MAP.get(marketplace, (None, None))
        prod_col = find_column_by_name_like(src_df, prod_name)
        var_col = find_column_by_name_like(src_df, var_name)

        start = v_start + len(columns_meta) + 2

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
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("Rubick OS Template Conversion")

marketplace = st.selectbox(
    "Select Template Type",
    ["General", "Amazon", "Flipkart", "Myntra", "Ajio", "TataCliq", "Zivame", "Celio"]
)

input_file = st.file_uploader("Upload Input Excel File", ["xlsx", "xls", "xlsm"])

if input_file:
    df = read_input_to_df(input_file, marketplace)
    st.subheader("Preview (first 5 rows)")
    st.dataframe(df.head())

    if st.button("Generate Output"):
        output = process_file(input_file, marketplace)
        st.success("✅ Output Generated")
        st.download_button(
            "📥 Download Output",
            data=output,
            file_name="output_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload a file to begin.")

st.caption("Built for Rubick.ai | Vishnu Sai G V")
