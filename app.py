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

# ───────────────────────── HELPERS ─────────────────────────
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


# ───────────────────────── INPUT READER ─────────────────────────
def read_input_to_df(
    input_file,
    marketplace,
    general_header_row=1,
    sheet_name=None,
):
    """
    Read Excel input with safe, deterministic header handling.
    """

    marketplace_configs = {
        "Amazon":   {"sheet": "Template", "header_row": 2, "sheet_index": None},
        "Flipkart": {"sheet": None,       "header_row": 1, "sheet_index": 2},
        "Myntra":   {"sheet": None,       "header_row": 3, "sheet_index": 1},
        "Ajio":     {"sheet": None,       "header_row": 2, "sheet_index": 2},
        "TataCliq": {"sheet": None,       "header_row": 4, "sheet_index": 0},
        "General":  {"sheet": None,       "header_row": general_header_row, "sheet_index": 0},
    }

    config = marketplace_configs[marketplace]

    # ───────────── GENERAL ─────────────
    if marketplace == "General":
        src_df = pd.read_excel(
            input_file,
            sheet_name=sheet_name,
            header=config["header_row"] - 1,
        )

    # ───────────── FLIPKART (manual) ─────────────
    elif marketplace == "Flipkart":
        xl = pd.ExcelFile(input_file)
        raw_df = xl.parse(xl.sheet_names[config["sheet_index"]], header=None)

        headers = raw_df.iloc[config["header_row"] - 1].tolist()
        src_df = raw_df.iloc[config["header_row"] :].copy()
        src_df.columns = headers

    # ───────────── FIXED SHEET (Amazon) ─────────────
    elif config["sheet"] is not None:
        src_df = pd.read_excel(
            input_file,
            sheet_name=config["sheet"],
            header=config["header_row"] - 1,
        )

    # ───────────── INDEX BASED ─────────────
    else:
        xl = pd.ExcelFile(input_file)
        src_df = xl.parse(
            xl.sheet_names[config["sheet_index"]],
            header=config["header_row"] - 1,
        )

    src_df.dropna(axis=1, how="all", inplace=True)
    return src_df


# ───────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config(page_title="SKU Template Automation", layout="wide")
st.title("Rubick OS Template Conversion")

marketplace_options = [
    "General", "Amazon", "Flipkart",
    "Myntra", "Ajio", "TataCliq",
]

marketplace_type = st.selectbox("Select Template Type", marketplace_options)

general_header_row = 1

if marketplace_type == "General":
    st.info("Header row is optional. Default = 1.")
    general_header_row = st.number_input(
        "Header row (1-indexed)",
        min_value=1,
        value=1
    )

input_file = st.file_uploader(
    "Upload Input Excel File",
    type=["xlsx", "xls", "xlsm"],
)

if not input_file:
    st.stop()

selected_sheet = None
if marketplace_type == "General":
    xl = pd.ExcelFile(input_file)
    selected_sheet = st.selectbox("Select sheet", xl.sheet_names)

src_df = read_input_to_df(
    input_file=input_file,
    marketplace=marketplace_type,
    general_header_row=general_header_row,
    sheet_name=selected_sheet,
)

st.subheader("Preview")
st.dataframe(src_df.head(5), use_container_width=True)

# ───────────────────────── OUTPUT GENERATION ─────────────────────────
def process_file_stub(df: pd.DataFrame) -> BytesIO:
    """
    Placeholder for your actual process_file() logic.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


if st.button("Generate Output"):
    with st.spinner("Processing…"):
        result = process_file_stub(src_df)

    st.success("✅ Output Generated!")
    st.download_button(
        "📥 Download Output",
        data=result,
        file_name="output_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("Built for Rubick.ai | By Vishnu Sai")
