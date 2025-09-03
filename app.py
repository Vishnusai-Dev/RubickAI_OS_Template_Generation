import streamlit as st
import pandas as pd
import openpyxl
import re
import os
from io import BytesIO

# ───────────────────────── FILE PATHS ─────────────────────────
TEMPLATE_PATH = "sku-template (4).xlsx"
MAPPING_PATH  = "Mapping - Automation.xlsx"

# ─────────────────── INTERNAL   COLUMN KEYS ───────────────────
ATTR_KEY   = "attributes"
TARGET_KEY = "fieldname"

# ───────────────────────── HELPERS ─────────────────────────
def read_excel_auto(file_obj):
    """Read Excel (xls, xlsx, xlsm) using correct engine."""
    ext = os.path.splitext(file_obj.name)[1].lower()
    file_bytes = file_obj.read()
    buffer = BytesIO(file_bytes)
    if ext == ".xls":
        engine = "xlrd"   # xlrd==1.2.0 must be installed
    else:
        engine = "openpyxl"
    return pd.ExcelFile(buffer, engine=engine)

def load_mapping():
    mapping_xl = pd.ExcelFile(MAPPING_PATH, engine="openpyxl")
    df = pd.read_excel(mapping_xl, sheet_name=0)
    df.columns = df.columns.str.strip().str.lower()
    if ATTR_KEY not in df.columns or TARGET_KEY not in df.columns:
        st.error("Mapping file must have 'attributes' and 'fieldname' columns.")
        st.stop()
    return df

def match_and_transform(input_df, mapping_df):
    """Row-wise match → Option1=size, Option2=color → keep all other columns."""
    input_df.columns = input_df.columns.str.strip().str.lower()
    mapping_df[ATTR_KEY] = mapping_df[ATTR_KEY].str.strip().str.lower()
    mapping_df[TARGET_KEY] = mapping_df[TARGET_KEY].str.strip().str.lower()

    output_df = input_df.copy()
    if "option1" not in output_df.columns:
        output_df["option1"] = ""
    if "option2" not in output_df.columns:
        output_df["option2"] = ""

    for idx, row in input_df.iterrows():
        for col in input_df.columns:
            attr = col.strip().lower()
            mapped_rows = mapping_df[mapping_df[ATTR_KEY] == attr]
            if not mapped_rows.empty:
                target_field = mapped_rows.iloc[0][TARGET_KEY]
                value = str(row[col]).strip()
                if "size" in target_field.lower():
                    output_df.at[idx, "option1"] = value
                elif "color" in target_field.lower():
                    output_df.at[idx, "option2"] = value
                else:
                    output_df.at[idx, col] = value
    return output_df

def fill_template(input_df, mapping_df):
    """Fill template → Values + Type tab (add unique values)."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    values_ws = wb["Values"]
    type_ws = wb["Type"]

    # Clear existing data
    for ws in [values_ws, type_ws]:
        for row in ws["A1:Z500"]:
            for cell in row:
                cell.value = None

    # ── Fill Values tab ──
    for c_idx, col_name in enumerate(input_df.columns, start=1):
        values_ws.cell(row=1, column=c_idx).value = col_name
    for r_idx, row in enumerate(input_df.values.tolist(), start=2):
        for c_idx, val in enumerate(row, start=1):
            values_ws.cell(row=r_idx, column=c_idx).value = val

    # ── Fill Type tab ──
    headers = input_df.columns.tolist()
    for c_idx, col_name in enumerate(headers, start=1):
        # Row1 & Row2 → headers
        type_ws.cell(row=1, column=c_idx).value = col_name
        type_ws.cell(row=2, column=c_idx).value = col_name

        # Row3 & Row4 → mapping
        mapped = mapping_df[mapping_df[ATTR_KEY] == col_name.lower()]
        if not mapped.empty:
            type_ws.cell(row=3, column=c_idx).value = mapped.iloc[0][TARGET_KEY]
            type_ws.cell(row=4, column=c_idx).value = mapped.iloc[0][TARGET_KEY]

        # Add unique values (starting row 5)
        unique_vals = input_df[col_name].dropna().unique().tolist()
        for r_offset, val in enumerate(unique_vals, start=5):
            type_ws.cell(row=r_offset, column=c_idx).value = val

    # Save to memory
    output_buffer = BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# ───────────────────────── STREAMLIT APP ─────────────────────────
st.set_page_config(page_title="Template Mapper", layout="wide")
st.title("Excel → SKU Template Mapper")

uploaded_file = st.file_uploader(
    "Upload any Excel file (XLS/XLSX/XLSM)",
    type=["xls", "xlsx", "xlsm"]
)

if uploaded_file:
    st.info("Processing file...")
    mapping_df = load_mapping()
    input_xl = read_excel_auto(uploaded_file)
    try:
        sheet_name = input_xl.sheet_names[0]
        input_df = pd.read_excel(input_xl, sheet_name=sheet_name)
        mapped_df = match_and_transform(input_df, mapping_df)
        output_buffer = fill_template(mapped_df, mapping_df)

        st.success("Template generated successfully! Click below to download.")
        st.download_button(
            label="Download SKU Template",
            data=output_buffer,
            file_name="sku_template_filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error while processing: {e}")
