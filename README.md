# SKU Template Mapper

A Streamlit app to map product data into a standardized SKU template using a predefined field mapping.

## Features
- Accepts .xls / .xlsx / .xlsm input files
- Uses Mapping - Automation.xlsx to map fields automatically
- Ensures Option1 = Size and Option2 = Color
- Generates output template with:
  - Values tab → input data with mapped attributes
  - Type tab → headers, mapped fields, and unique column values
- Single-click workflow: upload → download

## Setup Instructions
1. Clone repo:
   git clone https://github.com/your-username/sku-mapper.git
   cd sku-mapper

2. Install dependencies:
   pip install -r requirements.txt

3. Add template & mapping files:
   - Place sku-template (4).xlsx in the root folder
   - Place Mapping - Automation.xlsx in the root folder

4. Run app:
   streamlit run app.py

5. Use the app:
   - Upload any Excel input file
   - Download processed SKU template directly

## Notes
- For .xls files, xlrd==1.2.0 is required (already included).
- Do not modify app.py unless you know the mapping logic.
