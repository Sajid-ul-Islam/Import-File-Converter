import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="WooCommerce Export to Import Converter", layout="wide")

# Title and description
st.title("🔄 Export to Import Format Converter")
st.markdown("Upload your raw `Export.xlsx` file. This app will filter and format the columns to match the standard `Import.xlsx` structure.")

# The exact columns expected in the Import file
IMPORT_COLUMNS = [
    'ID', 'Type', 'SKU', 'GTIN, UPC, EAN, or ISBN', 'Name', 'Published', 'Is featured?', 
    'Visibility in catalog', 'Short description', 'Description', 'Date sale price starts', 
    'Date sale price ends', 'Tax status', 'Tax class', 'In stock?', 'Stock', 'Low stock amount', 
    'Backorders allowed?', 'Sold individually?', 'Weight (kg)', 'Length (cm)', 'Width (cm)', 
    'Height (cm)', 'Allow customer reviews?', 'Purchase note', 'Sale price', 'Regular price', 
    'Categories', 'Tags', 'Shipping class', 'Images', 'Download limit', 'Download expiry days', 
    'Parent', 'Grouped products', 'Upsells', 'Cross-sells', 'External URL', 'Button text', 
    'Position', 'Brands', 'Attribute 1 name', 'Attribute 1 value(s)', 'Attribute 1 visible', 
    'Attribute 1 global', 'Attribute 2 name', 'Attribute 2 value(s)', 'Attribute 2 visible', 
    'Attribute 2 global', 'Attribute 3 name', 'Attribute 3 value(s)', 'Attribute 3 visible', 
    'Attribute 3 global', 'Attribute 4 name', 'Attribute 4 value(s)', 'Attribute 4 visible', 
    'Attribute 4 global', 'Meta: _wp_page_template', 'Meta: size_and_fit', 'Meta: _size_and_fit', 
    'Meta: delivery_and_return', 'Meta: _delivery_and_return', 'Meta: site-sidebar-layout', 
    'Meta: ast-site-content-layout', 'Meta: site-content-style', 'Meta: site-sidebar-style', 
    'Meta: theme-transparent-header-meta', 'Meta: astra-migrate-meta-layouts'
]

# File uploader
uploaded_file = st.file_uploader("Upload Export File (Excel)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Load the uploaded file
        st.info("Loading file...")
        df = pd.read_excel(uploaded_file)
        
        # 1. Handle BOM in ID column if it exists
        if '\ufeffID' in df.columns:
            df = df.rename(columns={'\ufeffID': 'ID'})
            
        # 2. Process columns to match IMPORT_COLUMNS
        # Create a new dataframe that matches the import format perfectly
        st.info("Processing data...")
        df_out = pd.DataFrame()
        
        for col in IMPORT_COLUMNS:
            if col in df.columns:
                df_out[col] = df[col]
            else:
                # If the column doesn't exist in the export, create it as empty
                df_out[col] = None 
                
        # 3. Show a preview of the processed data
        st.success("File processed successfully!")
        st.subheader("Preview of Data to be Imported:")
        st.dataframe(df_out.head())
        
        # 4. Generate Excel file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_out.to_excel(writer, index=False, sheet_name='Sheet1')
        
        processed_data = output.getvalue()
        
        # Download button
        st.download_button(
            label="⬇️ Download Import Ready File",
            data=processed_data,
            file_name=uploaded_file.name.replace("Export", "Import"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
