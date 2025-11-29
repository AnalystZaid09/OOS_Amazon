import streamlit as st
import pandas as pd
import io

# Page configuration
st.set_page_config(
    page_title="OOS Amazon Analysis",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for colored cells
st.markdown("""
<style>
    .stApp {
        max-width: 100%;
    }
    .doc-legend {
        display: flex;
        gap: 15px;
        flex-wrap: wrap;
        margin: 20px 0;
    }
    .legend-item {
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .legend-box {
        width: 30px;
        height: 30px;
        border: 1px solid #ddd;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to color DOC values
def color_doc(val):
    """Apply color based on DOC value"""
    try:
        doc = float(val)
        if 0 <= doc < 7:
            return 'background-color: #FFE5E5'  # Light Red
        elif 7 <= doc < 15:
            return 'background-color: #FFF9E5'  # Light Yellow
        elif 15 <= doc < 30:
            return 'background-color: #E5F9E5'  # Light Green
        elif 30 <= doc < 45:
            return 'background-color: #E5F3FF'  # Light Blue
        elif 45 <= doc < 60:
            return 'background-color: #F0F0F0'  # Light Grey
        elif 60 <= doc < 90:
            return 'background-color: #FFFFFF'  # White
        else:
            return 'background-color: #FFFFFF'
    except:
        return ''

# Title and description
st.title("📊 Inventory Analysis Dashboard")
st.markdown("Upload your files and analyze inventory with Days of Coverage (DOC) calculations")

# Sidebar for configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    st.markdown("### Number of Days for Analysis")
    no_of_days = st.number_input(
        "Enter the number of days",
        min_value=1,
        value=31,
        step=1,
        help="This represents the time period for your sales data. The DRR (Daily Run Rate) will be calculated by dividing Total Order Items by this number."
    )
    
    st.info("""
    **What is Number of Days?**
    
    This is the time period covered by your sales data (e.g., 30 or 31 for monthly data).
    
    - **DRR** (Daily Run Rate) = Total Order Items ÷ Days
    - **DOC** (Days of Coverage) = Fulfillable Qty ÷ DRR
    
    DOC tells you how many days your current inventory will last at the current sales rate.
    """)

# File upload section
st.header("📁 Upload Files")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Business Report")
    business_file = st.file_uploader(
        "Upload Business Report CSV",
        type=['csv'],
        key="business",
        help="Upload the BusinessReport CSV file"
    )

with col2:
    st.subheader("PM Data")
    pm_file = st.file_uploader(
        "Upload PM Excel/CSV",
        type=['xlsx', 'csv'],
        key="pm",
        help="Upload the PM.xlsx or converted CSV file"
    )

with col3:
    st.subheader("Inventory Data")
    inventory_file = st.file_uploader(
        "Upload Manage Inventory CSV",
        type=['csv'],
        key="inventory",
        help="Upload the Manage Inventory CSV file"
    )

# DOC Color Legend
st.header("🎨 DOC Color Legend")
st.markdown("""
<div class="doc-legend">
    <div class="legend-item">
        <div class="legend-box" style="background-color: #FFE5E5;"></div>
        <span><b>0-7 days</b> (Critical)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #FFF9E5;"></div>
        <span><b>7-15 days</b> (Low)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #E5F9E5;"></div>
        <span><b>15-30 days</b> (Good)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #E5F3FF;"></div>
        <span><b>30-45 days</b> (Optimal)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #F0F0F0;"></div>
        <span><b>45-60 days</b> (High)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #FFFFFF; border: 1px solid #ccc;"></div>
        <span><b>60-90 days</b> (Excess)</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# Process button
if st.button("🚀 Process Data", type="primary", use_container_width=True):
    if business_file is None or pm_file is None or inventory_file is None:
        st.error("⚠️ Please upload all three files before processing!")
    elif no_of_days <= 0:
        st.error("⚠️ Number of days must be greater than 0!")
    else:
        with st.spinner("Processing data..."):
            try:
                # Read files
                st.info("📖 Reading files...")
                original = pd.read_csv(business_file)
                
                # ✅ Ensure SKU in Business Report is string
                if "SKU" in original.columns:
                    original["SKU"] = original["SKU"].astype(str)
                
                # Read PM file (Excel or CSV)
                if pm_file.name.endswith('.xlsx'):
                    pm = pd.read_excel(pm_file)
                else:
                    pm = pd.read_csv(pm_file)
                
                inventory = pd.read_csv(inventory_file)

                # ✅ Ensure first column of Inventory (SKU-like) is string
                inventory.columns = inventory.columns.str.strip()
                inventory.iloc[:, 0] = inventory.iloc[:, 0].astype(str)

                # Process PM data - select columns C, D, E, F, G (indices 2-6)
                pm = pm.iloc[:, 2:7]
                pm.columns = ["Amazon Sku Name", "D", "Brand Manager", "F", "Brand"]
                
                # ✅ Ensure Amazon Sku Name is string
                pm["Amazon Sku Name"] = pm["Amazon Sku Name"].astype(str)

                # Merge Brand Manager
                original = original.merge(
                    pm[["Amazon Sku Name", "Brand Manager"]],
                    how="left",
                    left_on="SKU",
                    right_on="Amazon Sku Name"
                )
                
                # Insert Brand Manager column
                insert_pos = original.columns.get_loc("Title")
                col = original.pop("Brand Manager")
                original.insert(insert_pos, "Brand Manager", col)
                
                # Merge Brand
                original = original.merge(
                    pm[["Amazon Sku Name", "Brand"]],
                    how="left",
                    left_on="SKU",
                    right_on="Amazon Sku Name"
                )
                
                # Insert Brand column
                insert_pos = original.columns.get_loc("Title")
                col = original.pop("Brand")
                original.insert(insert_pos, "Brand", col)
                
                # Strip whitespace from inventory columns (already done above)
                inventory.columns = inventory.columns.str.strip()
                
                # Add fulfillable quantity (11th column, index 10)
                return_col = inventory.columns[10]
                mi_map = inventory.set_index(inventory.columns[0])[return_col]
                original["afn-fulfillable-quantity"] = original["SKU"].map(mi_map)
                
                # Add reserved quantity (13th column, index 12)
                return_col_13 = inventory.columns[12]
                mi_res_map = inventory.set_index(inventory.columns[0])[return_col_13]
                original["afn-reserved-quantity"] = original["SKU"].map(mi_res_map)
                
                # Clean Total Order Items and calculate DRR
                original["Total Order Items"] = (
                    original["Total Order Items"]
                    .astype(str)
                    .str.replace("\u00A0", "", regex=False)
                    .str.replace(",", "", regex=False)
                    .str.replace(r"[^\d\.\-]", "", regex=True)
                )
                original["Total Order Items"] = pd.to_numeric(original["Total Order Items"], errors="coerce")
                
                # Calculate DRR
                original["DRR"] = (original["Total Order Items"] / no_of_days).round(2)
                
                # Convert fulfillable quantity to numeric
                original["afn-fulfillable-quantity"] = pd.to_numeric(
                    original["afn-fulfillable-quantity"], errors="coerce"
                )
                
                # Calculate DOC
                original["DOC"] = (original["afn-fulfillable-quantity"] / original["DRR"]).round(2)
                
                # Replace inf with 0
                original["DOC"] = original["DOC"].replace([float('inf'), float('-inf')], 0)
                
                st.success("✅ Data processed successfully!")
                
                # Display results
                st.header("📈 Processed Results")
                
                # Show summary statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Products", len(original))
                with col2:
                    critical_stock = len(original[original["DOC"] < 7])
                    st.metric("Critical Stock (< 7 days)", critical_stock)
                with col3:
                    avg_doc = original["DOC"].mean()
                    st.metric("Average DOC", f"{avg_doc:.2f} days")
                with col4:
                    total_orders = original["Total Order Items"].sum()
                    st.metric("Total Orders", f"{total_orders:,.0f}")
                
                st.markdown("---")
                
                # Select key columns for display
                display_cols = [
                    "(Child) ASIN", "Brand", "Brand Manager", "SKU", "Title",
                    "Units Ordered", "afn-fulfillable-quantity", "afn-reserved-quantity",
                    "DRR", "DOC"
                ]
                
                # Filter columns that exist
                display_cols = [col for col in display_cols if col in original.columns]
                display_df = original[display_cols].copy()
                
                # Apply styling to DOC column
                styled_df = display_df.style.map(
                    color_doc,
                    subset=['DOC']
                )
                
                # Display the dataframe
                st.dataframe(
                    styled_df,
                    use_container_width=True,
                    height=600
                )
                
                # Download buttons
                st.markdown("---")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Convert to CSV for download
                    csv_buffer = io.StringIO()
                    original.to_csv(csv_buffer, index=False)
                    csv_data = csv_buffer.getvalue()
                    
                    st.download_button(
                        label="📥 Download as CSV",
                        data=csv_data,
                        file_name=f"processed_inventory_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    # Convert to Excel with conditional formatting
                    from openpyxl import Workbook
                    from openpyxl.styles import PatternFill
                    from openpyxl.utils.dataframe import dataframe_to_rows
                    
                    excel_buffer = io.BytesIO()
                    
                    # Create workbook and sheet
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Inventory Analysis"
                    
                    # Write dataframe to worksheet
                    for r_idx, row in enumerate(dataframe_to_rows(original, index=False, header=True), 1):
                        for c_idx, value in enumerate(row, 1):
                            ws.cell(row=r_idx, column=c_idx, value=value)
                    
                    # Find DOC column index
                    doc_col_idx = None
                    for idx, cell in enumerate(ws[1], 1):
                        if cell.value == 'DOC':
                            doc_col_idx = idx
                            break
                    
                    # Apply conditional formatting to DOC column
                    if doc_col_idx:
                        for row_idx in range(2, ws.max_row + 1):
                            cell = ws.cell(row=row_idx, column=doc_col_idx)
                            try:
                                doc_value = float(cell.value) if cell.value else 0
                                
                                if 0 <= doc_value < 7:
                                    fill = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
                                elif 7 <= doc_value < 15:
                                    fill = PatternFill(start_color="FFF9E5", end_color="FFF9E5", fill_type="solid")
                                elif 15 <= doc_value < 30:
                                    fill = PatternFill(start_color="E5F9E5", end_color="E5F9E5", fill_type="solid")
                                elif 30 <= doc_value < 45:
                                    fill = PatternFill(start_color="E5F3FF", end_color="E5F3FF", fill_type="solid")
                                elif 45 <= doc_value < 60:
                                    fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
                                elif 60 <= doc_value < 90:
                                    fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                                else:
                                    fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                                
                                cell.fill = fill
                            except (ValueError, TypeError):
                                pass
                    
                    # Auto-adjust column widths
                    for column in ws.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        ws.column_dimensions[column_letter].width = adjusted_width
                    
                    # Save to buffer
                    wb.save(excel_buffer)
                    excel_data = excel_buffer.getvalue()
                    
                    st.download_button(
                        label="📥 Download as Excel (with colors)",
                        data=excel_data,
                        file_name=f"processed_inventory_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                # Store in session state for persistence
                st.session_state['processed_data'] = original
                
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.exception(e)

# Show previously processed data if available
elif 'processed_data' in st.session_state:
    st.header("📈 Previously Processed Results")
    
    original = st.session_state['processed_data']
    
    # Show summary statistics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Products", len(original))
    with col2:
        critical_stock = len(original[original["DOC"] < 7])
        st.metric("Critical Stock (< 7 days)", critical_stock)
    with col3:
        avg_doc = original["DOC"].mean()
        st.metric("Average DOC", f"{avg_doc:.2f} days")
    with col4:
        total_orders = original["Total Order Items"].sum()
        st.metric("Total Orders", f"{total_orders:,.0f}")
    
    st.markdown("---")
    
    # Select key columns for display
    display_cols = [
        "(Child) ASIN", "Brand", "Brand Manager", "SKU", "Title",
        "Units Ordered", "afn-fulfillable-quantity", "afn-reserved-quantity",
        "DRR", "DOC"
    ]
    
    # Filter columns that exist
    display_cols = [col for col in display_cols if col in original.columns]
    display_df = original[display_cols].copy()
    
    # Apply styling to DOC column
    styled_df = display_df.style.applymap(
        color_doc,
        subset=['DOC']
    )
    
    # Display the dataframe
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=600
    )
    
    # Download buttons
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        csv_buffer = io.StringIO()
        original.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        
        st.download_button(
            label="📥 Download as CSV",
            data=csv_data,
            file_name=f"processed_inventory_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col2:
        # Convert to Excel with conditional formatting
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        excel_buffer = io.BytesIO()
        
        # Create workbook and sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Inventory Analysis"
        
        # Write dataframe to worksheet
        for r_idx, row in enumerate(dataframe_to_rows(original, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        
        # Find DOC column index
        doc_col_idx = None
        for idx, cell in enumerate(ws[1], 1):
            if cell.value == 'DOC':
                doc_col_idx = idx
                break
        
        # Apply conditional formatting to DOC column
        if doc_col_idx:
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=doc_col_idx)
                try:
                    doc_value = float(cell.value) if cell.value else 0
                    
                    if 0 <= doc_value < 7:
                        fill = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
                    elif 7 <= doc_value < 15:
                        fill = PatternFill(start_color="FFF9E5", end_color="FFF9E5", fill_type="solid")
                    elif 15 <= doc_value < 30:
                        fill = PatternFill(start_color="E5F9E5", end_color="E5F9E5", fill_type="solid")
                    elif 30 <= doc_value < 45:
                        fill = PatternFill(start_color="E5F3FF", end_color="E5F3FF", fill_type="solid")
                    elif 45 <= doc_value < 60:
                        fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
                    elif 60 <= doc_value < 90:
                        fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    else:
                        fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    
                    cell.fill = fill
                except (ValueError, TypeError):
                    pass
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Save to buffer
        wb.save(excel_buffer)
        excel_data = excel_buffer.getvalue()
        
        st.download_button(
            label="📥 Download as Excel (with colors)",
            data=excel_data,
            file_name=f"processed_inventory_analysis_%s.xlsx" % pd.Timestamp.now().strftime('%Y%m%d_%H%M%S'),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>Inventory Analysis Dashboard | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)
