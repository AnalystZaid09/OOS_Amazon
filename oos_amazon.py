# app.py - Streamlit-ready (Template-first pivot; fallback workbook)
import streamlit as st
import pandas as pd
import io
import os
import traceback
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OOS Amazon Analysis", page_icon="📊", layout="wide")

# -------------------------
# CSS + DOC legend
# -------------------------
st.markdown(
    """
<style>
    .stApp { max-width: 100%; }
    .doc-legend { display: flex; gap: 15px; flex-wrap: wrap; margin: 20px 0; }
    .legend-item { display: flex; align-items: center; gap: 8px; }
    .legend-box { width: 30px; height: 30px; border: 1px solid #ddd; border-radius: 4px; }
</style>
""",
    unsafe_allow_html=True,
)

st.header("🎨 DOC Color Legend")
st.markdown(
    """
<div class="doc-legend">
    <div class="legend-item"><div class="legend-box" style="background-color: #8B0000;"></div><span><b>0-7 days</b> (Critical)</span></div>
    <div class="legend-item"><div class="legend-box" style="background-color: #FF8C00;"></div><span><b>7-15 days</b> (Low)</span></div>
    <div class="legend-item"><div class="legend-box" style="background-color: #006400;"></div><span><b>15-30 days</b> (Good)</span></div>
    <div class="legend-item"><div class="legend-box" style="background-color: #B8860B;"></div><span><b>30-45 days</b> (Optimal)</span></div>
    <div class="legend-item"><div class="legend-box" style="background-color: #0F52BA;"></div><span><b>45-60 days</b> (High)</span></div>
    <div class="legend-item"><div class="legend-box" style="background-color: #8B4513;"></div><span><b>60-90 days</b> (Excess)</span></div>
</div>
""",
    unsafe_allow_html=True,
)

# -------------------------
# Helpers
# -------------------------
def color_doc(val):
    try:
        doc = float(val)
        if 0 <= doc < 7:
            return "background-color: #8B0000; color: white"
        elif 7 <= doc < 15:
            return "background-color: #FF8C00; color: white"
        elif 15 <= doc < 30:
            return "background-color: #006400; color: white"
        elif 30 <= doc < 45:
            return "background-color: #B8860B; color: white"
        elif 45 <= doc < 60:
            return "background-color: #0F52BA; color: white"
        elif 60 <= doc < 90:
            return "background-color: #8B4513; color: white"
        else:
            return "background-color: #222222; color: white"
    except:
        return ""

def filter_oos(df: pd.DataFrame) -> pd.DataFrame:
    """OOS: sirf afn-fulfillable-quantity == 0"""
    col = "afn-fulfillable-quantity"
    if col not in df.columns:
        return df.iloc[0:0].copy()
    qty = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df[qty == 0].copy()

def filter_overstock(df: pd.DataFrame, threshold: float = 90.0) -> pd.DataFrame:
    """Overstock: DOC >= threshold (default 90)"""
    if "DOC" not in df.columns:
        return df.iloc[0:0].copy()
    doc = pd.to_numeric(df["DOC"], errors="coerce")
    return df[doc >= threshold].copy()

def create_excel_with_doc_format(df: pd.DataFrame) -> bytes:
    """
    Create an XLSX bytes object from df and apply DOC conditional formatting
    matching the app's color buckets. Returns bytes.
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    # Write header + rows (preserve types where possible)
    header = list(df.columns)
    ws.append(header)
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))

    # Find DOC column index (case-insensitive contains 'doc')
    header_row = [cell.value for cell in ws[1]]
    doc_idx = None
    for idx, h in enumerate(header_row, start=1):
        if h and "doc" in str(h).strip().lower():
            doc_idx = idx
            break

    # Color map (same as app)
    col_map = {
        "dark_red": "8B0000",
        "dark_orange": "FF8C00",
        "dark_green": "006400",
        "golden": "B8860B",
        "sky": "0F52BA",
        "saddle": "8B4513",
        "white": "FFFFFF",
    }

    if doc_idx is not None:
        try:
            last_row = ws.max_row
            col_letter = get_column_letter(doc_idx)
            first_data_row = 2
            rng = f"{col_letter}{first_data_row}:{col_letter}{last_row}"

            rules = [
                CellIsRule(operator='between', formula=['0','6.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["dark_red"], end_color=col_map["dark_red"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['7','14.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["dark_orange"], end_color=col_map["dark_orange"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['15','29.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["dark_green"], end_color=col_map["dark_green"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['30','44.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["golden"], end_color=col_map["golden"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['45','59.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["sky"], end_color=col_map["sky"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['60','89.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["saddle"], end_color=col_map["saddle"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
            ]

            for rrule in rules:
                ws.conditional_formatting.add(rng, rrule)
        except Exception:
            pass

    # Auto-adjust widths
    try:
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[column_letter].width = min(max_length + 2, 50)
    except Exception:
        pass

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

def fill_template_and_get_bytes(template_path: str, df: pd.DataFrame, table_name: str = "DataTable"):
    """
    Load an Excel template (xlsx/xlsm) and ensure there's an Excel Table named `table_name`.
    - If the table exists: replace header+rows and update the table.ref.
    - If the table is missing: create a new sheet named table_name or table_name_1 and add a Table.
    Additionally apply conditional formatting (CellIs rules) to the DOC column.
    Returns BytesIO of the modified workbook.
    """
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill, Font

    # keep_vba=True so macros are preserved when template is xlsm
    wb = load_workbook(template_path, keep_vba=True)
    table_sheet = None
    table_obj = None

    # 1) try to find existing table named table_name
    for ws in wb.worksheets:
        for tbl in ws._tables:
            if tbl.displayName == table_name:
                table_sheet = ws
                table_obj = tbl
                break
        if table_obj:
            break

    # helper: parse A1 addresses
    def cell_to_rowcol(cell):
        import re
        m = re.match(r"([A-Z]+)(\d+)", cell)
        if not m:
            raise RuntimeError("Unexpected table ref format")
        col_letters, row = m.groups()
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - ord('A') + 1)
        return int(row), col

    if table_obj is not None:
        # update existing table
        ref = table_obj.ref
        start_cell, end_cell = ref.split(":")
        start_row, start_col = cell_to_rowcol(start_cell)
        end_row, end_col = cell_to_rowcol(end_cell)

        # clear old rows under header
        for r in range(start_row + 1, end_row + 1):
            for c in range(start_col, end_col + 1):
                table_sheet.cell(row=r, column=c).value = None

        # write new header and rows
        header = list(df.columns)
        for idx, col_name in enumerate(header):
            table_sheet.cell(row=start_row, column=start_col + idx, value=col_name)

        for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=start_row + 1):
            for c_idx, v in enumerate(row, start=start_col):
                table_sheet.cell(row=r_idx, column=c_idx, value=v)

        # update ref
        new_end_row = start_row + len(df)
        new_end_col = start_col + len(header) - 1
        table_obj.ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(new_end_col)}{new_end_row}"
        target_ws = table_sheet
        header_row_idx = start_row
        first_data_row = start_row + 1
    else:
        # create a new sheet named table_name (avoid collisions)
        sheet_name = table_name
        base_name = sheet_name
        i = 1
        while sheet_name in [ws.title for ws in wb.worksheets]:
            sheet_name = f"{base_name}_{i}"
            i += 1
        ws_new = wb.create_sheet(sheet_name)

        header = list(df.columns)
        ws_new.append(header)
        for row in df.itertuples(index=False, name=None):
            ws_new.append(row)

        max_row = ws_new.max_row
        max_col = ws_new.max_column
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws_new.add_table(table)

        target_ws = ws_new
        header_row_idx = 1
        first_data_row = 2

    # --- apply conditional formatting to DOC column in target_ws ---
    # find DOC column index by header row
    header_cells = list(target_ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=False))[0]
    doc_col_idx = None
    for idx, cell in enumerate(header_cells, start=1):
        if cell.value and str(cell.value).strip().lower() == "doc":
            doc_col_idx = idx
            break
    if doc_col_idx is None:
        for idx, cell in enumerate(header_cells, start=1):
            if cell.value and "doc" in str(cell.value).strip().lower():
                doc_col_idx = idx
                break

    col_map = {
        "dark_red": "8B0000",
        "dark_orange": "FF8C00",
        "dark_green": "006400",
        "golden": "B8860B",
        "sky": "0F52BA",
        "saddle": "8B4513",
        "white": "FFFFFF",
    }

    if doc_col_idx is not None:
        try:
            last_row = target_ws.max_row
            col_letter = get_column_letter(doc_col_idx)
            rng = f"{col_letter}{first_data_row}:{col_letter}{last_row}"

            # create rules for buckets using CellIsRule
            rules = [
                CellIsRule(operator='between', formula=['0','6.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["dark_red"], end_color=col_map["dark_red"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['7','14.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["dark_orange"], end_color=col_map["dark_orange"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['15','29.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["dark_green"], end_color=col_map["dark_green"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['30','44.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["golden"], end_color=col_map["golden"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['45','59.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["sky"], end_color=col_map["sky"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
                CellIsRule(operator='between', formula=['60','89.9999'],
                           stopIfTrue=True,
                           fill=PatternFill(start_color=col_map["saddle"], end_color=col_map["saddle"], fill_type="solid"),
                           font=Font(color="FFFFFF")),
            ]

            for rrule in rules:
                target_ws.conditional_formatting.add(rng, rrule)
        except Exception:
            pass

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def create_fallback_workbook(df: pd.DataFrame, sort_desc: bool, sheet_name: str, parent_col: str = None, selected_brands = None):
    """
    Create fallback workbook bytes (openpyxl) with:
      - data sheet (DOC styling)
      - PivotSummary (aggregated sums by Brand + parent)
      - ChartData and Chart
      - HowToPivot sheet
    """
    colors = {
        "dark_red": "8B0000",
        "dark_orange": "FF8C00",
        "dark_green": "006400",
        "golden": "B8860B",
        "sky": "0F52BA",
        "saddle": "8B4513",
        "white": "FFFFFF",
    }

    working = df.copy()
    if selected_brands:
        working = working[working["Brand"].isin(selected_brands)].copy()
    working = working.sort_values(by="DOC", ascending=(not sort_desc)).reset_index(drop=True)

    if parent_col and parent_col in working.columns:
        agg = working.groupby(["Brand", parent_col], dropna=False)[["DOC", "DRR"]].sum().reset_index()
        agg["Brand_Parent"] = agg["Brand"].astype(str) + " | " + agg[parent_col].astype(str)
    elif "Brand" in working.columns:
        agg = working.groupby(["Brand"], dropna=False)[["DOC", "DRR"]].sum().reset_index()
        agg["Brand_Parent"] = agg["Brand"].astype(str)
    else:
        agg = pd.DataFrame(columns=["Brand_Parent", "DOC", "DRR"])

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # write data
    for r in dataframe_to_rows(working, index=False, header=True):
        ws.append(r)

    # apply DOC fills
    try:
        if "DOC" in working.columns:
            doc_idx = list(working.columns).index("DOC") + 1
            from openpyxl.styles import PatternFill, Font
            def fill_for_val(v):
                try:
                    v = float(v)
                except:
                    return PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid"), Font(color="000000")
                if 0 <= v < 7:
                    return PatternFill(start_color=colors["dark_red"], end_color=colors["dark_red"], fill_type="solid"), Font(color="FFFFFF")
                if 7 <= v < 15:
                    return PatternFill(start_color=colors["dark_orange"], end_color=colors["dark_orange"], fill_type="solid"), Font(color="FFFFFF")
                if 15 <= v < 30:
                    return PatternFill(start_color=colors["dark_green"], end_color=colors["dark_green"], fill_type="solid"), Font(color="FFFFFF")
                if 30 <= v < 45:
                    return PatternFill(start_color=colors["golden"], end_color=colors["golden"], fill_type="solid"), Font(color="FFFFFF")
                if 45 <= v < 60:
                    return PatternFill(start_color=colors["sky"], end_color=colors["sky"], fill_type="solid"), Font(color="FFFFFF")
                if 60 <= v < 90:
                    return PatternFill(start_color=colors["saddle"], end_color=colors["saddle"], fill_type="solid"), Font(color="FFFFFF")
                return PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid"), Font(color="000000")

            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=doc_idx)
                fill, font = fill_for_val(cell.value)
                cell.fill = fill
                cell.font = font
    except Exception:
        pass

    # add an Excel Table (DataTable)
    try:
        from openpyxl.worksheet.table import Table, TableStyleInfo
        max_row = ws.max_row
        max_col = ws.max_column
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName="DataTable", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(table)
    except Exception:
        pass

    # PivotSummary
    ws_pivot = wb.create_sheet("PivotSummary")
    for r in dataframe_to_rows(agg, index=False, header=True):
        ws_pivot.append(r)

    # apply DOC fills on PivotSummary
    try:
        headers = [c.value for c in ws_pivot[1]]
        if "DOC" in headers:
            doc_idx_p = headers.index("DOC") + 1
            from openpyxl.styles import PatternFill, Font
            for row in range(2, ws_pivot.max_row + 1):
                cell = ws_pivot.cell(row=row, column=doc_idx_p)
                try:
                    val = float(cell.value)
                except:
                    val = None
                if val is None:
                    cell.fill = PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid")
                    cell.font = Font(color="000000")
                elif 0 <= val < 7:
                    cell.fill = PatternFill(start_color=colors["dark_red"], end_color=colors["dark_red"], fill_type="solid")
                    cell.font = Font(color="FFFFFF")
                elif 7 <= val < 15:
                    cell.fill = PatternFill(start_color=colors["dark_orange"], end_color=colors["dark_orange"], fill_type="solid")
                    cell.font = Font(color="FFFFFF")
                elif 15 <= val < 30:
                    cell.fill = PatternFill(start_color=colors["dark_green"], end_color=colors["dark_green"], fill_type="solid")
                    cell.font = Font(color="FFFFFF")
                elif 30 <= val < 45:
                    cell.fill = PatternFill(start_color=colors["golden"], end_color=colors["golden"], fill_type="solid")
                    cell.font = Font(color="FFFFFF")
                elif 45 <= val < 60:
                    cell.fill = PatternFill(start_color=colors["sky"], end_color=colors["sky"], fill_type="solid")
                    cell.font = Font(color="FFFFFF")
                elif 60 <= val < 90:
                    cell.fill = PatternFill(start_color=colors["saddle"], end_color=colors["saddle"], fill_type="solid")
                    cell.font = Font(color="FFFFFF")
                else:
                    cell.fill = PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid")
                    cell.font = Font(color="000000")
    except Exception:
        pass

    # ChartData + Chart
    ws_chartdata = wb.create_sheet("ChartData")
    if not agg.empty:
        for r in dataframe_to_rows(agg[["Brand_Parent", "DOC", "DRR"]], index=False, header=True):
            ws_chartdata.append(r)
        try:
            from openpyxl.chart import BarChart, Reference
            if ws_chartdata.max_row > 1:
                chart = BarChart()
                cats = Reference(ws_chartdata, min_col=1, min_row=2, max_row=ws_chartdata.max_row)
                vals1 = Reference(ws_chartdata, min_col=2, min_row=2, max_row=ws_chartdata.max_row)
                vals2 = Reference(ws_chartdata, min_col=3, min_row=2, max_row=ws_chartdata.max_row)
                chart.add_data(vals1, titles_from_data=False)
                chart.add_data(vals2, titles_from_data=False)
                chart.set_categories(cats)
                chart.title = "Sum DOC and DRR by Brand + Parent"
                ws_chart = wb.create_sheet("Chart")
                ws_chart.add_chart(chart, "A1")
        except Exception:
            pass

    # HowToPivot
    try:
        how = wb.create_sheet("HowToPivot")
        how.append(["How to create interactive PivotTable + Slicer in Excel (no VBA)"])
        how.append([])
        how.append(["1) In Excel: Insert → PivotTable"])
        how.append([f"   - Select the table named 'DataTable' on the sheet: {sheet_name if 'sheet_name' in locals() else 'Data'}"])
        how.append(["   - Place the PivotTable on a new worksheet."])
        how.append([])
        how.append(["2) In PivotField list: drag 'Brand' and '(Parent) ASIN' (or your parent column) into Rows (Brand first)."])
        how.append(["   - Drag 'DOC' and 'DRR' into Values (set aggregation = Sum)."])
        how.append([])
        how.append(["3) To add a Slicer: Insert → Slicer → choose 'Brand' and '(Parent) ASIN'."])
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# -------------------------
# UI: file upload + processing
# -------------------------
st.title("📊 Inventory Analysis Dashboard")
st.markdown("Upload your files and analyze inventory with Days of Coverage (DOC) calculations")

with st.sidebar:
    st.header("⚙️ Configuration")
    no_of_days = st.number_input(
        "Enter the number of days",
        min_value=1,
        value=31,
        step=1,
        help="This represents the time period for your sales data. DRR = Total Order Items ÷ Days.",
    )
    st.info(
        """
- DRR (Daily Run Rate) = Total Order Items ÷ Days
- DOC (Days of Coverage) = afn-fulfillable-quantity ÷ DRR
"""
    )

st.header("📁 Upload Files")
col1, col2, col3 = st.columns(3)
with col1:
    business_file = st.file_uploader("Upload Business Report CSV", type=["csv"], key="business")
with col2:
    pm_file = st.file_uploader("Upload PM Excel/CSV", type=["xlsx", "csv"], key="pm")
with col3:
    inventory_file = st.file_uploader("Upload Manage Inventory CSV", type=["csv"], key="inventory")

st.markdown("---")

if st.button("🚀 Process Data"):
    if business_file is None or pm_file is None or inventory_file is None:
        st.error("⚠️ Please upload all three files before processing!")
    elif no_of_days <= 0:
        st.error("⚠️ Number of days must be greater than 0!")
    else:
        with st.spinner("Processing data..."):
            try:
                original = pd.read_csv(business_file)

                if "SKU" in original.columns:
                    original["SKU"] = original["SKU"].astype(str)

                if pm_file.name.endswith(".xlsx"):
                    pm = pd.read_excel(pm_file)
                else:
                    pm = pd.read_csv(pm_file)

                inventory = pd.read_csv(inventory_file)
                inventory.columns = inventory.columns.str.strip()
                inventory.iloc[:, 0] = inventory.iloc[:, 0].astype(str)

                # process pm columns C:G
                pm = pm.iloc[:, 2:7]
                pm.columns = ["Amazon Sku Name", "D", "Brand Manager", "F", "Brand"]
                pm["Amazon Sku Name"] = pm["Amazon Sku Name"].astype(str)

                original = original.merge(pm[["Amazon Sku Name", "Brand Manager"]], how="left", left_on="SKU", right_on="Amazon Sku Name")
                if "Title" in original.columns and "Brand Manager" in original.columns:
                    insert_pos = original.columns.get_loc("Title")
                    col = original.pop("Brand Manager")
                    original.insert(insert_pos, "Brand Manager", col)

                original = original.merge(pm[["Amazon Sku Name", "Brand"]], how="left", left_on="SKU", right_on="Amazon Sku Name")
                if "Title" in original.columns and "Brand" in original.columns:
                    insert_pos = original.columns.get_loc("Title")
                    col = original.pop("Brand")
                    original.insert(insert_pos, "Brand", col)

                # inventory mapping
                if inventory.shape[1] > 10:
                    return_col = inventory.columns[10]
                    mi_map = inventory.set_index(inventory.columns[0])[return_col]
                    original["afn-fulfillable-quantity"] = original["SKU"].map(mi_map)
                else:
                    original["afn-fulfillable-quantity"] = 0

                if inventory.shape[1] > 12:
                    return_col_13 = inventory.columns[12]
                    mi_res_map = inventory.set_index(inventory.columns[0])[return_col_13]
                    original["afn-reserved-quantity"] = original["SKU"].map(mi_res_map)
                else:
                    original["afn-reserved-quantity"] = 0

                # clean & compute DRR/DOC
                if "Total Order Items" in original.columns:
                    original["Total Order Items"] = (
                        original["Total Order Items"]
                        .astype(str)
                        .str.replace("\u00A0", "", regex=False)
                        .str.replace(",", "", regex=False)
                        .str.replace(r"[^\d\.\-]", "", regex=True)
                    )
                    original["Total Order Items"] = pd.to_numeric(original["Total Order Items"], errors="coerce")
                else:
                    original["Total Order Items"] = 0

                original["DRR"] = (original["Total Order Items"] / no_of_days).round(2)
                original["afn-fulfillable-quantity"] = pd.to_numeric(original["afn-fulfillable-quantity"], errors="coerce")
                original["DOC"] = (original["afn-fulfillable-quantity"] / original["DRR"]).round(2)
                original["DOC"] = original["DOC"].replace([float("inf"), float("-inf")], 0)

                st.success("✅ Data processed successfully!")

                # display metrics
                st.header("📈 Processed Results")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Total Products", len(original))
                with c2:
                    st.metric("Critical Stock (< 7 days)", int((original["DOC"] < 7).sum()))
                with c3:
                    st.metric("Average DOC", f"{original['DOC'].mean():.2f} days")
                with c4:
                    st.metric("Total Orders", f"{original['Total Order Items'].sum():,.0f}")

                st.markdown("---")

                display_cols = [
                    "(Child) ASIN",
                    "Brand",
                    "Brand Manager",
                    "SKU",
                    "Title",
                    "Units Ordered",
                    "afn-fulfillable-quantity",
                    "afn-reserved-quantity",
                    "DRR",
                    "DOC",
                ]
                display_cols = [col for col in display_cols if col in original.columns]
                display_df = original[display_cols].copy()
                styled_df = display_df.style.map(color_doc, subset=["DOC"])

                st.dataframe(styled_df, width="stretch", height=600)

                st.markdown("---")

                # Excel export UI (template-first)
                colA, colB = st.columns(2)
                with colA:
                    # plain CSV download
                    csv_buf = io.StringIO()
                    original.to_csv(csv_buf, index=False)
                    st.download_button(
                        "📥 Download CSV",
                        data=csv_buf.getvalue(),
                        file_name=f"processed_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                    )

                    # XLSX download that looks like CSV but with DOC conditional formatting
                    xlsx_bytes = create_excel_with_doc_format(original)
                    st.download_button(
                        "📥 Download CSV as Excel (DOC formatted)",
                        data=xlsx_bytes,
                        file_name=f"processed_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}_DOCformatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                with colB:
                    st.markdown("### Excel export (Template-first pivot)")
                    over = st.button("📥 Overstock (DOC ↓) - Export Excel", key="over_export")
                    oos = st.button("📥 OOS (DOC ↑) - Export Excel", key="oos_export")

                    if over or oos:
                        sort_desc = True if over else False
                        # detect parent col
                        parent_col = None
                        for c in original.columns:
                            cle = c.lower().replace(" ", "").replace("_", "").replace("(", "").replace(")", "")
                            if "parent" in cle:
                                parent_col = c
                                break

                        brands = sorted(original["Brand"].dropna().astype(str).unique().tolist()) if "Brand" in original.columns else []
                        selected = st.multiselect("Filter brands for export (leave empty = all)", options=brands, default=brands)

                        # 🔴 yahan sirf filtered df use ho raha hai
                        if over:
                            df_export = filter_overstock(original)     # DOC >= 90
                        else:
                            df_export = filter_oos(original)           # afn-fulfillable-quantity == 0

                        if selected:
                            df_export = df_export[df_export["Brand"].isin(selected)].copy()

                        df_export = df_export.sort_values(by="DOC", ascending=(not sort_desc)).reset_index(drop=True)

                        # template detection: prefer macro-enabled .xlsm if present
                        base_dir = os.path.dirname(__file__)
                        tmpl_xlsm = os.path.join(base_dir, "pivot_template.xlsm")
                        tmpl_xlsx = os.path.join(base_dir, "pivot_template.xlsx")
                        template_path = tmpl_xlsm if os.path.exists(tmpl_xlsm) else (tmpl_xlsx if os.path.exists(tmpl_xlsx) else None)

                        final_bytes = None
                        template_used = False
                        template_error = None

                        if template_path:
                            try:
                                buf = fill_template_and_get_bytes(template_path, df_export, table_name="DataTable")
                                final_bytes = buf.getvalue()
                                template_used = True
                                st.success("✅ Used pivot template — Pivot/Slicer included when opened in Excel.")
                            except Exception as te:
                                template_error = traceback.format_exc()
                                st.warning("⚠️ Template found but failed to be filled programmatically. Falling back to generated workbook.")
                                st.code(template_error)

                        if final_bytes is None:
                            fallback_buf = create_fallback_workbook(
                                df_export,
                                sort_desc=sort_desc,
                                sheet_name="Overstock" if sort_desc else "OOS",
                                parent_col=parent_col,
                                selected_brands=selected,
                            )
                            final_bytes = fallback_buf.getvalue()
                            st.info("ℹ️ Delivered fallback workbook (DataTable + PivotSummary + ChartData + HowToPivot).")

                        # pick correct extension & mime based on template vs fallback
                        if template_path and template_used and template_path.lower().endswith(".xlsm"):
                            dl_ext = ".xlsm"
                            dl_mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
                        else:
                            dl_ext = ".xlsx"
                            dl_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                        file_basename = f"{'overstock' if sort_desc else 'oos'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        file_name = file_basename + dl_ext

                        st.download_button(
                            label="Download Excel workbook",
                            data=final_bytes,
                            file_name=file_name,
                            mime=dl_mime,
                        )

                # persist processed data
                st.session_state["processed_data"] = original

            except Exception as e:
                st.error("❌ Processing failed.")
                st.exception(e)

# previously processed view
elif "processed_data" in st.session_state:
    orig = st.session_state["processed_data"]
    st.header("📈 Previously Processed Results")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Total Products", len(orig))
    with c2:
        st.metric("Critical Stock (< 7 days)", int((orig["DOC"] < 7).sum()))
    with c3:
        st.metric("Average DOC", f"{orig['DOC'].mean():.2f} days")
    with c4:
        st.metric("Total Orders", f"{orig['Total Order Items'].sum():,.0f}")

    st.markdown("---")
    display_cols = [
        "(Child) ASIN",
        "Brand",
        "Brand Manager",
        "SKU",
        "Title",
        "Units Ordered",
        "afn-fulfillable-quantity",
        "afn-reserved-quantity",
        "DRR",
        "DOC",
    ]
    display_cols = [col for col in display_cols if col in orig.columns]
    display_df = orig[display_cols].copy()
    styled_df = display_df.style.map(color_doc, subset=["DOC"])
    st.dataframe(styled_df, width="stretch", height=600)

    st.markdown("---")
    cA, cB = st.columns(2)
    with cA:
        # plain CSV download
        csv_buf = io.StringIO()
        orig.to_csv(csv_buf, index=False)
        st.download_button(
            "📥 Download CSV",
            data=csv_buf.getvalue(),
            file_name=f"processed_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )

        # XLSX download that looks like CSV but with DOC conditional formatting
        xlsx_bytes_prev = create_excel_with_doc_format(orig)
        st.download_button(
            "📥 Download CSV as Excel (DOC formatted)",
            data=xlsx_bytes_prev,
            file_name=f"processed_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}_DOCformatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with cB:
        st.markdown("### Excel export (Template-first pivot)")
        over2 = st.button("📥 Overstock (DOC ↓) - Export Excel (previous)", key="over_prev")
        oos2 = st.button("📥 OOS (DOC ↑) - Export Excel (previous)", key="oos_prev")

        if over2 or oos2:
            sort_desc = True if over2 else False
            parent_col = None
            for c in orig.columns:
                cle = c.lower().replace(" ", "").replace("_", "").replace("(", "").replace(")", "")
                if "parent" in cle:
                    parent_col = c
                    break
            brands = sorted(orig["Brand"].dropna().astype(str).unique().tolist()) if "Brand" in orig.columns else []
            selected = st.multiselect("Filter brands for export (leave empty = all)", options=brands, default=brands)

            # yahan bhi wahi filters apply:
            if over2:
                df_export = filter_overstock(orig)
            else:
                df_export = filter_oos(orig)

            if selected:
                df_export = df_export[df_export["Brand"].isin(selected)].copy()

            df_export = df_export.sort_values(by="DOC", ascending=(not sort_desc)).reset_index(drop=True)

            base_dir = os.path.dirname(__file__)
            tmpl_xlsm = os.path.join(base_dir, "pivot_template.xlsm")
            tmpl_xlsx = os.path.join(base_dir, "pivot_template.xlsx")
            template_path = tmpl_xlsm if os.path.exists(tmpl_xlsm) else (tmpl_xlsx if os.path.exists(tmpl_xlsx) else None)

            final_bytes = None
            template_used = False

            if template_path:
                try:
                    buf = fill_template_and_get_bytes(template_path, df_export, table_name="DataTable")
                    final_bytes = buf.getvalue()
                    template_used = True
                    st.success("✅ Used pivot template — Pivot/Slicer included when opened in Excel.")
                except Exception as te:
                    st.warning("⚠️ Template found but failed to be filled programmatically. Falling back to generated workbook.")
                    st.code(traceback.format_exc())

            if final_bytes is None:
                fallback_buf = create_fallback_workbook(
                    df_export,
                    sort_desc=sort_desc,
                    sheet_name="Overstock" if sort_desc else "OOS",
                    parent_col=parent_col,
                    selected_brands=selected,
                )
                final_bytes = fallback_buf.getvalue()
                st.info("ℹ️ Delivered fallback workbook (DataTable + PivotSummary + ChartData + HowToPivot).")

            # pick correct extension & mime based on template vs fallback
            if template_path and template_used and template_path.lower().endswith(".xlsm"):
                dl_ext = ".xlsm"
                dl_mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
            else:
                dl_ext = ".xlsx"
                dl_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            file_basename = f"{'overstock' if sort_desc else 'oos'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            file_name = file_basename + dl_ext

            st.download_button(
                label="Download Excel workbook",
                data=final_bytes,
                file_name=file_name,
                mime=dl_mime,
            )

# footer
st.markdown("---")
st.markdown("<div style='text-align: center; color: #666; padding: 10px;'>Inventory Analysis Dashboard | Built with Streamlit</div>", unsafe_allow_html=True)
