import streamlit as st
import pandas as pd
import io
import os
import sys
import tempfile
import traceback
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------------
# COM helper: create_pivot_with_com (FULL - creates PivotTable, PivotChart, Slicer)
# ---------------------
def create_pivot_with_com(xlsx_path: str,
                          data_sheet_name: str = "Overstock",
                          table_name: str = "DataTable",
                          pivot_sheet_name: str = "PivotTable",
                          parent_field_hint: str = "(Parent) ASIN"):
    """
    COM-based Pivot creation wrapper that ensures COM is initialized on the calling thread.
    Returns (True, None) on success, or (False, traceback_str) on failure.
    Requires: Windows + Excel + pywin32.
    """
    import traceback, os
    try:
        import pythoncom
        import win32com.client as win32
        from win32com.client import constants
    except Exception as e:
        tb = traceback.format_exc()
        return False, f"pywin32/pythoncom import failed: {e}\n\n{tb}"

    excel = None
    wb = None
    initialized = False
    try:
        # Initialize COM on this thread (idempotent if already called)
        # Use Apartment-threaded model which Excel expects.
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            initialized = True
        except Exception:
            # fallback to simple CoInitialize (older pywin32 builds)
            try:
                pythoncom.CoInitialize()
                initialized = True
            except Exception:
                # if both fail, continue and attempt EnsureDispatch (we'll capture failure)
                initialized = False

        # Now do the COM dispatch
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

        # find data sheet
        ws_data = None
        for ws in wb.Worksheets:
            if ws.Name == data_sheet_name:
                ws_data = ws
                break
        if ws_data is None:
            ws_data = wb.Worksheets(1)

        used = ws_data.UsedRange
        if used is None:
            raise RuntimeError("Could not determine UsedRange on data sheet.")
        addr = used.Address
        source = f"'{ws_data.Name}'!{addr}"

        pivot_cache = wb.PivotCaches().Create(SourceType=1, SourceData=source)

        try:
            pivot_ws = wb.Worksheets(pivot_sheet_name)
        except Exception:
            pivot_ws = wb.Worksheets.Add()
            pivot_ws.Name = pivot_sheet_name

        pivot_table_name = "PivotFromCOM"
        pivot_table = pivot_cache.CreatePivotTable(TableDestination=f"'{pivot_ws.Name}'!R3C1", TableName=pivot_table_name)

        xlRowField = constants.xlRowField
        xlSum = constants.xlSum

        # read headers
        headers = []
        first_row = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(1, ws_data.UsedRange.Columns.Count))
        for i in range(1, first_row.Columns.Count + 1):
            val = first_row.Cells(1, i).Value
            headers.append(str(val).strip() if val is not None else "")

        brand_field = None
        parent_field = None
        for h in headers:
            if h and h.strip().lower() == "brand":
                brand_field = h
            if parent_field_hint and parent_field_hint.lower() in h.lower():
                parent_field = h
        if not brand_field:
            for h in headers:
                if "brand" in h.lower():
                    brand_field = h
                    break
        if not parent_field:
            for h in headers:
                if "parent" in h.lower():
                    parent_field = h
                    break

        if brand_field:
            pivot_table.PivotFields(brand_field).Orientation = xlRowField
        if parent_field:
            pivot_table.PivotFields(parent_field).Orientation = xlRowField

        for field in ("DOC", "DRR"):
            try:
                pivot_table.AddDataField(pivot_table.PivotFields(field), f"Sum of {field}", xlSum)
            except Exception:
                pass

        # add pivot chart (best effort)
        try:
            chart_obj = pivot_ws.Shapes.AddChart2(201, 51, 400, 10, 600, 300)
            chart = chart_obj.Chart
            try:
                chart.SetSourceData(pivot_table.TableRange2)
            except Exception:
                pass
            try:
                chart.ChartTitle.Text = "Sum DOC and DRR by Brand + Parent"
            except Exception:
                pass
        except Exception:
            pass

        # add slicer (best effort)
        try:
            if brand_field:
                slicer_cache = wb.SlicerCaches.Add(pivot_table.PivotFields(brand_field))
                slicer_cache.Slicers.Add(pivot_ws, Name=f"Slicer_{brand_field}", Caption=brand_field, Left=10, Top=10, Width=140, Height=200)
        except Exception:
            pass

        wb.Save()
        wb.Close(SaveChanges=True)
        try:
            excel.Quit()
        except Exception:
            pass

        return True, None

    except Exception as e:
        tb = traceback.format_exc()
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass
        return False, tb

    finally:
        # Uninitialize COM if we initialized it here
        try:
            if initialized:
                pythoncom.CoUninitialize()
        except Exception:
            pass


# ---------------------
# Helper: create workbook (pivot attempt) + apply DOC conditional formatting for both engines
# ---------------------
def create_workbook_with_brand_parent_pivot_and_doc_format(df_full, sort_desc: bool, sheet_name: str, selected_brands, parent_col):
    """
    Create workbook bytes:
    - Try programmatic pivot creation using xlsxwriter (if available).
    - If not, create openpyxl workbook with DataTable, PivotSummary, ChartData, Chart, HowToPivot.
    - Always apply DOC conditional formatting on exported sheets (data & pivot summary).
    Returns (BytesIO, pivot_created_flag, pivot_error_or_none)
    """
    from openpyxl.utils import get_column_letter

    # filter & sort
    working = df_full[df_full["Brand"].isin(selected_brands)].copy() if selected_brands else df_full.copy()
    working = working.sort_values(by="DOC", ascending=(not sort_desc)).reset_index(drop=True)

    # Build aggregated ChartData by Brand+Parent
    if parent_col and parent_col in working.columns:
        agg = working.groupby(["Brand", parent_col], dropna=False)[["DOC", "DRR"]].sum().reset_index()
        agg["Brand_Parent"] = agg["Brand"].astype(str) + " | " + agg[parent_col].astype(str)
    elif "Brand" in working.columns:
        agg = working.groupby(["Brand"], dropna=False)[["DOC", "DRR"]].sum().reset_index()
        agg["Brand_Parent"] = agg["Brand"].astype(str)
    else:
        agg = pd.DataFrame(columns=["Brand_Parent", "DOC", "DRR"])

    # color hex codes for Excel rules (without '#')
    colors = {
        "dark_red": "8B0000",
        "dark_orange": "FF8C00",
        "dark_green": "006400",
        "golden": "B8860B",
        "sky": "0F52BA",
        "saddle": "8B4513",
        "black": "222222",
        "white": "FFFFFF",
    }

    # --- Try xlsxwriter route (programmatic pivot + conditional formatting) ---
    try:
        import xlsxwriter
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            working.to_excel(writer, sheet_name=sheet_name, index=False)
            agg.to_excel(writer, sheet_name="ChartData", index=False)

            workbook = writer.book
            ws_data = writer.sheets[sheet_name]
            ws_chartdata = writer.sheets["ChartData"]

            # Add Excel table for data (improves pivot source)
            try:
                from xlsxwriter.utility import xl_range
                nrows = working.shape[0] + 1
                ncols = working.shape[1]
                table_range = xl_range(0, 0, nrows - 1, ncols - 1)
                ws_data.add_table(table_range, {'name': 'DataTable', 'columns': [{'header': h} for h in working.columns.tolist()]})
            except Exception:
                nrows = working.shape[0] + 1
                ncols = working.shape[1]

            pivot_created = False
            pivot_error = None

            # Attempt pivot creation using xlsxwriter pivot API if present
            try:
                if hasattr(workbook, "add_pivot_table"):
                    rows = [{"field": "Brand"}]
                    parent_field = "(Parent) ASIN" if "(Parent) ASIN" in working.columns else parent_col
                    if parent_field and parent_field in working.columns:
                        rows.append({"field": parent_field})
                    pivot_options = {
                        "data": {"sheet": sheet_name, "ref": None, "table": "DataTable"},
                        "rows": rows,
                        "columns": [],
                        "filters": [],
                        "values": [{"field": "DOC", "function": "sum"}, {"field": "DRR", "function": "sum"}],
                        "location": {"sheet": "PivotTable", "ref": "A3"},
                    }
                    workbook.add_worksheet("PivotTable")
                    workbook.add_pivot_table(pivot_options)
                    pivot_created = True
                else:
                    pivot_error = AttributeError("xlsxwriter pivot API not present.")
                    pivot_created = False
            except Exception as e_p:
                pivot_error = e_p
                pivot_created = False

            # Add a chart sheet from ChartData as visual
            try:
                chart = workbook.add_chart({'type': 'column'})
                rows_agg = agg.shape[0]
                if rows_agg > 0:
                    chart.add_series({'name': 'Sum of DOC', 'categories': f"=ChartData!$A$2:$A${rows_agg+1}", 'values': f"=ChartData!$B$2:$B${rows_agg+1}"})
                    chart.add_series({'name': 'Sum of DRR', 'categories': f"=ChartData!$A$2:$A${rows_agg+1}", 'values': f"=ChartData!$C$2:$C${rows_agg+1}"})
                    chart.set_title({'name': 'Sum DOC and DRR by Brand + Parent'})
                    chart.set_x_axis({'name': 'Brand | ParentASIN'})
                    chart_sheet = workbook.add_worksheet("Chart")
                    chart_sheet.insert_chart('A1', chart)
            except Exception:
                pass

            # Apply conditional formatting to DOC column in data sheet (xlsxwriter)
            try:
                if "DOC" in working.columns:
                    doc_col_idx = working.columns.get_loc("DOC")
                    from xlsxwriter.utility import xl_col_to_name
                    doc_col_letter = xl_col_to_name(doc_col_idx)
                    data_range = f"{doc_col_letter}2:{doc_col_letter}{nrows}"

                    fmt_dark_red = workbook.add_format({'bg_color': f"#{colors['dark_red']}", 'font_color': "#FFFFFF"})
                    fmt_orange = workbook.add_format({'bg_color': f"#{colors['dark_orange']}", 'font_color': "#FFFFFF"})
                    fmt_green = workbook.add_format({'bg_color': f"#{colors['dark_green']}", 'font_color': "#FFFFFF"})
                    fmt_gold = workbook.add_format({'bg_color': f"#{colors['golden']}", 'font_color': "#FFFFFF"})
                    fmt_sky = workbook.add_format({'bg_color': f"#{colors['sky']}", 'font_color': "#FFFFFF"})
                    fmt_saddle = workbook.add_format({'bg_color': f"#{colors['saddle']}", 'font_color': "#FFFFFF"})
                    fmt_default = workbook.add_format({'bg_color': f"#{colors['white']}", 'font_color': "#000000"})

                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': '<', 'value': 7, 'format': fmt_dark_red})
                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': 'between', 'minimum': 7, 'maximum': 14.999, 'format': fmt_orange})
                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': 'between', 'minimum': 15, 'maximum': 29.999, 'format': fmt_green})
                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': 'between', 'minimum': 30, 'maximum': 44.999, 'format': fmt_gold})
                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': 'between', 'minimum': 45, 'maximum': 59.999, 'format': fmt_sky})
                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': 'between', 'minimum': 60, 'maximum': 89.999, 'format': fmt_saddle})
                    ws_data.conditional_format(data_range, {'type': 'cell', 'criteria': '>=', 'value': 90, 'format': fmt_default})
            except Exception:
                pass

        out.seek(0)
        return out, pivot_created, pivot_error

    except Exception as xlsx_err:
        pivot_err = xlsx_err

    # --- Fallback: openpyxl (deterministic) ---
    try:
        wb = Workbook()
        ws_sorted = wb.active
        ws_sorted.title = sheet_name

        # write sorted data
        for r in dataframe_to_rows(working, index=False, header=True):
            ws_sorted.append(r)

        # apply DOC conditional formatting (by direct fills) on sorted sheet
        try:
            if "DOC" in working.columns:
                doc_idx = list(working.columns).index("DOC") + 1
                from openpyxl.styles import PatternFill, Font
                def _fill_for_val(v):
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

                for row in range(2, ws_sorted.max_row + 1):
                    cell = ws_sorted.cell(row=row, column=doc_idx)
                    fill, font = _fill_for_val(cell.value)
                    cell.fill = fill
                    cell.font = font
        except Exception:
            pass

        # try to add a table named DataTable
        try:
            from openpyxl.worksheet.table import Table, TableStyleInfo
            max_row = ws_sorted.max_row
            max_col = ws_sorted.max_column
            ref = f"A1:{get_column_letter(max_col)}{max_row}"
            table = Table(displayName="DataTable", ref=ref)
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            ws_sorted.add_table(table)
        except Exception:
            pass

        # PivotSummary aggregated sheet
        ws_pivot = wb.create_sheet("PivotSummary")
        for r in dataframe_to_rows(agg, index=False, header=True):
            ws_pivot.append(r)

        # Apply DOC formatting to PivotSummary if DOC present
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

        # ChartData sheet
        ws_chartdata = wb.create_sheet("ChartData")
        for r in dataframe_to_rows(agg[["Brand_Parent", "DOC", "DRR"]], index=False, header=True):
            ws_chartdata.append(r)

        # Create a bar chart on Chart sheet
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

        # HowToPivot sheet
        try:
            how = wb.create_sheet("HowToPivot")
            how.append(["How to create interactive PivotTable + Slicer in Excel (no VBA)"])
            how.append([])
            how.append(["1) In Excel: Insert → PivotTable"])
            how.append([f"   - Select the table named 'DataTable' on the sheet: {sheet_name}"])
            how.append(["   - Place the PivotTable on a new worksheet."])
            how.append([])
            how.append(["2) In PivotField list: drag 'Brand' and '(Parent) ASIN' (or your parent column) into Rows (Brand first)."])
            how.append(["   - Drag 'DOC' and 'DRR' into Values (set aggregation = Sum)."])
            how.append([])
            how.append(["3) To add a Slicer: Insert → Slicer → choose 'Brand'."])
            how.append(["   - The slicer will filter both the Pivot and any PivotChart based on that Pivot."])
        except Exception:
            pass

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf, False, pivot_err

    except Exception as fb_err:
        raise RuntimeError(f"Both pivot attempt and fallback failed: {pivot_err}; fallback_err={fb_err}")

# ---------------------
# UI & existing app code (kept your original flow; only wiring changed to call COM optionally)
# ---------------------

# Page configuration
st.set_page_config(
    page_title="OOS Amazon Analysis",
    page_icon="📊",
    layout="wide",
)

# Custom CSS for colored cells
st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

# Helper function to color DOC values (returns CSS style string)
def color_doc(val):
    """Apply dark color based on DOC value"""
    try:
        doc = float(val)
        if 0 <= doc < 7:
            return "background-color: #8B0000; color: white"  # Dark Red
        elif 7 <= doc < 15:
            return "background-color: #FF8C00; color: white"  # Dark Orange
        elif 15 <= doc < 30:
            return "background-color: #006400; color: white"  # Dark Green
        elif 30 <= doc < 45:
            return "background-color: #B8860B; color: white"  # Dark Goldenrod (yellowish)
        elif 45 <= doc < 60:
            return "background-color: #0F52BA; color: white"  # Dark Sky Blue / Sapphire
        elif 60 <= doc < 90:
            return "background-color: #8B4513; color: white"  # Saddle Brown
        else:
            return "background-color: #222222; color: white"
    except:
        return ""

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
        help="This represents the time period for your sales data. The DRR (Daily Run Rate) will be calculated by dividing Total Order Items by this number.",
    )

    st.info(
        """
    **What is Number of Days?**

    This is the time period covered by your sales data (e.g., 30 or 31 for monthly data).

    - **DRR** (Daily Run Rate) = Total Order Items ÷ Days
    - **DOC** (Days of Coverage) = Fulfillable Qty ÷ DRR

    DOC tells you how many days your current inventory will last at the current sales rate.
    """
    )

    # Diagnostics small button
    st.markdown("---")
    st.subheader("Diagnostics")
    if st.button("Check xlsxwriter pivot & COM support"):
        try:
            import xlsxwriter
            st.write("xlsxwriter version:", xlsxwriter.__version__)
            has_pivot_api = hasattr(xlsxwriter.Workbook, "add_pivot_table")
            st.write("xlsxwriter add_pivot_table available:", bool(has_pivot_api))
        except Exception as ex:
            st.write("xlsxwriter not available:", ex)
        # COM check
        if os.name == 'nt':
            try:
                import win32com.client as win32
                st.write("pywin32 available (win32com).")
            except Exception as ex:
                st.write("pywin32 not available:", ex)
        else:
            st.write("COM automation only available on Windows.")

# File upload section
st.header("📁 Upload Files")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Business Report")
    business_file = st.file_uploader(
        "Upload Business Report CSV",
        type=["csv"],
        key="business",
        help="Upload the BusinessReport CSV file",
    )

with col2:
    st.subheader("PM Data")
    pm_file = st.file_uploader(
        "Upload PM Excel/CSV",
        type=["xlsx", "csv"],
        key="pm",
        help="Upload the PM.xlsx or converted CSV file",
    )

with col3:
    st.subheader("Inventory Data")
    inventory_file = st.file_uploader(
        "Upload Manage Inventory CSV",
        type=["csv"],
        key="inventory",
        help="Upload the Manage Inventory CSV file",
    )

# DOC Color Legend
st.header("🎨 DOC Color Legend")
st.markdown(
    """
<div class="doc-legend">
    <div class="legend-item">
        <div class="legend-box" style="background-color: #8B0000;"></div>
        <span><b>0-7 days</b> (Critical)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #FF8C00;"></div>
        <span><b>7-15 days</b> (Low)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #006400;"></div>
        <span><b>15-30 days</b> (Good)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #B8860B;"></div>
        <span><b>30-45 days</b> (Optimal)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #0F52BA;"></div>
        <span><b>45-60 days</b> (High)</span>
    </div>
    <div class="legend-item">
        <div class="legend-box" style="background-color: #8B4513;"></div>
        <span><b>60-90 days</b> (Excess)</span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown("---")

# Process button (main)
if st.button("🚀 Process Data"):
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

                # Ensure SKU in Business Report is string
                if "SKU" in original.columns:
                    original["SKU"] = original["SKU"].astype(str)

                # Read PM file (Excel or CSV)
                if pm_file.name.endswith(".xlsx"):
                    pm = pd.read_excel(pm_file)
                else:
                    pm = pd.read_csv(pm_file)

                inventory = pd.read_csv(inventory_file)

                # Ensure first column of Inventory (SKU-like) is string
                inventory.columns = inventory.columns.str.strip()
                inventory.iloc[:, 0] = inventory.iloc[:, 0].astype(str)

                # Process PM data - select columns C, D, E, F, G (indices 2-6)
                pm = pm.iloc[:, 2:7]
                pm.columns = ["Amazon Sku Name", "D", "Brand Manager", "F", "Brand"]

                # Ensure Amazon Sku Name is string
                pm["Amazon Sku Name"] = pm["Amazon Sku Name"].astype(str)

                # Merge Brand Manager
                original = original.merge(
                    pm[["Amazon Sku Name", "Brand Manager"]],
                    how="left",
                    left_on="SKU",
                    right_on="Amazon Sku Name",
                )

                # Insert Brand Manager column
                if "Title" in original.columns and "Brand Manager" in original.columns:
                    insert_pos = original.columns.get_loc("Title")
                    col = original.pop("Brand Manager")
                    original.insert(insert_pos, "Brand Manager", col)

                # Merge Brand
                original = original.merge(
                    pm[["Amazon Sku Name", "Brand"]],
                    how="left",
                    left_on="SKU",
                    right_on="Amazon Sku Name",
                )

                # Insert Brand column
                if "Title" in original.columns and "Brand" in original.columns:
                    insert_pos = original.columns.get_loc("Title")
                    col = original.pop("Brand")
                    original.insert(insert_pos, "Brand", col)

                # Strip whitespace from inventory columns (already done above)
                inventory.columns = inventory.columns.str.strip()

                # Add fulfillable quantity (11th column, index 10) if exists
                if inventory.shape[1] > 10:
                    return_col = inventory.columns[10]
                    mi_map = inventory.set_index(inventory.columns[0])[return_col]
                    original["afn-fulfillable-quantity"] = original["SKU"].map(mi_map)
                else:
                    original["afn-fulfillable-quantity"] = 0

                # Add reserved quantity (13th column, index 12) if exists
                if inventory.shape[1] > 12:
                    return_col_13 = inventory.columns[12]
                    mi_res_map = inventory.set_index(inventory.columns[0])[return_col_13]
                    original["afn-reserved-quantity"] = original["SKU"].map(mi_res_map)
                else:
                    original["afn-reserved-quantity"] = 0

                # Clean Total Order Items and calculate DRR
                if "Total Order Items" in original.columns:
                    original["Total Order Items"] = (
                        original["Total Order Items"]
                        .astype(str)
                        .str.replace("\u00A0", "", regex=False)
                        .str.replace(",", "", regex=False)
                        .str.replace(r"[^\d\.\-]", "", regex=True)
                    )
                    original["Total Order Items"] = pd.to_numeric(
                        original["Total Order Items"], errors="coerce"
                    )
                else:
                    original["Total Order Items"] = 0

                # Calculate DRR
                original["DRR"] = (original["Total Order Items"] / no_of_days).round(2)

                # Convert fulfillable quantity to numeric
                original["afn-fulfillable-quantity"] = pd.to_numeric(
                    original["afn-fulfillable-quantity"], errors="coerce"
                )

                # Calculate DOC
                original["DOC"] = (original["afn-fulfillable-quantity"] / original["DRR"]).round(2)

                # Replace inf with 0
                original["DOC"] = original["DOC"].replace([float("inf"), float("-inf")], 0)

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

                # Filter columns that exist
                display_cols = [col for col in display_cols if col in original.columns]
                display_df = original[display_cols].copy()

                # Apply styling to DOC column (element-wise)
                styled_df = display_df.style.map(color_doc, subset=["DOC"])

                # Display the dataframe using container width instead of invalid width string
                st.dataframe(styled_df, use_container_width=True, height=600)

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
                    )

                # ---------- Excel export block (process branch) ----------
                with col2:
                    # Offer Excel export that tries to create PivotTable (Brand + (Parent) ASIN)
                    try:
                        parent_col = None
                        for c in original.columns:
                            cle = c.lower().replace(" ", "").replace("_", "").replace("(", "").replace(")", "")
                            if "parent" in cle:
                                parent_col = c
                                break
                    except Exception:
                        parent_col = None

                    st.markdown("### Excel exports (Pivot attempt: Brand + (Parent) ASIN, DOC formatting applied)")
                    over_btn = st.button("📥 Overstock (DOC ↓) — create Excel with Pivot attempt", key="overstock_export")
                    oos_btn = st.button("📥 OOS (DOC ↑) — create Excel with Pivot attempt", key="oos_export")

                    df_to_use = original

                    if over_btn or oos_btn:
                        sort_desc = True if over_btn else False
                        brands = sorted(df_to_use["Brand"].dropna().astype(str).unique().tolist()) if "Brand" in df_to_use.columns else []
                        selected_brands = st.multiselect("Filter brands for Excel export (leave empty = all)", options=brands, default=brands)

                        try:
                            buf, pivot_ok, pivot_info = create_workbook_with_brand_parent_pivot_and_doc_format(df_to_use, sort_desc=sort_desc, sheet_name="Overstock" if sort_desc else "OOS", selected_brands=selected_brands, parent_col=parent_col)
                            final_bytes = buf.getvalue()
                            com_attempted = False
                            com_ok = False
                            com_error = None
                            # Attempt COM only on Windows
                            if os.name == "nt":
                                # check pywin32
                                try:
                                    import win32com.client  # type: ignore
                                    has_pywin = True
                                except Exception:
                                    has_pywin = False

                                if has_pywin:
                                    com_attempted = True
                                    # write to temp file and call COM helper
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                                        tmp_name = tmp.name
                                        tmp.write(final_bytes)
                                        tmp.flush()
                                    try:
                                        ok, com_tb = create_pivot_with_com(tmp_name, data_sheet_name="Overstock" if sort_desc else "OOS", table_name="DataTable", pivot_sheet_name="PivotTable", parent_field_hint="(Parent) ASIN")
                                        if ok:
                                            com_ok = True
                                            with open(tmp_name, "rb") as f:
                                                final_bytes = f.read()
                                        else:
                                            com_ok = False
                                            com_error = com_tb
                                    except Exception as ce:
                                        com_ok = False
                                        com_error = traceback.format_exc()
                                    finally:
                                        try:
                                            os.remove(tmp_name)
                                        except Exception:
                                            pass

                            # deliver final_bytes to user
                            st.download_button(
                                label="Download Excel workbook",
                                data=final_bytes,
                                file_name=f"{'overstock' if sort_desc else 'oos'}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                            # messages about what happened
                            if com_attempted:
                                if com_ok:
                                    st.success("✅ COM automation succeeded — PivotTable, PivotChart and Slicer added in Excel (Windows + Excel).")
                                else:
                                    st.warning("⚠️ COM automation failed — delivered workbook without COM pivot. See pivot attempt info below.")
                                    if com_error:
                                        st.code(com_error)
                            else:
                                # no COM attempted (non-Windows or pywin32 missing)
                                if pivot_ok:
                                    st.success("✅ Programmatic PivotTable (xlsxwriter) created. Open workbook and use Pivot UI to filter by Brand.")
                                else:
                                    st.warning("⚠️ Programmatic pivot creation not possible — exported workbook contains DataTable + PivotSummary + Chart. See 'HowToPivot' sheet for manual steps.")
                                    if pivot_info:
                                        st.info(f"Pivot attempt info: {str(pivot_info)}")
                        except Exception as e:
                            st.error(f"Failed to create workbook: {e}")
                            st.exception(e)
                # ---------- end Excel export block (process branch) ----------

                # Store in session state for persistence
                st.session_state["processed_data"] = original

            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.exception(e)

# Show previously processed data if available
elif "processed_data" in st.session_state:
    st.header("📈 Previously Processed Results")

    original = st.session_state["processed_data"]

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

    # Filter columns that exist
    display_cols = [col for col in display_cols if col in original.columns]
    display_df = original[display_cols].copy()

    # Apply styling to DOC column
    styled_df = display_df.style.map(color_doc, subset=["DOC"])

    # Display the dataframe
    st.dataframe(styled_df, use_container_width=True, height=600)

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
        )

    with col2:
        # Convert to Excel with conditional formatting (previous results)
        try:
            parent_col = None
            for c in original.columns:
                cle = c.lower().replace(" ", "").replace("_", "").replace("(", "").replace(")", "")
                if "parent" in cle:
                    parent_col = c
                    break
        except Exception:
            parent_col = None

        st.markdown("### Excel exports (Pivot attempt: Brand + (Parent) ASIN, DOC formatting applied)")
        over_btn = st.button("📥 Overstock (DOC ↓) — create Excel with Pivot attempt (previous results)", key="overstock_export_prev")
        oos_btn = st.button("📥 OOS (DOC ↑) — create Excel with Pivot attempt (previous results)", key="oos_export_prev")

        df_to_use = original

        if over_btn or oos_btn:
            sort_desc = True if over_btn else False
            brands = sorted(df_to_use["Brand"].dropna().astype(str).unique().tolist()) if "Brand" in df_to_use.columns else []
            selected_brands = st.multiselect("Filter brands for Excel export (leave empty = all)", options=brands, default=brands)

            try:
                buf, pivot_ok, pivot_info = create_workbook_with_brand_parent_pivot_and_doc_format(df_to_use, sort_desc=sort_desc, sheet_name="Overstock" if sort_desc else "OOS", selected_brands=selected_brands, parent_col=parent_col)
                final_bytes = buf.getvalue()
                com_attempted = False
                com_ok = False
                com_error = None
                if os.name == "nt":
                    try:
                        import win32com.client
                        has_pywin = True
                    except Exception:
                        has_pywin = False
                    if has_pywin:
                        com_attempted = True
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                            tmp_name = tmp.name
                            tmp.write(final_bytes)
                            tmp.flush()
                        try:
                            ok, com_tb = create_pivot_with_com(tmp_name, data_sheet_name="Overstock" if sort_desc else "OOS", table_name="DataTable", pivot_sheet_name="PivotTable", parent_field_hint="(Parent) ASIN")
                            if ok:
                                with open(tmp_name, "rb") as f:
                                    final_bytes = f.read()
                                com_ok = True
                            else:
                                com_ok = False
                                com_error = com_tb
                        except Exception as ce:
                            com_ok = False
                            com_error = traceback.format_exc()
                        finally:
                            try:
                                os.remove(tmp_name)
                            except Exception:
                                pass

                st.download_button(
                    label="Download Excel workbook",
                    data=final_bytes,
                    file_name=f"{'overstock' if sort_desc else 'oos'}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if com_attempted:
                    if com_ok:
                        st.success("✅ COM automation succeeded — PivotTable, PivotChart and Slicer added in Excel (Windows + Excel).")
                    else:
                        st.warning("⚠️ COM automation failed — delivered workbook without COM pivot. See pivot attempt info below.")
                        if com_error:
                            st.code(com_error)
                else:
                    if pivot_ok:
                        st.success("✅ Programmatic PivotTable (xlsxwriter) created. Open workbook and use Pivot UI to filter by Brand.")
                    else:
                        st.warning("⚠️ Programmatic pivot creation not possible — exported workbook contains DataTable + PivotSummary + Chart. See 'HowToPivot' sheet for manual steps.")
                        if pivot_info:
                            st.info(f"Pivot attempt info: {str(pivot_info)}")
            except Exception as e:
                st.error(f"Failed to create workbook: {e}")
                st.exception(e)

# Footer
st.markdown("---")
st.markdown(
    """
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>Inventory Analysis Dashboard | Built with Streamlit</p>
</div>
""",
    unsafe_allow_html=True,
)
