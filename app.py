import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import datetime

# =========================================================
# STREAMLIT PAGE CONFIG & STYLING
# =========================================================
st.set_page_config(
    page_title="Fleet Mileage Analytics & Executive Reports",
    page_icon="🚛",
    layout="wide"
)

# Dark Futuristic Custom CSS
st.markdown("""
    <style>
    .main {
        background-color: #0B0F19;
        color: #F8FAFC;
    }
    .stApp {
        background-color: #0B0F19;
    }
    .css-1d39121, .stSidebar {
        background-color: #111827 !important;
    }
    h1, h2, h3 {
        color: #38BDF8 !important;
        font-family: 'Segoe UI', sans-serif;
    }
    .metric-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 18px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
    }
    .stButton>button {
        background: linear-gradient(90deg, #0EA5E9 0%, #2563EB 100%);
        color: white;
        font-weight: bold;
        border: none;
        border-radius: 8px;
        padding: 10px 24px;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #0284C7 0%, #1D4ED8 100%);
        box-shadow: 0 0 15px rgba(14, 165, 233, 0.5);
    }
    </style>
""", unsafe_allow_html=True)

# =========================================================
# HELPER FUNCTIONS & EXCEL BUILDERS
# =========================================================

def clean_and_parse_input(uploaded_file):
    """Reads raw Mileage Report excel and normalizes column headers and types."""
    df_raw = pd.read_excel(uploaded_file)
    
    # Header detection
    if '#' in str(df_raw.columns[0]) or 'Device' in str(df_raw.columns[1]):
        df = df_raw.copy()
    else:
        headers = [str(val).strip() for val in df_raw.iloc[0].values]
        df = df_raw.iloc[1:].copy()
        df.columns = headers

    # Clean whitespace
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    # Cast numeric & date formats
    df['Distance_num'] = pd.to_numeric(df['Distance'], errors='coerce').fillna(0)
    df['Date_parsed'] = pd.to_datetime(df['Date'], errors='coerce')
    
    return df

def build_sunday_report_excel(df, city_name, month_name):
    """Generates Sunday Working Vehicles Excel with a separate sheet for each Sunday."""
    # Filter 0 period and 0 distance
    df_valid = df[
        (df['Period'] != '00:00:00') & 
        (df['Period'] != '0:00:00') & 
        (df['Distance_num'] > 0)
    ].copy()

    # Identify Sundays (dayofweek == 6)
    sundays = sorted(df_valid[df_valid['Date_parsed'].dt.dayofweek == 6]['Date_parsed'].dropna().unique())
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default initial sheet

    # If dataset has no Sundays (e.g. single weekday test file), process available dates
    dates_to_process = sundays if len(sundays) > 0 else sorted(df_valid['Date_parsed'].dropna().unique())

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for idx, dt in enumerate(dates_to_process, 1):
        dt_obj = pd.to_datetime(dt)
        dt_str = dt_obj.strftime('%d-%b-%Y')
        sheet_title = f"Sunday {idx} ({dt_obj.strftime('%b %d')})" if len(sundays) > 0 else f"Day {idx} ({dt_obj.strftime('%b %d')})"
        
        ws = wb.create_sheet(title=sheet_title[:31])
        ws.views.sheetView[0].showGridLines = True

        # Filter for exact Sunday & Sort Distance Descending
        df_day = df_valid[df_valid['Date_parsed'] == dt].sort_values(by='Distance_num', ascending=False)

        # 1. Main Title Header Block
        ws.merge_cells('A1:K1')
        t_cell = ws['A1']
        t_cell.value = f"SUNDAY WORKING VEHICLES MILEAGE REPORT — {city_name.upper()}"
        t_cell.font = Font(name='Segoe UI', size=13, bold=True, color='FFFFFF')
        t_cell.fill = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
        t_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 38

        # 2. Subtitle Meta Bar
        ws.merge_cells('A2:K2')
        s_cell = ws['A2']
        s_cell.value = f"Date: {dt_str}   |   Month: {month_name}   |   Active Sunday Vehicles: {len(df_day)}"
        s_cell.font = Font(name='Segoe UI', size=9.5, italic=True, color='38BDF8')
        s_cell.fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
        s_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 22

        # 3. Table Column Headers
        headers = ['Sr #', 'Device ID', 'Reg #', 'Group / Region', 'Vehicle Type', 'Purpose', 'TH Kms', 'TH Time', 'Date', 'Distance (Km)', 'Period (HH:MM)']
        ws.row_dimensions[4].height = 26
        
        h_fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
        h_font = Font(name='Segoe UI', size=9.5, bold=True, color='FFFFFF')

        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=4, column=col_idx, value=h)
            c.fill = h_fill
            c.font = h_font
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = thin_border

        # 4. Data Rows
        start_row = 5
        for r_idx, (_, row) in enumerate(df_day.iterrows(), start_row):
            ws.row_dimensions[r_idx].height = 20
            bg_color = 'F8FAFC' if r_idx % 2 == 0 else 'FFFFFF'
            r_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

            d_vals = [
                r_idx - start_row + 1,
                row['Device'],
                row['Reg#'],
                row['Group'],
                row['Vehicle Type'],
                row['Purpose'],
                float(row['TH Kms']) if str(row['TH Kms']).replace('.','',1).isdigit() else row['TH Kms'],
                row['TH Tme'],
                row['Date'],
                float(row['Distance_num']),
                row['Period']
            ]

            for c_idx, val in enumerate(d_vals, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.fill = r_fill
                cell.border = thin_border
                cell.font = Font(name='Segoe UI', size=9)

                if c_idx in [1, 2, 3, 7, 8, 9, 11]:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif c_idx == 10:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0.00'
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

        # 5. Summary Row
        tot_row = start_row + len(df_day)
        ws.row_dimensions[tot_row].height = 24
        tot_fill = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')
        tot_font = Font(name='Segoe UI', size=9.5, bold=True, color='0F172A')

        ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=9)
        tot_label = ws.cell(row=tot_row, column=1, value="TOTAL SUNDAY DISTANCE (KM)")
        tot_label.font = tot_font
        tot_label.alignment = Alignment(horizontal='right', vertical='center')

        for c in range(1, 10):
            ws.cell(row=tot_row, column=c).border = thin_border
            ws.cell(row=tot_row, column=c).fill = tot_fill

        if len(df_day) > 0:
            tot_val = ws.cell(row=tot_row, column=10, value=f"=SUM(J{start_row}:J{tot_row-1})")
        else:
            tot_val = ws.cell(row=tot_row, column=10, value=0)
            
        tot_val.font = tot_font
        tot_val.fill = tot_fill
        tot_val.alignment = Alignment(horizontal='right', vertical='center')
        tot_val.number_format = '#,##0.00'
        tot_val.border = thin_border

        empty_c = ws.cell(row=tot_row, column=11, value="")
        empty_c.fill = tot_fill
        empty_c.border = thin_border

        # Auto column width adjustment
        for col in ws.columns:
            max_l = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_l + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def build_evening_report_excel(df, city_name, month_name, target_vehicles):
    """Generates Evening Shift Report Excel with stacked vehicle sections in one sheet."""
    # Filter 0 period and 0 distance
    df_valid = df[
        (df['Period'] != '00:00:00') & 
        (df['Period'] != '0:00:00') & 
        (df['Distance_num'] > 0)
    ].copy()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Evening Shift Report"
    ws.views.sheetView[0].showGridLines = True

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # 1. Main Header Title
    ws.merge_cells('A1:K1')
    t_cell = ws['A1']
    t_cell.value = f"EVENING VEHICLES SHIFT MILEAGE REPORT — {city_name.upper()}"
    t_cell.font = Font(name='Segoe UI', size=13, bold=True, color='FFFFFF')
    t_cell.fill = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
    t_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 38

    ws.merge_cells('A2:K2')
    s_cell = ws['A2']
    s_cell.value = f"Month Period: {month_name}   |   Target Evening Shift Fleet: {', '.join([v.upper() for v in target_vehicles])}"
    s_cell.font = Font(name='Segoe UI', size=9.5, italic=True, color='F5A623')
    s_cell.fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
    s_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 22

    current_row = 4

    headers = ['Sr #', 'Device ID', 'Reg #', 'Group / Region', 'Vehicle Type', 'Purpose', 'TH Kms', 'TH Time', 'Date', 'Distance (Km)', 'Period (HH:MM)']

    for veh in target_vehicles:
        # Filter for specific registration number
        df_veh = df_valid[df_valid['Reg#'].str.upper() == veh.upper()].sort_values(by='Distance_num', ascending=False)

        # Vehicle Section Banner
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
        v_cell = ws.cell(row=current_row, column=1, value=f"VEHICLE REGISTRATION: {veh.upper()}   (Total Monthly Entries: {len(df_veh)})")
        v_cell.font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
        v_cell.fill = PatternFill(start_color='334155', end_color='334155', fill_type='solid')
        v_cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[current_row].height = 25
        current_row += 1

        # Section Header Row
        ws.row_dimensions[current_row].height = 24
        h_fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
        h_font = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')

        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=current_row, column=col_idx, value=h)
            c.fill = h_fill
            c.font = h_font
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = thin_border

        current_row += 1

        if len(df_veh) == 0:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
            empty_info = ws.cell(row=current_row, column=1, value="No active evening shift entries recorded for this vehicle during the selected month.")
            empty_info.font = Font(name='Segoe UI', size=9, italic=True, color='64748B')
            empty_info.alignment = Alignment(horizontal='center', vertical='center')
            for c in range(1, 12):
                ws.cell(row=current_row, column=c).border = thin_border
            ws.row_dimensions[current_row].height = 22
            current_row += 1
        else:
            start_v_row = current_row
            for r_idx, (_, row) in enumerate(df_veh.iterrows(), start_v_row):
                ws.row_dimensions[r_idx].height = 20
                bg_color = 'F8FAFC' if r_idx % 2 == 0 else 'FFFFFF'
                r_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

                d_vals = [
                    r_idx - start_v_row + 1,
                    row['Device'],
                    row['Reg#'],
                    row['Group'],
                    row['Vehicle Type'],
                    row['Purpose'],
                    float(row['TH Kms']) if str(row['TH Kms']).replace('.','',1).isdigit() else row['TH Kms'],
                    row['TH Tme'],
                    row['Date'],
                    float(row['Distance_num']),
                    row['Period']
                ]

                for c_idx, val in enumerate(d_vals, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.fill = r_fill
                    cell.border = thin_border
                    cell.font = Font(name='Segoe UI', size=9)

                    if c_idx in [1, 2, 3, 7, 8, 9, 11]:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif c_idx == 10:
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                        cell.number_format = '#,##0.00'
                    else:
                        cell.alignment = Alignment(horizontal='left', vertical='center')

                current_row += 1

            # Vehicle Total Summary Row
            tot_row = current_row
            ws.row_dimensions[tot_row].height = 23
            tot_fill = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')
            tot_font = Font(name='Segoe UI', size=9, bold=True, color='0F172A')

            ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=9)
            tot_label = ws.cell(row=tot_row, column=1, value=f"TOTAL MONTHLY MILEAGE FOR {veh.upper()} (KM)")
            tot_label.font = tot_font
            tot_label.alignment = Alignment(horizontal='right', vertical='center')

            for c in range(1, 10):
                ws.cell(row=tot_row, column=c).border = thin_border
                ws.cell(row=tot_row, column=c).fill = tot_fill

            tot_val = ws.cell(row=tot_row, column=10, value=f"=SUM(J{start_v_row}:J{tot_row-1})")
            tot_val.font = tot_font
            tot_val.fill = tot_fill
            tot_val.alignment = Alignment(horizontal='right', vertical='center')
            tot_val.number_format = '#,##0.00'
            tot_val.border = thin_border

            empty_c = ws.cell(row=tot_row, column=11, value="")
            empty_c.fill = tot_fill
            empty_c.border = thin_border

            current_row += 1

        current_row += 2  # Visual space before next vehicle section

    # Auto column width adjustment
    for col in ws.columns:
        max_l = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_l + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================================================
# UI LAYOUT & APPLICATION CONTROLLER
# =========================================================

st.title("⚡ Executive Fleet Mileage Report Generator")
st.markdown("Upload raw Excel mileage data to generate executive Sunday Working and Evening Shift reports.")

st.divider()

col_left, col_right = st.columns([1, 2])

with col_left:
    st.subheader("📋 Configuration Parameters")
    
    uploaded_file = st.file_uploader(
        "Upload Monthly Raw Mileage File (.xlsx)",
        type=["xlsx", "xls"],
        help="Upload the raw Excel dump containing complete month mileage entries."
    )
    
    tehsil_choice = st.selectbox(
        "Select Tehsil / City Name",
        ["Kamoke", "Nowshera Virkan", "Custom"]
    )
    
    if tehsil_choice == "Custom":
        city_name = st.text_input("Enter Tehsil/City Name", value="Kamoke")
    else:
        city_name = tehsil_choice

    month_name = st.selectbox(
        "Select Month",
        ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
        index=6
    )
    
    st.markdown("---")
    st.markdown("##### 🚗 Evening Target Vehicles Configuration")
    if city_name.lower() == "kamoke":
        default_vehs = "GBA-915, GBA-917"
    elif city_name.lower() in ["nowshera virkan", "nowshera"]:
        default_vehs = "RIC-159, RIC-165, TT-240"
    else:
        default_vehs = "GBA-915, GBA-917"

    veh_input = st.text_input(
        "Target Registration Numbers (Comma Separated)",
        value=default_vehs
    )
    target_vehicles = [v.strip() for v in veh_input.split(",") if v.strip()]

with col_right:
    st.subheader("📊 Data Inspection & Report Generation")
    
    if uploaded_file is not None:
        try:
            df_clean = clean_and_parse_input(uploaded_file)
            
            st.success("✅ Excel file loaded & parsed successfully!")
            
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("Total Raw Entries", len(df_clean))
            with m2:
                active_entries = len(df_clean[(df_clean['Period'] != '00:00:00') & (df_clean['Period'] != '0:00:00')])
                st.metric("Active Entries (>0 Time)", active_entries)
            with m3:
                unique_vehs = df_clean['Reg#'].nunique()
                st.metric("Unique Fleet Count", unique_vehs)

            st.markdown("##### Raw Data Snapshot (First 5 Rows)")
            st.dataframe(df_clean[['Device', 'Reg#', 'Group', 'Vehicle Type', 'Date', 'Distance', 'Period']].head(5), use_container_width=True)
            
            st.divider()
            
            # GENERATE REPORTS
            st.markdown("### 📥 Download Executive Reports")
            
            c_btn1, c_btn2 = st.columns(2)
            
            # File 1: Sunday Working Report
            sunday_filename = f"Mileage Report {month_name} All Sunday Working vehicles report Tehsil {city_name}.xlsx"
            sunday_bytes = build_sunday_report_excel(df_clean, city_name, month_name)
            
            with c_btn1:
                st.download_button(
                    label="📊 Download Sunday Working Report",
                    data=sunday_bytes,
                    file_name=sunday_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.caption(f"📁 `{sunday_filename}`")
                
            # File 2: Evening Shift Report
            evening_filename = f"Mileage report {month_name} evening vehicles shift report tehsil {city_name}.xlsx"
            evening_bytes = build_evening_report_excel(df_clean, city_name, month_name, target_vehicles)
            
            with c_btn2:
                st.download_button(
                    label="🌙 Download Evening Shift Report",
                    data=evening_bytes,
                    file_name=evening_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.caption(f"📁 `{evening_filename}`")

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.info("Please make sure the uploaded file is a valid raw Mileage Report excel.")
    else:
        st.info("👈 Upload your Excel file from the left sidebar panel to begin processing.")
