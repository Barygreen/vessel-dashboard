import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import io
from datetime import datetime

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Vessel Schedule Workspace", layout="wide")

# --- REGION CONFIGURATIONS ---
REGION_CONFIG = {
    "USA East": ["NY", "Charleston", "Sav", "Norfolk"],
    "USA West": ["LA"],
    "China": ["Shanghai", "Qingdao", "Ningbo"],
    "Europe": ["Hamburg", "Rotterdam", "Antwerp"],
    "Spain": ["Barcelona", "Valencia"],
    "Vietnam": ["Cat Lai"],
    "Indonesia": ["Jakarta"],
    "Mexico": ["Manzanillo"],
    "Malaysia": ["Port Klang", "Penang", "Pasir Gudang"],
    "Bangladesh": ["Chittagong"],
    "Other": ["Colombo"]
}

PORT_REGION_MAP = {
    "NEW YORK": ("USA East", "NY"), "NEWARK": ("USA East", "NY"),
    "CHARLESTON": ("USA East", "Charleston"), "SAVANNAH": ("USA East", "Sav"), "NORFOLK": ("USA East", "Norfolk"),
    "LOS ANGELES": ("USA West", "LA"), "LONG BEACH": ("USA West", "LA"),
    "SHANGHAI": ("China", "Shanghai"), "NINGBO": ("China", "Ningbo"), "QINGDAO": ("China", "Qingdao"),
    "HAMBURG": ("Europe", "Hamburg"), "ROTTERDAM": ("Europe", "Rotterdam"), "ANTWERP": ("Europe", "Antwerp"),
    "BARCELONA": ("Spain", "Barcelona"), "VALENCIA": ("Spain", "Valencia"),
    "CAT LAI": ("Vietnam", "Cat Lai"), "HO CHI MINH": ("Vietnam", "Cat Lai"),
    "JAKARTA": ("Indonesia", "Jakarta"), "MANZANILLO": ("Mexico", "Manzanillo"),
    "PORT KLANG": ("Malaysia", "Port Klang"), "PENANG": ("Malaysia", "Penang"), "PASIR GUDANG": ("Malaysia", "Pasir Gudang"),
    "CHITTAGONG": ("Bangladesh", "Chittagong"), "COLOMBO": ("Other", "Colombo")
}

# --- DYNAMIC DATE ENGINE ---
def normalize_date(val, default_year):
    if pd.isna(val) or not str(val).strip(): return "-"
    s = str(val).strip()
    
    # 1. Strip structural noise from carriers (like CY Cutoff: or SI Cutoff:)
    s = re.sub(r'(?i)(CY|SI)\s*Cutoff:', '', s)
    
    # 2. Strip weekdays (e.g., (Mon), (Tue)) and parentheses
    s = re.sub(r'(?i)\([a-z]{3}\)', '', s)
    s = re.sub(r'(?i)\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\b', '', s)
    s = s.replace('()', '').replace('(/)', '')
    
    # 3. Strip times (e.g., 23:00)
    s = re.sub(r'\d{1,2}:\d{2}', '', s)
    
    s = re.sub(r'\s+', ' ', s).strip()
    if '|' in s: s = s.split('|')[0].strip()
    
    # 4. Look for a complete date with a year
    match = re.search(r'(\d{1,2}[-/\s][A-Za-z]{3,9}[-/\s]\d{2,4})|(\d{4}[-/\s]\d{1,2}[-/\s]\d{1,2})', s)
    if match:
        s_clean = match.group(0)
    else:
        # 5. Fallback for OOCL/ONE "DD MMM" format (e.g., "04 Jul") -> Force the dynamically extracted year
        match_no_year = re.search(r'(\d{1,2})[-/\s]([A-Za-z]{3,9})', s)
        if match_no_year and default_year:
            s_clean = f"{match_no_year.group(1)} {match_no_year.group(2)} {default_year}"
        else:
            return "-"

    # Format Date to DD-Mon-YY
    formats_to_try = [
        '%d %b %Y', '%Y-%m-%d', '%d-%b-%Y', '%d-%B-%Y', 
        '%Y/%m/%d', '%d/%m/%Y', '%d-%b-%y', '%d %B %Y', '%d-%b %Y'
    ]
    for fmt in formats_to_try:
        try:
            return datetime.strptime(s_clean, fmt).strftime('%d-%b-%y')
        except ValueError:
            continue
    return "-"

def clean_vessel_name(name):
    if pd.isna(name): return "-"
    return re.sub(r'\[.*?\]', '', str(name).upper()).strip()

def extract_voyage(voy):
    if pd.isna(voy): return "-"
    return str(voy).upper().strip().lstrip('0')

# --- PARSING ENGINES ---
def parse_file(uploaded_file):
    name = uploaded_file.name.lower()
    rows = []
    try:
        if name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, on_bad_lines='skip')
        else:
            engine = 'xlrd' if name.endswith('.xls') else 'openpyxl'
            df = pd.read_excel(uploaded_file, engine=engine)
            
    except Exception as e:
        error_msg = str(e)
        if "Expected BOF record" in error_msg or "<html" in error_msg.lower() or "<div" in error_msg.lower():
            try:
                uploaded_file.seek(0)
                tables = pd.read_html(uploaded_file)
                df = tables[0]
            except Exception as html_err:
                st.error(f"Failed to read {uploaded_file.name} as an HTML table: {html_err}")
                return []
        else:
            st.error(f"Failed to read {uploaded_file.name}: {e}")
            return []

    cols_str = " ".join([str(c).lower().strip() for c in df.columns])
    sample_blob = df.iloc[:5].to_string().lower()
    
    # Extract dynamic year for future-proofing missing-year dates
    year_match = re.search(r'20\d{2}', cols_str + " " + sample_blob)
    file_year = year_match.group(0) if year_match else str(datetime.now().year)

    try:
        # 1. CMA CGM Engine
        if 'vessel name' in cols_str and 'voyage ref.' in cols_str:
            for _, r in df.iterrows():
                dest = str(r.get('Destination', '')).upper()
                reg, pod = next((v for k, v in PORT_REGION_MAP.items() if k in dest), ("USA East", "NY"))
                
                service = str(r.get('Service Name', '-')).strip()
                if service.lower() == 'nan': service = "-"
                
                ts_val = str(r.get('Transhipment', '')).strip()
                if ts_val in ['1.0', '1', 'Yes', 'Y']:
                    service = f"{service} (T/S)" if service != "-" else "Transshipment"
                
                rows.append({
                    "region": reg, "carrier": "CMA CGM", "vessel_name": clean_vessel_name(r.get('Vessel name', '-')), 
                    "voyage_no": extract_voyage(r.get('Voyage ref.', '-')),
                    "doc_cutoff": normalize_date(r.get('SI cut-off', '-'), file_year), "fcl_cutoff": normalize_date(r.get('Standard booking cut-off', '-'), file_year),
                    "etd": normalize_date(r.get('Departure Date', '-'), file_year), "etas": {pod: normalize_date(r.get('Arrival Date', '-'), file_year)},
                    "tts": {pod: str(r.get('Transit Time', '-')).split('.')[0]}, "service": service
                })

        # 2. Hapag Lloyd Engine
        elif 'doc cut-off/fcl cut-off/vgm cut-off' in cols_str:
            for _, r in df.iterrows():
                dest = str(r.get('Arrival location', '')).upper()
                reg, pod = next((v for k, v in PORT_REGION_MAP.items() if k in dest), ("USA East", "NY"))
                cell = str(r.get('Doc Cut-off/FCL Cut-off/VGM Cut-off', ''))
                splits = cell.split('/') if '/' in cell else [cell, cell]
                
                service = str(r.get('Service name', '-')).strip()
                if service.lower() == 'nan': service = "-"

                rows.append({
                    "region": reg, "carrier": "Hapag Lloyd", "vessel_name": str(r.get('Vessel name', '-')).upper(), "voyage_no": "-",
                    "doc_cutoff": normalize_date(splits[0], file_year), "fcl_cutoff": normalize_date(splits[1] if len(splits)>1 else splits[0], file_year),
                    "etd": normalize_date(r.get('Departure time and date', '-'), file_year), "etas": {pod: normalize_date(r.get('Arrival time and date', '-'), file_year)},
                    "tts": {pod: str(r.get('Transit time', '-')).split('.')[0]}, "service": service
                })

        # 3. ONE Line Engine
        elif 'unnamed: 10' in cols_str or 'go to one schedule page' in cols_str:
            for i in range(2, len(df)):
                r = df.iloc[i]
                vv = str(r.iloc[10] if len(r) > 10 else "").split('\n')[-1].strip()
                if not vv or vv.lower().startswith('vessel'): continue
                
                parts = vv.rsplit(' ', 1)
                dest = str(r.iloc[4]).upper()
                reg, pod = next((v for k, v in PORT_REGION_MAP.items() if k in dest), ("USA East", "NY"))
                
                service = str(r.iloc[9]).replace('\n', ' ').strip()
                if service.lower() == 'nan': service = "-"

                ts_raw = str(r.iloc[2]).strip()
                ts_port = re.sub(r'\d{4}-\d{2}-\d{2}.*', '', ts_raw).strip().title()
                if pd.notna(r.iloc[2]) and ts_port and ts_port.lower() != "nan":
                    service = f"{service} (T/S {ts_port})" if service != "-" else f"T/S {ts_port}"

                rows.append({
                    "region": reg, "carrier": "ONE Line", "vessel_name": parts[0].upper(), "voyage_no": extract_voyage(parts[1] if len(parts)>1 else "-"),
                    "doc_cutoff": normalize_date(str(r.iloc[6]).split('\n')[0], file_year), "fcl_cutoff": normalize_date(str(r.iloc[7]).split('\n')[0], file_year),
                    "etd": normalize_date(str(r.iloc[1]).split('\n')[-1], file_year), "etas": {pod: normalize_date(str(r.iloc[4]).split('\n')[-1], file_year)},
                    "tts": {pod: str(r.iloc[5]).split(' ')[0]}, "service": service
                })

        # 4. OOCL Engine
        elif 'vessel voyage' in sample_blob or 'unnamed: 8' in cols_str:
            for i in range(7, len(df) - 1, 2):
                r_top = df.iloc[i]
                r_bot = df.iloc[i+1]
                vv = str(r_top.iloc[8] if len(r_top) > 8 else "").strip()
                if not vv or vv.lower().startswith('vessel'): continue
                
                parts = vv.rsplit(' ', 1)
                dest = str(r_top.iloc[2]).upper()
                reg, pod = next((v for k, v in PORT_REGION_MAP.items() if k in dest), ("USA East", "NY"))
                
                fcl_raw = str(r_top.iloc[4]).replace('CY Cutoff:', '')
                doc_raw = str(r_bot.iloc[4]).replace('SI Cutoff:', '')
                
                service = str(r_top.iloc[7]).strip()
                if service.lower() == 'nan': service = "-"

                ts_raw = str(r_top.iloc[1]).strip().title()
                if pd.notna(r_top.iloc[1]) and ts_raw and ts_raw.lower() != "nan":
                    service = f"{service} (T/S {ts_raw})" if service != "-" else f"T/S {ts_raw}"

                rows.append({
                    "region": reg, "carrier": "OOCL", "vessel_name": parts[0].upper(), "voyage_no": extract_voyage(parts[1] if len(parts)>1 else "-"),
                    "doc_cutoff": normalize_date(doc_raw, file_year), "fcl_cutoff": normalize_date(fcl_raw, file_year),
                    "etd": normalize_date(r_bot.iloc[5], file_year), "etas": {pod: normalize_date(r_bot.iloc[2], file_year)},
                    "tts": {pod: str(r_top.iloc[3]).split(' ')[0]}, "service": service
                })

        # 5. Maersk Engine
        elif 'deadline cy' in cols_str:
            for _, r in df.iterrows():
                dest = str(r.get('Arrival', '')).upper().split('-')[0].strip()
                reg, pod = next((v for k, v in PORT_REGION_MAP.items() if k in dest), ("USA East", "NY"))
                
                rows.append({
                    "region": reg, "carrier": "Maersk", "vessel_name": str(r.get('Vessel', '-')).upper(), "voyage_no": extract_voyage(r.get('Voyage Number', '-')),
                    "doc_cutoff": normalize_date(r.get('Deadline SI - NON AMS', '-'), file_year), "fcl_cutoff": normalize_date(r.get('Deadline CY', '-'), file_year),
                    "etd": normalize_date(r.get('Departure Date', '-'), file_year), "etas": {pod: normalize_date(r.get('Arrival Date', '-'), file_year)},
                    "tts": {pod: str(r.get('Transit Time', '-')).split(' ')[0]}, "service": "-"
                })
    except Exception as e:
        st.warning(f"Error parsing layout in {name}: {e}")
    
    return rows

# --- FRONTEND UI ---
st.title("🚢 Master Vessel Schedule Aggregator")
st.markdown("Drag and drop raw carrier exports. The engine will automatically identify the carrier, fix the dates, and build the 9-tab corporate Excel file.")

uploaded_files = st.file_uploader("Upload Carrier Exports (.xlsx, .xls, .csv)", accept_multiple_files=True)

if uploaded_files:
    database = {}
    with st.spinner('Parsing matrices and aligning timelines...'):
        for f in uploaded_files:
            extracted = parse_file(f)
            for r in extracted:
                v_key = re.sub(r'[^A-Z0-9]', '', r['vessel_name'])
                voy_key = re.sub(r'[^A-Z0-9]', '', r['voyage_no'])
                comp_key = f"{r['region']}|||{r['carrier']}|||{v_key}|||{voy_key}"
                
                if comp_key not in database:
                    database[comp_key] = r
                else:
                    database[comp_key]['etas'].update(r['etas'])
                    database[comp_key]['tts'].update(r['tts'])
                    if database[comp_key]['doc_cutoff'] == "-": database[comp_key]['doc_cutoff'] = r['doc_cutoff']
                    if database[comp_key]['fcl_cutoff'] == "-": database[comp_key]['fcl_cutoff'] = r['fcl_cutoff']
                    if database[comp_key].get('service', '-') == "-" and r.get('service', '-') != "-": 
                        database[comp_key]['service'] = r['service']

    all_items = list(database.values())
    st.success(f"✅ Successfully compiled {len(all_items)} unique vessel rotations!")
    
    # --- BUILD EXCEL FILE IN MEMORY ---
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    fill_header = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")
    fill_meta = PatternFill(start_color="EAD1DC", end_color="EAD1DC", fill_type="solid")
    fill_sop = PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for reg_name, ports in REGION_CONFIG.items():
        ws = wb.create_sheet(title=reg_name)
        ws.append(["JUNE 2026 VESSEL SCHEDULE"])
        ws.append(["STANDARD OPERATING PARAMETERS"])
        ws.append(["· Gate Open: 5 Days Prior to Vessel ETD"])
        ws.append(["· Gate Close: 2 Days Prior to Vessel ETD"])
        ws.append(["· Documentation Cut-Off: 48 Hours Prior to Vessel ETD"])
        ws.append([])
        ws.append([f"{reg_name.upper()} MATRIX OVERVIEW"])
        ws.append(["POL: Mundra / Nhava Sheva"])
        ws.append([f"POD Target Gateways: {', '.join(ports)}"])
        ws.append([])
        
        headers = ["Carrier", "Vessel Name", "Voyage No.", "Doc Cut-Off", "FCL Cut-Off", "ETD Local"]
        for p in ports: headers.extend([f"ETA {p}", f"{p} TT"])
        headers.append("Service")
        ws.append(headers)
        
        # Sort rows: Group by Carrier FIRST, then Date
        reg_rows = [r for r in all_items if r['region'] == reg_name]
        def sort_key(x):
            try: dt = datetime.strptime(x['etd'], '%d-%b-%y')
            except: dt = datetime(9999, 12, 31)
            return (x['carrier'].lower(), dt)
            
        reg_rows.sort(key=sort_key)

        for r in reg_rows:
            cells = [r['carrier'], r['vessel_name'], r['voyage_no'], r['doc_cutoff'], r['fcl_cutoff'], r['etd']]
            for p in ports: cells.extend([r['etas'].get(p, "-"), r['tts'].get(p, "-")])
            cells.append(r.get('service', '-'))
            ws.append(cells)

        # Style the Excel Sheet
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=len(headers)), 1):
            if r_idx in [1, 2, 7, 8, 9]:
                for cell in row: 
                    cell.fill = fill_meta
                    cell.font = Font(bold=True)
                    cell.alignment = align_left
            elif r_idx in [3, 4, 5]:
                for cell in row: 
                    cell.fill = fill_sop
                    cell.alignment = align_left
            elif r_idx == 11:
                for cell in row: 
                    cell.fill = fill_header
                    cell.font = Font(bold=True)
                    cell.border = border
                    cell.alignment = align_center
            elif r_idx > 11:
                for cell in row: 
                    cell.border = border
                    cell.alignment = align_center
                
        # Dynamic Auto-Fitter for Column Widths
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    lines = str(cell.value).split('\n')
                    for line in lines:
                        if len(line) > max_len: max_len = len(line)
            adjusted_width = min(max_len + 4, 45)
            adjusted_width = max(adjusted_width, 14)
            ws.column_dimensions[col_letter].width = adjusted_width

        # --- VERTICAL MERGE ENGINE FOR CARRIER COLUMN ---
        if ws.max_row >= 12:
            start_merge_row = 12
            current_carrier = ws.cell(row=12, column=1).value
            
            for r in range(13, ws.max_row + 2):
                cell_val = ws.cell(row=r, column=1).value if r <= ws.max_row else None
                if cell_val != current_carrier:
                    if r - 1 > start_merge_row:
                        ws.merge_cells(start_row=start_merge_row, start_column=1, end_row=r-1, end_column=1)
                    start_merge_row = r
                    current_carrier = cell_val

    out = io.BytesIO()
    wb.save(out)
    
    st.markdown("---")
    st.download_button(
        label="📥 Download Compiled Master Excel Sheet",
        data=out.getvalue(),
        file_name=f"{datetime.now().strftime('%B')} Vessel Schedule Master.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
