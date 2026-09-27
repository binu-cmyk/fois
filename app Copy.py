import os
import sys
import time
import datetime
import ssl
import subprocess
import requests
import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Ignore SSL warnings for Indian Railways sites as their certs are often invalid or outdated
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = "fois_tracker_secret_key_123"

# Global session to maintain JSESSIONID and security cookies across requests
session = None
last_session_time = 0

# Global mappings and modification tracker for dynamic reloading from Excel
_last_lookup_load_time = 0
_last_mtime = 0
SIDING_MAP = {}
COMMODITY_MAP = {}

def get_lookup_mappings():
    global _last_lookup_load_time, _last_mtime, SIDING_MAP, COMMODITY_MAP
    excel_path = os.path.join(os.path.dirname(__file__), "Rail Siding Codes.xlsm")
    if not os.path.exists(excel_path):
        print(f"Lookup file not found at: {excel_path}. Mappings will be empty.")
        return SIDING_MAP, COMMODITY_MAP, f"Siding codes spreadsheet not found at: {excel_path}."
        
    warning_msg = None
    try:
        mtime = os.path.getmtime(excel_path)
        # Reload if file changed or if we haven't loaded yet
        if mtime > _last_mtime or not SIDING_MAP:
            print(f"Detected change or first load for {excel_path}. Reloading mappings...")
            s_map = {}
            c_map = {}
            # Load Siding Codes
            siding_df = pd.read_excel(excel_path, sheet_name='Siding Code')
            for _, row in siding_df.iterrows():
                if 'Name' in row and 'Siding code' in row:
                    name = str(row['Name']).strip()
                    code = str(row['Siding code']).strip().upper()
                    if name and code and name != 'nan' and code != 'nan':
                        s_map[code] = name
                        
            # Load Commodity Codes
            comm_df = pd.read_excel(excel_path, sheet_name='Commodity Code')
            for _, row in comm_df.iterrows():
                if 'Commodity Code' in row and 'COMMODITY NAME' in row:
                    code = str(row['Commodity Code']).strip().upper()
                    name = str(row['COMMODITY NAME']).strip()
                    if code and name and code != 'nan' and name != 'nan':
                        c_map[code] = name
                        
            SIDING_MAP = s_map
            COMMODITY_MAP = c_map
            _last_mtime = mtime
            print(f"Successfully loaded {len(SIDING_MAP)} siding codes and {len(COMMODITY_MAP)} commodity codes.")
    except PermissionError:
        warning_msg = "Siding codes sheet 'Rail Siding Codes.xlsm' is currently open in Excel and locked. Please close it so the dashboard can load your new codes!"
        print(f"Error loading lookup mappings: {warning_msg}")
    except Exception as e:
        warning_msg = f"Failed to reload siding/commodity codes: {e}"
        print(f"Error loading lookup mappings: {warning_msg}")
        
    return SIDING_MAP, COMMODITY_MAP, warning_msg

def enrich_siding(val, siding_map):
    if not val or pd.isna(val):
        return val
    val_str = str(val).strip()
    val_upper = val_str.upper()
    if val_upper in siding_map:
        return siding_map[val_upper]
    return val_str

def enrich_commodity(val, commodity_map):
    if not val or pd.isna(val):
        return val
    val_str = str(val).strip()
    val_upper = val_str.upper()
    if val_upper in commodity_map:
        return commodity_map[val_upper]
    return val_str

def load_last_saved_data(sheet_name):
    excel_path = os.path.abspath("FOIS_MIS_Report.xlsx")
    if not os.path.exists(excel_path):
        return None, None
        
    try:
        print(f"Loading last saved data from sheet '{sheet_name}' in {excel_path}...")
        # Skip the corporate title banner rows (rows 1-3)
        # Header is on row 4 (skiprows=3)
        df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=3)
        
        # Clean column names in case they have spaces
        df.columns = [" ".join(str(c).split()) for c in df.columns]
        
        # Get file modification time as formatted string
        mtime = os.path.getmtime(excel_path)
        mtime_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        
        # Convert all columns to string to prevent numerical float conversion issues on display
        for col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '')
            
        return df, mtime_str
    except Exception as e:
        print(f"Error loading last saved Excel data: {e}")
        return None, None

def is_utcl_value(val, s_map_reverse):
    if not val or pd.isna(val):
        return False
    val_str = str(val).strip().upper()
    
    # Check if code itself starts with UCL/UTC or contains UTCL/ULTRATECH/ULTRA
    if val_str.startswith(("UCL", "UTC")) or any(term in val_str for term in ["UTCL", "ULTRATECH", "ULTRA"]):
        return True
        
    # Check if mapped name's code starts with UCL/UTC
    if val_str in s_map_reverse:
        code = s_map_reverse[val_str]
        if code.startswith(("UCL", "UTC")):
            return True
            
    return False

def get_active_session():
    global session, last_session_time
    # Reset session if older than 15 minutes to prevent timeout
    now = time.time()
    if session is None or (now - last_session_time > 900):
        print("Initializing new FOIS HTTP session...")
        session = requests.Session()
        session.verify = False  # Skip SSL verification
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
        })
        
        # Hit the landing page first to establish cookies
        try:
            url = "https://www.fois.indianrail.gov.in/FOISWebPortal/pages/FWP_ODROtsgDtls.jsp"
            for attempt in range(3):
                try:
                    session.get(url, timeout=30)
                    break
                except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                    if attempt == 2:
                        raise e
                    print(f"Landing page attempt {attempt+1} failed. Retrying in 2 seconds...")
                    time.sleep(2)
            last_session_time = time.time()
            print("Session established successfully. Cookies:")
            for cookie in session.cookies:
                print(f"  {cookie.name} = {cookie.value}")
        except Exception as e:
            print(f"Error establishing landing page session: {e}")
            
    return session

def parse_fois_table(html, qry="ODR_RK_OTSG"):
    soup = BeautifulSoup(html, 'html.parser')
    tables = soup.find_all('table')
    if not tables:
        return None
    
    # We want to find the main data table (usually the largest table with records)
    main_table = None
    max_rows = 0
    for table in tables:
        rows = table.find_all('tr')
        if len(rows) > max_rows:
            max_rows = len(rows)
            main_table = table
            
    if not main_table or max_rows <= 1:
        return None
        
    rows = main_table.find_all('tr')
    
    # 1. Extract headers dynamically handling colspan and rowspan from merged headers
    first_row_cells = rows[0].find_all(['th', 'td'])
    C = sum(int(cell.get('colspan', 1)) for cell in first_row_cells)
    
    grid = [[None] * C for _ in range(2)]
    col_idx = 0
    for cell in first_row_cells:
        text = cell.text.strip().replace('\n', ' ').replace('\r', '')
        colspan = int(cell.get('colspan', 1))
        rowspan = int(cell.get('rowspan', 1))
        for c in range(colspan):
            grid[0][col_idx + c] = text
            if rowspan > 1:
                grid[1][col_idx + c] = text
        col_idx += colspan
        
    start_row = 1
    if len(rows) > 1:
        second_row_cells = rows[1].find_all(['th', 'td'])
        subheader_terms = ['NO.', 'DATE', 'TIME', 'TYPE', 'UNTS', '8W']
        has_subheaders = any(cell.name == 'th' or cell.text.strip().upper() in subheader_terms for cell in second_row_cells)
        
        if has_subheaders:
            sub_col_idx = 0
            for cell in second_row_cells:
                text = cell.text.strip().replace('\n', ' ').replace('\r', '')
                while sub_col_idx < C and grid[1][sub_col_idx] is not None:
                    sub_col_idx += 1
                if sub_col_idx < C:
                    grid[1][sub_col_idx] = text
                    sub_col_idx += 1
            start_row = 2
            
    headers = []
    for c in range(C):
        h0 = grid[0][c] or f"Column_{c}"
        h1 = grid[1][c] or ""
        h0_clean = " ".join(h0.split())
        h1_clean = " ".join(h1.split())
        if h0_clean == h1_clean or not h1_clean:
            headers.append(h0_clean)
        else:
            headers.append(f"{h0_clean} {h1_clean}")
            
    # Ensure unique headers
    unique_headers = []
    for i, h in enumerate(headers):
        if not h:
            h = f"Column_{i}"
        while h in unique_headers:
            h = h + "_dup"
        unique_headers.append(h)
        
    data = []
    for r in rows[start_row:]:
        cells = r.find_all(['td', 'th'])
        row_data = [cell.text.strip().replace('\xa0', ' ') for cell in cells]
        if len(row_data) == len(unique_headers):
            data.append(row_data)
        elif len(row_data) > 0:
            if len(row_data) < len(unique_headers):
                row_data += [''] * (len(unique_headers) - len(row_data))
            else:
                row_data = row_data[:len(unique_headers)]
            data.append(row_data)
            
    df = pd.DataFrame(data, columns=unique_headers)
    
    # Clean columns: split and rejoin by whitespace to completely sanitize newlines and duplicate spaces
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    
    # Apply dynamic column remapping based on User Request to extract ONLY the requested columns
    mapping = {}
    for col in df.columns:
        col_clean = " ".join(str(col).upper().replace('_', ' ').replace('-', ' ').split())
        
        # 1. Serial Number
        if any(term in col_clean for term in ['S.NO', 'SR. NO', 'SR.NO', 'S. NO', 'SERIAL NUMBER']):
            mapping[col] = 'Serial Number'
            
        # 2. Division Code
        elif any(term in col_clean for term in ['DVSN', 'DIVISION']):
            mapping[col] = 'Division Code'
            
        # 3. Station From
        elif any(term in col_clean for term in ['STTN FROM', 'STTNFRM', 'STTNFROM', 'STTN FRM', 'STATION FROM', 'STN FROM', 'STNFROM', 'STATIONFROM', 'DISPATCH']) or col_clean in ['SOURCE', 'SRCE']:
            mapping[col] = 'Station From'
            
        # 4. Sender
        elif any(term in col_clean for term in ['CNSR', 'CONSIGNOR', 'SENDER', 'CONSIGNER']):
            mapping[col] = 'Sender'
            
        # 5. Receiver
        elif any(term in col_clean for term in ['CNSG', 'CONSIGNEE', 'RECEIVER', 'RECVR']):
            mapping[col] = 'Receiver'
            
        # 7. Commodity Type (RAKE CMDT e.g. GYP) - check this first before Commodity!
        elif any(term in col_clean for term in ['RAKE CMDT', 'RAKE_CMDT', 'RAKECMDT', 'RAKE COMMODITY']):
            mapping[col] = 'Commodity Type'
            
        # 6. Commodity
        elif any(term in col_clean for term in ['CMDT', 'COMMODITY']):
            mapping[col] = 'Commodity'
            
        # 8. Destination
        elif any(term in col_clean for term in ['DSTN', 'DEST', 'DESTINATION', 'STTN TO', 'STN TO']) or col_clean in ['TO']:
            mapping[col] = 'Destination'

        # 9. Indent No.
        elif any(term in col_clean for term in ['INDENT NO', 'INDENT NUMBER', 'INDT NO', 'INDT_NO', 'INDENTNO']):
            mapping[col] = 'Indent No.'

        # 10. Indent Date
        elif any(term in col_clean for term in ['INDENT DATE', 'INDENT DT', 'INDT DATE', 'INDT DT', 'INDT_DT', 'INDENT_DT', 'INDENTDATE']):
            mapping[col] = 'Indent Date'

        # 11. Matured Date
        elif any(term in col_clean for term in ['MATURED DATE', 'MATURED DT', 'MATUR DATE', 'MATUR DT', 'MATURITY DATE', 'MATURITY DT', 'MATURED_DATE', 'MATURED_DT', 'MATUREDDATE']):
            mapping[col] = 'Matured Date'

        # 12. Priority
        elif any(term in col_clean for term in ['PRIORITY', 'PRTY', 'PRIOR']):
            mapping[col] = 'Priority'

        # 13. Rakes
        elif any(term in col_clean for term in ['RAKES', 'NO OF RAKES', 'NO. OF RAKES', 'NO_OF_RAKES']) or col_clean == 'RAKE':
            mapping[col] = 'Rakes'
            
    df = df.rename(columns=mapping)
    
    # Clean up duplicate columns if any (keep first occurrence) to prevent pandas grouping failures
    df = df.loc[:, ~df.columns.duplicated()]
    
    # Filter columns to only keep the requested columns in exact order
    if qry == "MATURED_INDENTS":
        desired_cols = ['Serial Number', 'Station From', 'Indent No.', 'Indent Date', 'Matured Date', 'Priority', 'Sender', 'Receiver', 'Commodity', 'Commodity Type', 'Destination', 'Rakes']
    else:
        desired_cols = ['Serial Number', 'Station From', 'Sender', 'Receiver', 'Commodity', 'Commodity Type', 'Destination']
        
    for col in desired_cols:
        if col not in df.columns:
            df[col] = ''
    df = df[desired_cols]
    
    # Clean up DataFrame: filter out empty rows or row-aggregates containing "Total"
    def is_total_row(row):
        return any("total" in str(val).lower() or "grand" in str(val).lower() for val in row)
    
    df = df[~df.apply(is_total_row, axis=1)]
    
    return df

def save_to_excel(df, sheet_name, filename="FOIS_MIS_Report.xlsx"):
    # If Excel file already exists, open it, else create new
    if os.path.exists(filename):
        try:
            wb = load_workbook(filename)
        except Exception:
            wb = Workbook()
    else:
        wb = Workbook()
        
    # Ensure sheet exists and reset it completely to avoid openpyxl residues/warnings
    if sheet_name in wb.sheetnames:
        temp_ws = None
        if len(wb.sheetnames) == 1:
            temp_ws = wb.create_sheet(title="TempSheetForReset")
        wb.remove(wb[sheet_name])
        ws = wb.create_sheet(title=sheet_name)
        if temp_ws:
            wb.remove(temp_ws)
    else:
        ws = wb.create_sheet(title=sheet_name)
        
    # Delete default Sheet if we have created other sheets
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])
        
    # Enable Gridlines
    ws.views.sheetView[0].showGridLines = True
    
    last_col_letter = get_column_letter(len(df.columns))
    
    # 1. Add corporate Title Banner (Merged A1 to last column)
    ws.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws["A1"]
    title_cell.value = f"FOIS {sheet_name.upper()} MIS REPORT"
    title_cell.font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Sapphire Blue
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 45
    
    # 2. Add Info Block (Merged A2 to last column)
    ws.merge_cells(f"A2:{last_col_letter}2")
    meta_cell = ws["A2"]
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta_cell.value = f"Report Generated: {now_str} | Data Source: Indian Railways FOIS Web Portal"
    meta_cell.font = Font(name="Segoe UI", size=9, italic=True, color="475569")
    meta_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Light grey-blue
    meta_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22
    
    # Row 3 is a spacer
    ws.row_dimensions[3].height = 10
    
    # Row 4 contains the table headers
    headers = list(df.columns)
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx)
        cell.value = str(header).upper()
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid") # Dark Charcoal
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style='thin', color='475569'),
            right=Side(style='thin', color='475569'),
            top=Side(style='thin', color='0F172A'),
            bottom=Side(style='medium', color='1E293B')
        )
    ws.row_dimensions[4].height = 28
    
    # Alignments & borders for data cells
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # slate-50
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    for row_idx, row_data in enumerate(df.values, 5):
        ws.row_dimensions[row_idx].height = 20
        fill_to_use = zebra_fill if row_idx % 2 == 0 else white_fill
        
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = Font(name="Segoe UI", size=10)
            cell.fill = fill_to_use
            cell.border = thin_border
            
            val_str = str(val).strip()
            # Center Serial Numbers, Dates, and Priority codes
            header_name = str(df.columns[col_idx - 1]).upper()
            is_serial = "SERIAL" in header_name or "SR" in header_name or "S.NO" in header_name
            is_date = "DATE" in header_name or "DT" in header_name
            is_priority = "PRIORITY" in header_name or "PRTY" in header_name or "PRIOR" in header_name
            
            try:
                if val_str.isdigit():
                    cell.value = int(val_str)
                    if is_serial:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0'
                else:
                    float_val = float(val_str)
                    cell.value = float_val
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0.00'
            except ValueError:
                if is_date or is_priority:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif len(val_str) <= 6 and (val_str.isupper() or val_str.isalnum()):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
    # Freeze panes below headers
    ws.freeze_panes = "A5"
    
    # 3. Add Excel Auto-Filters on Headers (Row 4)
    last_col_letter = get_column_letter(len(headers))
    last_row = len(df) + 4
    if last_row > 4:
        ws.auto_filter.ref = f"A4:{last_col_letter}{last_row}"
        
    # Autofit Column Widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col[3:]:  # Exclude banner titles
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        header_val = ws.cell(row=4, column=col[0].column).value
        if header_val:
            max_len = max(max_len, len(str(header_val)))
            
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    try:
        wb.save(filename)
        return True, None
    except PermissionError:
        return False, f"Output file '{filename}' is currently open in Excel and locked. Please close the file and click Submit again to update it!"
    except Exception as e:
        return False, str(e)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/get_captcha")
def get_captcha():
    # Fetch Captcha image from FOIS using the current active session
    sess = get_active_session()
    captcha_url = f"https://www.fois.indianrail.gov.in/FOISWebPortal/portalCaptcha.jpg?rand={int(time.time() * 1000)}"
    try:
        r = sess.get(captcha_url, timeout=10)
        if r.status_code == 200:
            # Save the captcha image temporarily inside the static or workspace dir
            captcha_path = os.path.join(os.path.dirname(__file__), "static")
            if not os.path.exists(captcha_path):
                os.makedirs(captcha_path)
            
            img_file = os.path.join(captcha_path, "current_captcha.jpg")
            with open(img_file, "wb") as f:
                f.write(r.content)
            return send_file(img_file, mimetype="image/jpeg", as_attachment=False)
        else:
            return jsonify({"success": False, "message": f"HTTP status {r.status_code}"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": f"Connection error: {e}"}), 500

@app.route("/submit", methods=["POST"])
def submit_query():
    data = request.json or {}
    qry = data.get("Qry", "ODR_RK_OTSG")
    zone = data.get("Zone", "CR")
    optn = data.get("Optn", "7")
    captcha_text = data.get("captchaText", "").strip()
    
    if not captcha_text:
        return jsonify({"success": False, "message": "Captcha input cannot be empty."})
        
    sess = get_active_session()
    submit_url = "https://www.fois.indianrail.gov.in/FOISWebPortal/FWP_RakeOtsg"
    
    payload = {
        "Qry": qry,
        "Zone": zone,
        "captchaText": captcha_text
    }
    
    if qry == "MATURED_INDENTS":
        payload["Optn"] = optn
        
    # Determine Sheet Name & Excel path
    sheet_name = "Outstanding ODR data" if qry == "ODR_RK_OTSG" else "Matured Indents"
    excel_file = "FOIS_MIS_Report.xlsx"
    
    # Fallback monitoring variables
    used_cache = False
    cache_time = None
    df = None
    
    try:
        print(f"Submitting query to FOIS: {qry} | Zone: {zone} | Captcha: {captcha_text}")
        
        # Enforce a 30-second timeout for the railway request to give the portal a solid chance before falling back
        r = sess.post(submit_url, data=payload, timeout=30)
        
        if r.status_code != 200:
            raise requests.exceptions.HTTPError(f"FOIS portal returned HTTP status {r.status_code}")
            
        # Parse the output
        html_response = r.text
        
        # Check if captcha failed
        if "Invalid Captcha" in html_response or "enter captcha" in html_response.lower():
            return jsonify({"success": False, "message": "Incorrect Captcha text. Please try again."})
            
        df = parse_fois_table(html_response, qry)
        
        if df is None or len(df) == 0:
            # Let's check if there's a real error message on the page
            soup = BeautifulSoup(html_response, 'html.parser')
            errmsg_div = soup.find(id="errmsg") or soup.find(class_="alert-danger")
            msg = errmsg_div.text.strip() if errmsg_div else "No records found for the selected Zone / Query."
            
            # If it's a real "no records" return, we shouldn't show cache, we should just tell the user!
            if "no record" in msg.lower() or "not found" in msg.lower() or "does not exist" in msg.lower():
                return jsonify({"success": False, "message": msg})
            else:
                raise ValueError("Empty or malformed table response from railway portal.")
                
    except Exception as e:
        print(f"Railway portal issue or timeout (>18s): {e}. Falling back to last successfully saved Excel data...")
        
        # Load last successfully updated data from the local Excel file
        df, cache_time = load_last_saved_data(sheet_name)
        
        if df is not None:
            used_cache = True
            print(f"Successfully loaded cached data from sheet '{sheet_name}' (saved on {cache_time})")
        else:
            err_msg = str(e)
            print(f"Cache fallback failed: {err_msg}")
            if "Timeout" in err_msg or "time" in err_msg.lower():
                user_msg = "The official Indian Railways FOIS portal took too long to respond (> 20s). No locally saved report was found in your folder yet. Please check your network connection or try again."
            else:
                user_msg = f"Unable to connect to the FOIS portal: {err_msg}. No cached data was found in your folder yet."
            return jsonify({"success": False, "message": user_msg})
            
    # Drop Division Code column if present (e.g. from cached loads)
    if df is not None and 'Division Code' in df.columns:
        df = df.drop(columns=['Division Code'])
        
    # Load Siding and Commodity Lookup Mappings
    s_map, c_map, mapping_warning = get_lookup_mappings()
    s_map_reverse = {str(name).strip().upper(): code for code, name in s_map.items()}
    
    # 1. Compute key MIS metrics for dashboard visualization
    total_val = len(df)
    if qry == "MATURED_INDENTS" and 'Rakes' in df.columns:
        try:
            total_rakes = pd.to_numeric(df['Rakes'], errors='coerce').fillna(0).astype(int).sum()
            total_val = f"{len(df)} ({total_rakes} Rakes)"
        except Exception:
            pass
            
    kpis = {
        "total_count": total_val,
        "columns": list(df.columns)
    }
    
    col_map = {c.upper(): c for c in df.columns}
    sttn_col = col_map.get("STATION FROM") or col_map.get("DISPATCH") or col_map.get("STATION") or col_map.get("STTN") or col_map.get("SOURCE")
    cmd_col = col_map.get("COMMODITY") or col_map.get("CMDT") or col_map.get("CMDT_CODE")
    snd_col = col_map.get("SENDER") or col_map.get("CNSR")
    rcvr_col = col_map.get("RECEIVER") or col_map.get("CNSG")
    
    if qry == "MATURED_INDENTS" and 'Rakes' in df.columns:
        # We sum Rakes for KPIs and Charts
        try:
            df_temp = df.copy()
            df_temp['Rakes_num'] = pd.to_numeric(df_temp['Rakes'], errors='coerce').fillna(1).astype(int)
            
            # Top Station From
            if sttn_col and not df_temp.empty:
                grouped_sttn = df_temp.groupby(sttn_col)['Rakes_num'].sum()
                top_sttn = grouped_sttn.idxmax()
                max_sttn_rakes = grouped_sttn.max()
                top_sttn_name = s_map.get(str(top_sttn).upper(), top_sttn)
                kpis["top_station"] = f"{top_sttn_name} ({max_sttn_rakes})"
            else:
                kpis["top_station"] = "N/A"
                
            # Top Commodity & Commodity Chart
            if cmd_col and not df_temp.empty:
                grouped_cmd = df_temp.groupby(cmd_col)['Rakes_num'].sum()
                top_cmd = grouped_cmd.idxmax()
                max_cmd_rakes = grouped_cmd.max()
                top_cmd_name = c_map.get(str(top_cmd).upper(), top_cmd)
                kpis["top_commodity"] = f"{top_cmd_name} ({max_cmd_rakes})"
                
                cmd_counts = grouped_cmd.sort_values(ascending=False).head(10).to_dict()
                chart_labels = [c_map.get(str(k).upper(), k) for k in cmd_counts.keys()]
                kpis["commodity_chart"] = {"labels": chart_labels, "values": [int(v) for v in cmd_counts.values()]}
            else:
                kpis["top_commodity"] = "N/A"
                kpis["commodity_chart"] = {"labels": [], "values": []}
                
            # UTCL Count (sum Rakes)
            utcl_count = 0
            if not df_temp.empty:
                for _, row in df_temp.iterrows():
                    match_snd = False
                    match_rcv = False
                    if snd_col and row[snd_col]:
                        match_snd = is_utcl_value(row[snd_col], s_map_reverse)
                    if rcvr_col and row[rcvr_col]:
                        match_rcv = is_utcl_value(row[rcvr_col], s_map_reverse)
                    if match_snd or match_rcv:
                        utcl_count += int(row['Rakes_num'])
            kpis["utcl_count"] = f"{utcl_count} Rakes"
            
            # Receiver Chart
            if rcvr_col and not df_temp.empty:
                grouped_rcvr = df_temp.groupby(rcvr_col)['Rakes_num'].sum()
                rcvr_counts = grouped_rcvr.sort_values(ascending=False).head(8).to_dict()
                chart_labels = [s_map.get(str(k).upper(), k) for k in rcvr_counts.keys()]
                kpis["receiver_chart"] = {"labels": chart_labels, "values": [int(v) for v in rcvr_counts.values()]}
            else:
                kpis["receiver_chart"] = {"labels": [], "values": []}
                
        except Exception as kpi_err:
            print(f"Error compiling matured indents KPIs: {kpi_err}")
            # Fallback to counts by overriding qry value for the rest of KPI builder
            qry = "ODR_RK_OTSG"
            
    # Standard ODR count-based KPI logic
    if qry != "MATURED_INDENTS" or 'Rakes' not in df.columns:
        # Top Station
        if sttn_col and not df.empty:
            top_sttn = df[sttn_col].value_counts().idxmax()
            top_sttn_name = s_map.get(str(top_sttn).upper(), top_sttn)
            kpis["top_station"] = f"{top_sttn_name} ({df[sttn_col].value_counts().max()})"
        else:
            kpis["top_station"] = "N/A"
            
        # Top Commodity
        if cmd_col and not df.empty:
            top_cmd = df[cmd_col].value_counts().idxmax()
            top_cmd_name = c_map.get(str(top_cmd).upper(), top_cmd)
            kpis["top_commodity"] = f"{top_cmd_name} ({df[cmd_col].value_counts().max()})"
            
            cmd_counts = df[cmd_col].value_counts().head(10).to_dict()
            chart_labels = [c_map.get(str(k).upper(), k) for k in cmd_counts.keys()]
            kpis["commodity_chart"] = {"labels": chart_labels, "values": [int(v) for v in cmd_counts.values()]}
        else:
            kpis["top_commodity"] = "N/A"
            kpis["commodity_chart"] = {"labels": [], "values": []}
            
        # UTCL Outstanding Count
        utcl_count = 0
        if not df.empty:
            for _, row in df.iterrows():
                match_snd = False
                match_rcv = False
                if snd_col and row[snd_col]:
                    match_snd = is_utcl_value(row[snd_col], s_map_reverse)
                if rcvr_col and row[rcvr_col]:
                    match_rcv = is_utcl_value(row[rcvr_col], s_map_reverse)
                if match_snd or match_rcv:
                    utcl_count += 1
        kpis["utcl_count"] = utcl_count
            
        # Receiver Share Breakdown list for Chart
        if rcvr_col and not df.empty:
            rcvr_counts = df[rcvr_col].value_counts().head(8).to_dict()
            chart_labels = [s_map.get(str(k).upper(), k) for k in rcvr_counts.keys()]
            kpis["receiver_chart"] = {"labels": chart_labels, "values": [int(v) for v in rcvr_counts.values()]}
        else:
            kpis["receiver_chart"] = {"labels": [], "values": []}
        
    # 2. Enrich the DataFrame for Excel output and display (if it was live data)
    # If loaded from Excel cache, columns are already enriched!
    excel_warning = None
    if not used_cache:
        try:
            df_enriched = df.copy()
            # Enrich siding-related columns
            siding_cols = ['Station From', 'Destination', 'Sender', 'Receiver']
            for col in siding_cols:
                if col in df_enriched.columns:
                    df_enriched[col] = df_enriched[col].apply(lambda x: enrich_siding(x, s_map))
            # Enrich commodity-related columns
            commodity_cols = ['Commodity', 'Commodity Type']
            for col in commodity_cols:
                if col in df_enriched.columns:
                    df_enriched[col] = df_enriched[col].apply(lambda x: enrich_commodity(x, c_map))
            
            # 3. Generate the Excel MIS Report sheet using enriched data
            excel_ok, excel_err = save_to_excel(df_enriched, sheet_name, excel_file)
            if not excel_ok:
                excel_warning = excel_err
        except Exception as ee:
            print(f"Error enriching DataFrame: {ee}")
            df_enriched = df
            excel_warning = f"Enrichment error: {ee}"
    else:
        df_enriched = df
        
    # 4. Prepare records for frontend display
    records = df_enriched.to_dict(orient="records")
    
    # Compile a clear and concise warning if either mapping or saving failed
    active_warning = mapping_warning or excel_warning
    
    if used_cache:
        msg = f"The Indian Railways FOIS portal took more than 20s to respond. Showing the last successfully updated data from your local folder (saved on {cache_time}) so you can work without delay!"
    else:
        if active_warning:
            msg = f"Warning: {active_warning}"
        else:
            msg = f"Successfully parsed {len(df)} records and updated sheet '{sheet_name}' in Excel!"
        
    return jsonify({
        "success": True,
        "cached": used_cache,
        "cache_time": cache_time,
        "message": msg,
        "kpis": kpis,
        "records": records,
        "excel_path": os.path.abspath(excel_file)
    })

@app.route("/open_excel", methods=["POST"])
def open_excel():
    excel_path = os.path.abspath("FOIS_MIS_Report.xlsx")
    if os.path.exists(excel_path):
        try:
            # os.startfile is Windows specific, perfect for our user!
            os.startfile(excel_path)
            return jsonify({"success": True, "message": "Excel opened successfully."})
        except Exception as e:
            return jsonify({"success": False, "message": f"Failed to open Excel: {e}"})
    else:
        return jsonify({"success": False, "message": "Excel file does not exist yet. Run a query first!"})

if __name__ == "__main__":
    # Create static directory if missing
    static_path = os.path.join(os.path.dirname(__file__), "static")
    if not os.path.exists(static_path):
        os.makedirs(static_path)
        
    print("-------------------------------------------------------")
    print("     Starting FOIS Tracker Server on Port 5000         ")
    print("-------------------------------------------------------")
    app.run(debug=True, host="0.0.0.0", port=5000)
