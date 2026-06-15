import io
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import pdfplumber
import csv
import zipfile
from xml.etree import ElementTree

# ============================================================================
# CONFIGURATION
# ============================================================================

MONTHS_EN = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

MONTHS_FR = [
    "janvier", "février", "fevrier", "mars", "avril", "mai", "juin",
    "juillet", "août", "aout", "septembre", "octobre", "novembre", "décembre", "decembre"
]

SHEET_PATTERNS = {
    "Budget Initial": [
        "budget initial",
        "budget"
    ],
    "Fiche Stationnement": [
        "fiche stationnement",
        "fiche de stationnement",
        "stationnement",
        "1. fiche"
    ],
    "Donnees Historiques": [
        "donnees historiques",
        "données historiques",
        "historiques",
        "historique",
        "2. donnees",
        "2. données"
    ],
}

DONNEES_HISTORIQUES_LABELS = {
    "Transient Revenue": [
        "revenus horaires",
        "transient",
        "transitoire",
        "stationnement horaire"
    ],
    "Monthly Revenues": [
        "revenus mensuels",
        "monthly",
        "mensuels",
        "stationnement mensuel"
    ],
    "Car-Wash Revenue": [
        "revenus lave-auto",
        "lave-auto",
        "car-wash",
        "car wash",
        "lavage"
    ],
    "Hotel Revenue": [
        "revenus hôtel",
        "revenus hotel",
        "hôtel",
        "hotel",
        "hotelier"
    ],
    "Interests": [
        "revenus d'intérêts",
        "revenus d'interets",
        "intérêts",
        "interets",
        "interests",
        "interest"
    ],
    "Miscellaneous": [
        "autres revenus",
        "miscellaneous",
        "misc",
        "divers"
    ],
    "Parking Revenue": [
        "total revenus bruts",
        "revenus bruts",
        "parking revenue",
        "total revenus",
        "revenus stationnement"
    ],
    "Discount-Gratuities - Transient": [
        "gratuités",
        "gratuites",
        "gratuities",
        "discount",
        "escomptes",
        "courtoisies"
    ],
    "Discount-Gratuities - Monthly": [
        "rabais",
        "discount monthly",
        "escomptes mensuels",
        "rabais mensuel"
    ],
    "TOTAL REVENUE": [
        "total revenus",
        "total revenue",
        "revenu total",
        "total des revenus",
        "total revenu"
    ],
    "Parking wages": [
        "salaire stationnement",
        "parking wages",
        "salaires",
        "main d'oeuvre",
        "main d'œuvre",
        "masse salariale"
    ],
    "Total Operation expenses": [
        "total operation expenses",
        "total dépenses",
        "total depenses",
        "operation expenses",
        "dépenses d'opération",
        "depenses d'operation",
        "total des dépenses"
    ],
    "OPERATION SURPLUS": [
        "operation surplus",
        "surplus d'opération",
        "surplus d'operation",
        "bénéfice",
        "benefice",
        "excédent",
        "excedent"
    ],
    "Percent Management fee": [
        "percent management fee",
        "frais de gestion",
        "% management",
        "management fee",
        "honoraires de gestion"
    ],
}

FICHE_STATIONNEMENT_MAP = [
    ("Parking Revenue", "K17"),
    ("Monthly Revenues", "K18"),
    ("Transient Revenue", "K19"),
    ("TOTAL REVENUE", "K20"),
    ("Total Operation expenses", "K21"),
    ("Parking wages", "K22"),
    ("OPERATION SURPLUS", "K23"),
    ("Percent Management fee", "K24"),
]

# ============================================================================
# FILE TYPE HANDLERS - Accept ANY format
# ============================================================================

def is_excel_file(file_bytes_or_obj):
    """Check if file is Excel format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        return name.endswith('.xlsx') or name.endswith('.xls') or name.endswith('.xlsm')
    return False


def is_csv_file(file_bytes_or_obj):
    """Check if file is CSV format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        return name.endswith('.csv') or name.endswith('.tsv')
    return False


def is_pdf_file(file_bytes_or_obj):
    """Check if file is PDF format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        return name.endswith('.pdf')
    return False


def is_docx_file(file_bytes_or_obj):
    """Check if file is DOCX format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        return name.endswith('.docx')
    return False


def is_txt_file(file_bytes_or_obj):
    """Check if file is text format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        return name.endswith(('.txt', '.tsv', '.text'))
    return False


def get_file_bytes(uploaded_file):
    """Get bytes from uploaded file, handling various types"""
    if hasattr(uploaded_file, 'read'):
        uploaded_file.seek(0)
        return uploaded_file.read()
    if hasattr(uploaded_file, 'getvalue'):
        return uploaded_file.getvalue()
    return uploaded_file


def read_excel_to_dataframe(uploaded_file):
    """Read Excel file (any format) to pandas DataFrame dict"""
    try:
        file_bytes = get_file_bytes(uploaded_file)
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        sheets = {}
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                sheets[sheet_name] = df
            except Exception:
                pass
        return sheets
    except Exception:
        return None


def read_csv_to_dataframe(uploaded_file):
    """Read CSV file to pandas DataFrame"""
    try:
        file_bytes = get_file_bytes(uploaded_file)
        text = file_bytes.decode('utf-8', errors='ignore')
        df = pd.read_csv(io.StringIO(text))
        return {"Sheet1": df}
    except Exception:
        try:
            file_bytes = get_file_bytes(uploaded_file)
            text = file_bytes.decode('latin-1', errors='ignore')
            df = pd.read_csv(io.StringIO(text))
            return {"Sheet1": df}
        except Exception:
            return None


def read_pdf_to_dataframe(uploaded_file):
    """Extract tables from PDF and convert to DataFrames"""
    try:
        file_bytes = get_file_bytes(uploaded_file)
        sheets = {}
        
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table_num, table in enumerate(tables):
                    if table and len(table) > 1:
                        headers = table[0] if table[0] else [f"Col{i}" for i in range(len(table[1]))]
                        data = table[1:] if table[0] else table
                        
                        clean_headers = []
                        for h in headers:
                            if h is None:
                                clean_headers.append("")
                            else:
                                clean_headers.append(str(h).strip())
                        
                        df = pd.DataFrame(data, columns=clean_headers)
                        sheet_key = f"Page{page_num+1}_Table{table_num+1}"
                        sheets[sheet_key] = df
        
        return sheets if sheets else None
    except Exception:
        return None


def read_docx_to_dataframe(uploaded_file):
    """Extract tables from Word document"""
    try:
        file_bytes = get_file_bytes(uploaded_file)
        sheets = {}
        
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            if 'word/document.xml' in z.namelist():
                xml_content = z.read('word/document.xml')
                tree = ElementTree.fromstring(xml_content)
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                
                tables = tree.findall('.//w:tbl', ns)
                for table_num, table in enumerate(tables):
                    rows = table.findall('.//w:tr', ns)
                    table_data = []
                    for row in rows:
                        cells = row.findall('.//w:tc', ns)
                        row_data = []
                        for cell in cells:
                            texts = cell.findall('.//w:t', ns)
                            cell_text = ''.join(t.text for t in texts if t.text)
                            row_data.append(cell_text.strip())
                        if row_data:
                            table_data.append(row_data)
                    
                    if table_data and len(table_data) >= 2:
                        headers = table_data[0]
                        data = table_data[1:]
                        df = pd.DataFrame(data, columns=headers)
                        sheets[f"Table_{table_num+1}"] = df
        
        return sheets if sheets else None
    except Exception:
        return None


def read_txt_to_dataframe(uploaded_file):
    """Try to parse text file as tabular data"""
    try:
        file_bytes = get_file_bytes(uploaded_file)
        text = file_bytes.decode('utf-8', errors='ignore')
        lines = text.strip().split('\n')
        
        if not lines:
            return None
        
        # Try tab-separated
        if '\t' in lines[0]:
            reader = csv.reader(io.StringIO(text), delimiter='\t')
            data = list(reader)
            if data and len(data) >= 2:
                headers = data[0]
                df = pd.DataFrame(data[1:], columns=headers)
                return {"Sheet1": df}
        
        # Try comma-separated
        if ',' in lines[0]:
            reader = csv.reader(io.StringIO(text))
            data = list(reader)
            if data and len(data) >= 2:
                headers = data[0]
                df = pd.DataFrame(data[1:], columns=headers)
                return {"Sheet1": df}
        
        return None
    except Exception:
        return None


def read_any_file_to_dataframes(uploaded_file):
    """
    Universal file reader - handles Excel, PDF, CSV, DOCX, TXT, and more.
    Returns (dict of {sheet_name: DataFrame}, file_type) or (None, None) if unreadable.
    """
    if uploaded_file is None:
        return None, None
    
    # Try Excel first
    if is_excel_file(uploaded_file):
        result = read_excel_to_dataframe(uploaded_file)
        if result:
            return result, "excel"
    
    # Try CSV
    if is_csv_file(uploaded_file):
        result = read_csv_to_dataframe(uploaded_file)
        if result:
            return result, "csv"
    
    # Try PDF
    if is_pdf_file(uploaded_file):
        result = read_pdf_to_dataframe(uploaded_file)
        if result:
            return result, "pdf"
    
    # Try DOCX
    if is_docx_file(uploaded_file):
        result = read_docx_to_dataframe(uploaded_file)
        if result:
            return result, "docx"
    
    # Try TXT
    if is_txt_file(uploaded_file):
        result = read_txt_to_dataframe(uploaded_file)
        if result:
            return result, "txt"
    
    # Last resort: try Excel anyway (some files don't have proper extension)
    try:
        result = read_excel_to_dataframe(uploaded_file)
        if result:
            return result, "excel"
    except Exception:
        pass
    
    # Try CSV anyway
    try:
        result = read_csv_to_dataframe(uploaded_file)
        if result:
            return result, "csv"
    except Exception:
        pass
    
    return None, None


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    """Extract parking code (like CMO142) from filename"""
    if not filename:
        return None
    name = filename.rsplit('.', 1)[0]
    match = re.search(r'(CMO\d+)', name, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    if 'LUNA' in name.upper():
        return 'LUNA'
    parts = name.replace('(', ' ').replace(')', ' ').split('_')[0].split()[0]
    return parts.upper()


def find_sheet_by_pattern(wb, patterns):
    """Find a sheet in an openpyxl workbook by matching patterns against sheet names"""
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            pattern_clean = pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
            if pattern_clean in sheet_lower:
                return sheet_name
    return None


def find_sheet_in_dict(sheets_dict, parking_code):
    """
    Find the sheet in a dict of DataFrames that matches the parking code.
    Returns the sheet name or None.
    """
    if not sheets_dict:
        return None
    
    # Exact match first
    for sheet_name in sheets_dict:
        if sheet_name.upper().strip() == parking_code.upper().strip():
            return sheet_name
    
    # Partial match
    for sheet_name in sheets_dict:
        if parking_code.upper() in sheet_name.upper():
            return sheet_name
    
    # If no match, return the sheet with the most rows
    if sheets_dict:
        best_sheet = None
        max_rows = 0
        for name, df in sheets_dict.items():
            if len(df) > max_rows:
                max_rows = len(df)
                best_sheet = name
        return best_sheet
    
    return None


def detect_year_from_filename(filename):
    """Try to extract year from filename like 'P&L_2025.xlsx' or '2025_P&L.pdf'"""
    if not filename:
        return None
    years = re.findall(r'(20\d{2})', filename)
    if years:
        return int(years[0])
    return None


def detect_year_from_data(df):
    """Try to detect year from DataFrame content by searching for year references"""
    if df is None:
        return None
    
    for col_idx in range(min(20, len(df.columns))):
        for row_idx in range(min(5, len(df))):
            try:
                cell = str(df.iloc[row_idx, col_idx])
                years = re.findall(r'(20\d{2})', cell)
                if years:
                    from collections import Counter
                    year_counts = Counter(years)
                    most_common_year = year_counts.most_common(1)[0][0]
                    return int(most_common_year)
            except Exception:
                continue
    
    return None


def safe_float(value, default=0.0):
    """Safely convert a value to float, handling strings, None, and special characters"""
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '').replace('€', '')
            value = value.replace('(', '').replace(')', '')
            if value.startswith('-'):
                value = value[1:]
                return -safe_float(value, default)
        return float(value)
    except (ValueError, TypeError):
        return default


def get_cell_display_value(ws, row, col):
    """
    Get the cell value from an openpyxl worksheet.
    If it's a formula, try to get the computed value.
    For data_only workbooks, this returns the cached value.
    """
    try:
        cell = ws.cell(row=row, column=col)
        val = cell.value
        
        if val is None or (isinstance(val, str) and not val.strip()):
            return ""
        
        # If it's a formula (starts with =), return empty as we can't compute it
        if isinstance(val, str) and val.startswith('='):
            return ""
        
        return str(val).strip()
    except Exception:
        return ""


def find_template_rows_by_label(ws, french_labels, max_rows=100):
    """
    Search Column A of a worksheet for French labels.
    Returns list of matching row numbers.
    """
    matching_rows = []
    
    for row_idx in range(1, max_rows + 1):
        cell_value = get_cell_display_value(ws, row_idx, 1)
        
        if not cell_value:
            continue
        
        cell_lower = cell_value.lower()
        # Clean: remove special characters, keep letters/numbers/spaces
        cell_clean = re.sub(r'[^a-zéèêëàâîïôûùç0-9 ]', ' ', cell_lower)
        cell_clean = re.sub(r'\s+', ' ', cell_clean).strip()
        
        for french_label in french_labels:
            label_lower = french_label.lower()
            label_clean = re.sub(r'[^a-zéèêëàâîïôûùç0-9 ]', ' ', label_lower)
            label_clean = re.sub(r'\s+', ' ', label_clean).strip()
            
            # Check if the cell contains the label (partial match either way)
            if label_clean in cell_clean or cell_clean in label_clean:
                matching_rows.append(row_idx)
                break
    
    return matching_rows


def read_year_mapping_from_template(ws):
    """
    Read the year mapping from Donnees Historiques template.
    Looks for a row that has years like 2026, 2025, etc. in columns B-M.
    Returns dict: {month_index: year} or fallback if not found.
    """
    for row_idx in range(35, 55):
        year_map = {}
        for col_idx in range(2, 14):  # Columns B through M
            cell_value = get_cell_display_value(ws, row_idx, col_idx)
            year_match = re.search(r'(20\d{2})', cell_value)
            if year_match:
                year_map[col_idx - 2] = int(year_match.group(1))  # 0-based month index
        
        if len(year_map) >= 6:  # If we found at least 6 months with years
            return year_map
    
    # Fallback: current year for Jan-Apr, previous year for May-Dec
    current_year = datetime.now().year
    year_map = {}
    for i in range(4):
        year_map[i] = current_year
    for i in range(4, 12):
        year_map[i] = current_year - 1
    return year_map


def get_template_labels_for_debug(ws, max_rows=100):
    """Debug helper to get all non-empty labels from column A"""
    labels = []
    for row_idx in range(1, max_rows + 1):
        val = get_cell_display_value(ws, row_idx, 1)
        if val and len(val) > 2:
            labels.append(f"R{row_idx}:{val[:80]}")
    return labels


# ============================================================================
# P&L DATA EXTRACTION - Works with ANY file format
# ============================================================================

def extract_pnl_data_from_dataframe(df, sheet_name_hint=None):
    """
    Extract P&L data from a pandas DataFrame.
    Handles different P&L layouts (labels in column A, monthly data in B-M, yearly in N).
    Returns dict with 'monthly' and 'yearly' data.
    """
    result = {
        'monthly': {},
        'yearly': {}
    }
    
    if df is None or len(df) == 0:
        return result
    
    # Try to find header row (where month names are)
    header_row = None
    for row_idx in range(min(20, len(df))):
        for col_idx in range(min(14, len(df.columns))):
            try:
                cell_val = str(df.iloc[row_idx, col_idx]).lower().strip()
                if cell_val in ['january', 'february', 'march', 'janvier', 'février', 'fevrier', 'mars']:
                    header_row = row_idx
                    break
            except Exception:
                continue
        if header_row is not None:
            break
    
    # Start searching from after header row (or from row 1 if no header found)
    data_start = (header_row + 1) if header_row else 1
    
    for row_idx in range(data_start, len(df)):
        try:
            label = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        except Exception:
            continue
        
        if not label:
            continue
        
        label_lower = label.lower()
        if label_lower in ['code', 'profit & loss', '', 'nan', 'none']:
            continue
        
        # Skip date rows like "01-01-2025"
        if re.search(r'\d{2}[-/]\d{2}[-/]\d{2}', label):
            continue
        
        # Check for any data in the row
        has_any_data = False
        for col_idx in range(1, min(14, len(df.columns))):
            if safe_float(df.iloc[row_idx, col_idx]) != 0:
                has_any_data = True
                break
        
        if not has_any_data:
            continue
        
        # Monthly data (columns B-M = indices 1-12)
        monthly = {}
        for month_idx in range(12):
            col_idx = month_idx + 1
            if col_idx < len(df.columns):
                val = safe_float(df.iloc[row_idx, col_idx])
                monthly[MONTHS_EN[month_idx]] = val
        
        # Yearly total (column N = index 13, or last column)
        yearly_total = 0
        if len(df.columns) > 13:
            yearly_total = safe_float(df.iloc[row_idx, 13])
        elif len(df.columns) > 1:
            # Use last column as total
            yearly_total = safe_float(df.iloc[row_idx, -1])
        
        # Store using cleaned label
        clean_label = label.strip().replace('  ', ' ')
        result['monthly'][clean_label] = monthly
        result['yearly'][clean_label] = yearly_total
    
    return result


def extract_pnl_data(uploaded_file, parking_code):
    """
    Extract P&L data from ANY file format (Excel, PDF, CSV, DOCX, TXT).
    Returns (pnl_data_dict, file_type) or (None, None).
    """
    sheets_dict, file_type = read_any_file_to_dataframes(uploaded_file)
    
    if sheets_dict is None:
        return None, None
    
    # Find the right sheet for this parking code
    sheet_name = find_sheet_in_dict(sheets_dict, parking_code)
    
    if sheet_name is None:
        # Try any sheet that might contain the parking code
        for name in sheets_dict:
            if parking_code.upper() in name.upper():
                sheet_name = name
                break
    
    if sheet_name is None:
        # Just use the first/largest sheet as fallback
        if sheets_dict:
            max_len = 0
            for name, df in sheets_dict.items():
                if len(df) > max_len:
                    max_len = len(df)
                    sheet_name = name
        else:
            return None, None
    
    if sheet_name is None:
        return None, None
    
    df = sheets_dict[sheet_name]
    pnl_data = extract_pnl_data_from_dataframe(df, sheet_name)
    
    return pnl_data, file_type


def find_pnl_value(pnl_data, label_alternatives):
    """
    Search through P&L data for a matching label.
    Tries exact match first, then partial match.
    Returns the yearly value or 0.
    """
    if pnl_data is None:
        return 0
    
    yearly = pnl_data.get('yearly', {})
    
    if not yearly:
        return 0
    
    # Try exact match first
    for alt in label_alternatives:
        alt_lower = alt.lower().strip()
        for key in yearly:
            if alt_lower == key.lower().strip():
                return yearly[key]
    
    # Try partial match - label contains key or key contains label
    for alt in label_alternatives:
        alt_lower = alt.lower().strip()
        for key in yearly:
            key_lower = key.lower().strip()
            if alt_lower in key_lower or key_lower in alt_lower:
                return yearly[key]
    
    return 0


def merge_monthly_data(current_year_data, previous_year_data, year_map):
    """
    Merge monthly data from current and previous years based on year_map.
    year_map: {month_index: year} - tells which year each month comes from.
    Returns merged monthly data dict.
    """
    merged = {}
    current_year = datetime.now().year
    
    for month_idx, year in year_map.items():
        month_name = MONTHS_EN[month_idx]
        
        # Determine which data source to use based on the year
        if year == current_year and current_year_data:
            source = current_year_data
        elif year == current_year - 1 and previous_year_data:
            source = previous_year_data
        elif year == current_year - 2 and previous_year_data:
            # If year mapping says 2 years ago but we only have "previous" data, use it
            source = previous_year_data
        elif current_year_data:
            # Ultimate fallback
            source = current_year_data
        else:
            continue
        
        # Get monthly values from the chosen source
        if source and 'monthly' in source:
            for label, monthly in source['monthly'].items():
                if label not in merged:
                    merged[label] = {}
                merged[label][month_name] = monthly.get(month_name, 0)
    
    return merged


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, previous_year_data, parking_code):
    """
    Update Budget Initial sheet.
    Sets cell S8 to the PREVIOUS year TOTAL REVENUE.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["❌ Budget Initial: Sheet not found"]
        
        ws = wb[sheet_name]
        
        total = 0
        if previous_year_data:
            total_alternatives = [
                "TOTAL REVENUE",
                "Total Revenue",
                "total revenus",
                "revenu total",
                "REVENU TOTAL",
                "Total Revenus",
                "TOTAL DES REVENUS",
                "Total des revenus",
                "Revenus Totals"
            ]
            total = find_pnl_value(previous_year_data, total_alternatives)
        
        if total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f} (previous year)")
        else:
            # Try other common labels
            backup_alternatives = [
                "Parking Revenue",
                "Total Revenus Bruts",
                "Revenue",
                "Total",
                "Total Revenue Bruts"
            ]
            found = False
            for alt in backup_alternatives:
                if previous_year_data:
                    total = previous_year_data['yearly'].get(alt, 0)
                    if total > 0:
                        ws["S8"] = total
                        ws["S8"].number_format = '#,##0.00'
                        updates.append(f"✅ Budget Initial: S8 = ${total:,.2f} (from '{alt}')")
                        found = True
                        break
            
            if not found:
                updates.append("⚠️ Budget Initial: No TOTAL REVENUE found in previous year P&L")
                if previous_year_data and 'yearly' in previous_year_data:
                    available_keys = list(previous_year_data['yearly'].keys())[:10]
                    updates.append(f"   Available labels: {available_keys}")
    except Exception as e:
        updates.append(f"❌ Budget Initial: {str(e)}")
    return updates


def update_fiche_stationnement(wb, year_minus_2_data, parking_code, word_data=None):
    """
    Update Fiche Stationnement sheet.
    Uses data from 2 YEARS AGO (year - 2 relative to budget year).
    Fills K17-K26 with yearly values.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["❌ Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        
        if year_minus_2_data is None:
            updates.append("⚠️ Fiche Stationnement: No data available (need P&L from 2 years ago)")
            return updates
        
        for pnl_label, cell in FICHE_STATIONNEMENT_MAP:
            yearly_value = find_pnl_value(year_minus_2_data, [pnl_label])
            
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00'
                updates.append(f"✅ {cell} = {pnl_label}: ${yearly_value:,.2f}")
            else:
                updates.append(f"⚠️ {cell} = {pnl_label}: Not found in P&L")
        
        # K26 = TOTAL REVENUE
        total_revenue = find_pnl_value(
            year_minus_2_data,
            ["TOTAL REVENUE", "Total Revenue", "total revenus", "Total des revenus"]
        )
        ws["K26"] = total_revenue
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"✅ K26 = TOTAL REVENUE: ${total_revenue:,.2f}")
        
        # Apply Word data if provided
        if word_data:
            for row in range(43, 56):
                if "Nb abonnés" in word_data:
                    ws[f"H{row}"] = word_data["Nb abonnés"]
                if "Informations" in word_data:
                    ws[f"I{row}"] = word_data["Informations"]
                if "Avant taxes" in word_data:
                    ws[f"J{row}"] = word_data["Avant taxes"]
                    ws[f"L{row}"] = word_data["Avant taxes"]
            updates.append("✅ Updated Word data in rows 43-55")
    except Exception as e:
        updates.append(f"❌ Fiche Stationnement: {str(e)}")
    return updates


def update_donnees_historiques(wb, merged_monthly_data, parking_code):
    """
    Update Donnees Historiques sheet.
    Fills monthly data from the correct years based on year mapping.
    Jan-Apr from current year, May-Dec from previous year.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["❌ Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # Read year mapping from template
        year_map = read_year_mapping_from_template(ws)
        updates.append(
            f"📅 Year mapping: Jan={year_map.get(0)}, Feb={year_map.get(1)}, "
            f"Mar={year_map.get(2)}, Apr={year_map.get(3)}, "
            f"May={year_map.get(4)}, Jun={year_map.get(5)}, "
            f"Jul={year_map.get(6)}, Aug={year_map.get(7)}, "
            f"Sep={year_map.get(8)}, Oct={year_map.get(9)}, "
            f"Nov={year_map.get(10)}, Dec={year_map.get(11)}"
        )
        
        # Debug: show what labels we found in the template
        debug_labels = get_template_labels_for_debug(ws, max_rows=80)
        if debug_labels:
            updates.append(f"🔍 Template labels found: {'; '.join(debug_labels[:6])}")
        
        if not merged_monthly_data:
            updates.append("⚠️ Donnees Historiques: No merged monthly data available")
            return updates
        
        cells_updated = 0
        rows_filled = []
        
        for pnl_label, french_labels in DONNEES_HISTORIQUES_LABELS.items():
            # Find matching value in merged monthly data
            monthly_values = None
            matched_key = None
            
            # Try exact match first
            if pnl_label in merged_monthly_data:
                monthly_values = merged_monthly_data[pnl_label]
                matched_key = pnl_label
            else:
                # Try partial match
                for key in merged_monthly_data:
                    if pnl_label.lower() in key.lower() or key.lower() in pnl_label.lower():
                        monthly_values = merged_monthly_data[key]
                        matched_key = key
                        break
            
            if monthly_values is None:
                continue
            
            # Check if there's any non-zero data
            if all(v == 0 for v in monthly_values.values()):
                continue
            
            # Find matching row in template
            matching_rows = find_template_rows_by_label(ws, french_labels, max_rows=80)
            
            if not matching_rows:
                continue
            
            template_row = matching_rows[0]
            row_cells = 0
            
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values and monthly_values[month_name] != 0:
                    col_letter = get_column_letter(month_idx + 2)  # B=2, C=3, etc.
                    cell_ref = f"{col_letter}{template_row}"
                    ws[cell_ref] = monthly_values[month_name]
                    ws[cell_ref].number_format = '#,##0.00'
                    cells_updated += 1
                    row_cells += 1
            
            if row_cells > 0:
                rows_filled.append(f"  Row {template_row}: {pnl_label} ({row_cells} months)")
        
        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: {cells_updated} cells updated in {len(rows_filled)} rows")
            for row_info in rows_filled:
                updates.append(row_info)
        else:
            updates.append("⚠️ Donnees Historiques: No matches found between P&L data and template labels")
            updates.append(f"   Template has {len(debug_labels)} labels in column A")
            updates.append(f"   P&L data has {len(merged_monthly_data)} data rows")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {str(e)}")
    return updates


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(
    excel_file,
    pnl_current_year,
    pnl_previous_year=None,
    pnl_two_years_ago=None,
    parking_code=None,
    word_data=None
):
    """
    Main function to process the Excel template with P&L data from MULTIPLE years.
    
    Args:
        excel_file: Uploaded Excel template (must be .xlsx)
        pnl_current_year: P&L file for CURRENT year (e.g., 2026 for 2027 budget)
                          Can be Excel, PDF, CSV, DOCX, or TXT
        pnl_previous_year: P&L file for PREVIOUS year (e.g., 2025 for 2027 budget)
                          Can be any format. Used for Donnees Historiques May-Dec
        pnl_two_years_ago: P&L file from 2 years ago (e.g., 2024 for 2027 budget)
                          Can be any format. Used for Fiche Stationnement
        parking_code: Parking code (extracted from filename if None)
        word_data: Optional dict for Fiche Stationnement H-M rows
    
    Year Usage per Sheet:
        - Budget Initial: Previous year total (year - 1)
        - Fiche Stationnement: 2 years ago data (year - 2)
        - Donnees Historiques: Jan-Apr = current year, May-Dec = previous year
    
    Returns:
        (output_bytesio, updates_list) or (None, error_list)
    """
    updates = []
    
    # ── Extract parking code ────────────────────────────────────────────
    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)
    
    if not parking_code:
        return None, ["❌ Could not determine parking code from filename"]
    
    updates.append(f"🔍 Processing: {parking_code}")
    
    # ── Detect years from uploaded files ────────────────────────────────
    current_year_name = "?"
    previous_year_name = "?"
    two_years_ago_name = "?"
    
    if hasattr(pnl_current_year, 'name'):
        detected = detect_year_from_filename(pnl_current_year.name)
        if detected:
            current_year_name = str(detected)
    
    if pnl_previous_year and hasattr(pnl_previous_year, 'name'):
        detected = detect_year_from_filename(pnl_previous_year.name)
        if detected:
            previous_year_name = str(detected)
    
    if pnl_two_years_ago and hasattr(pnl_two_years_ago, 'name'):
        detected = detect_year_from_filename(pnl_two_years_ago.name)
        if detected:
            two_years_ago_name = str(detected)
    
    updates.append(
        f"📂 Files: Current={current_year_name} | "
        f"Previous={previous_year_name} | "
        f"2YA={two_years_ago_name}"
    )
    
    # ── Extract data from P&L files (any format!) ───────────────────────
    current_year_data, current_file_type = extract_pnl_data(pnl_current_year, parking_code)
    previous_year_data, prev_file_type = (None, None)
    two_years_ago_data, two_ya_file_type = (None, None)
    
    if pnl_previous_year:
        previous_year_data, prev_file_type = extract_pnl_data(pnl_previous_year, parking_code)
    
    if pnl_two_years_ago:
        two_years_ago_data, two_ya_file_type = extract_pnl_data(pnl_two_years_ago, parking_code)
    
    # Check if we got ANY data
    if current_year_data is None and previous_year_data is None and two_years_ago_data is None:
        return None, [
            f"❌ Could not find P&L data for {parking_code} in any uploaded file. "
            f"Please check that the parking code matches the sheet name or data in the P&L file."
        ]
    
    if current_year_data:
        keys = list(current_year_data['yearly'].keys())[:8]
        updates.append(f"📊 Current year keys ({current_file_type}): {keys}")
    
    if previous_year_data:
        keys = list(previous_year_data['yearly'].keys())[:8]
        updates.append(f"📊 Previous year keys ({prev_file_type}): {keys}")
    
    if two_years_ago_data:
        keys = list(two_years_ago_data['yearly'].keys())[:8]
        updates.append(f"📊 2YA keys ({two_ya_file_type}): {keys}")
    
    # ── Read template - need TWO workbooks for formula handling ─────────
    try:
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        file_bytes = excel_file.read()
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        
        # wb_read: data_only=True to get computed values (for label matching)
        # wb_write: data_only=False to preserve formulas (for writing)
        wb_read = load_workbook(io.BytesIO(file_bytes), data_only=True)
        wb_write = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception as e:
        return None, [f"❌ Error reading template: {str(e)}"]
    
    # ── Read year mapping from Donnees Historiques ──────────────────────
    dh_sheet_name = find_sheet_by_pattern(wb_read, SHEET_PATTERNS["Donnees Historiques"])
    year_map = None
    if dh_sheet_name:
        year_map = read_year_mapping_from_template(wb_read[dh_sheet_name])
    
    # ── Merge monthly data based on year mapping ────────────────────────
    merged_monthly = {}
    if year_map and current_year_data and previous_year_data:
        # Use current_year for Jan-Apr, previous_year for May-Dec (as per year_map)
        merged_monthly = merge_monthly_data(current_year_data, previous_year_data, year_map)
    
    if not merged_monthly and current_year_data:
        # Fallback: use current year for all months
        merged_monthly = current_year_data['monthly']
    
    # ── Determine which data to use for each sheet ──────────────────────
    # Budget Initial = previous year (year - 1)
    budget_initial_data = previous_year_data if previous_year_data else current_year_data
    
    # Fiche Stationnement = 2 years ago (year - 2)
    # If 2YA not provided, try previous year, then current year
    fiche_data = two_years_ago_data
    if fiche_data is None:
        fiche_data = previous_year_data
    if fiche_data is None:
        fiche_data = current_year_data
    
    # Donnees Historiques = merged (current + previous)
    dh_data = merged_monthly if merged_monthly else {}
    if not dh_data and current_year_data:
        dh_data = current_year_data['monthly']
    
    # ── Update all sheets ───────────────────────────────────────────────
    updates.extend(update_budget_initial(wb_write, budget_initial_data, parking_code))
    updates.extend(update_fiche_stationnement(wb_write, fiche_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb_write, dh_data, parking_code))
    
    # ── Summary ─────────────────────────────────────────────────────────
    success_count = sum(1 for u in updates if u.startswith("✅"))
    warning_count = sum(1 for u in updates if u.startswith("⚠️"))
    
    if success_count == 0:
        updates.append("💡 No updates were made. Check that P&L files contain matching data for the parking code and labels.")
    
    # ── Reset file pointers ─────────────────────────────────────────────
    if hasattr(pnl_current_year, 'seek'):
        pnl_current_year.seek(0)
    if pnl_previous_year and hasattr(pnl_previous_year, 'seek'):
        pnl_previous_year.seek(0)
    if pnl_two_years_ago and hasattr(pnl_two_years_ago, 'seek'):
        pnl_two_years_ago.seek(0)
    
    # ── Save output ─────────────────────────────────────────────────────
    output = io.BytesIO()
    wb_write.save(output)
    output.seek(0)
    
    return output, updates
