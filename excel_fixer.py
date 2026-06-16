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

MONTH_NAMES_MAP = {
    "january": "January", "february": "February", "march": "March",
    "april": "April", "may": "May", "june": "June",
    "july": "July", "august": "August", "september": "September",
    "october": "October", "november": "November", "december": "December",
    "janvier": "January", "février": "February", "fevrier": "February",
    "mars": "March", "avril": "April", "mai": "May", "juin": "June",
    "juillet": "July", "août": "August", "aout": "August",
    "septembre": "September", "octobre": "October",
    "novembre": "November", "décembre": "December", "decembre": "December",
}

SHEET_PATTERNS = {
    "Budget Initial": ["budget initial", "budget"],
    "Fiche Stationnement": ["fiche stationnement", "fiche de stationnement", "stationnement", "1. fiche"],
    "Donnees Historiques": ["donnees historiques", "données historiques", "historiques", "historique", "2. donnees", "2. données"],
}

# ============================================================================
# VERIFIED ROW MAPPING: Donnees Historiques row -> P&L labels
# ============================================================================
DH_ROW_MAPPING = {
    12: ["Transient Revenue", "transient revenue"],
    13: ["Monthly Revenues", "monthly revenues"],
    14: ["Car-Wash Revenue", "car-wash revenue", "lave-auto"],
    15: ["Hotel Revenue", "hotel revenue", "revenus hotel"],
    16: ["Interests", "interests", "intérêts", "interets"],
    17: ["Miscellaneous", "miscellaneous", "autres revenus", "Other Monthly revenue", "other monthly revenue", "other revenue", "Violation", "violation"],
    20: ["Discount-Gratuities - Transient", "gratuities transient"],
    22: ["Discount-Gratuities - Monthly", "rabais", "discount monthly"],
    29: ["Parking wages", "parking wages", "salaire stationnement"],
    30: ["Other wages", "other wages", "salaire superviseur", "supervisor"],
    31: ["Training & Recr.", "training", "formation", "recrutement"],
    32: ["Uniforms", "uniforms", "uniformes"],
    35: ["R&M - Cleaning", "cleaning", "nettoyage"],
    36: ["R&M - General", "maintenance", "entretien stationnement"],
    37: ["R&M - Equipement", "equipment", "entretien équipement", "entretien equipement"],
    38: ["R&M - Signs", "signs", "signalisation", "signage"],
    39: ["R&M - Lines", "lines", "lignage", "line painting"],
    40: ["Snow Removal", "snow removal", "déneigement", "deneigement", "snow"],
    41: ["Parking supplies", "parking supplies", "fournitures stationnement", "fournitures"],
    42: ["Misc. Re-Billing", "re-billing", "refacturations diverses", "refacturations", "rebilling"],
    43: ["R&M - General", "amenagement", "aménagement stationnement", "aménagement"],
    46: ["Public services", "public services", "services publics", "utilities"],
    49: ["Office expenses", "office expenses", "fournitures de bureau", "fournitures bureau"],
    50: ["Telecommunication", "telecommunication", "telecommunications", "télécommunications", "telecom"],
    51: ["Rent", "rent", "loyer"],
    52: ["Travel expenses", "travel", "frais de déplacement", "frais de deplacement", "déplacement"],
    53: ["Credit Card fees", "credit card", "frais de cartes de crédit", "frais de cartes de credit", "cartes de crédit"],
    54: ["Bank fees", "bank fees", "intérêts et frais de banque", "interets et frais de banque", "frais de banque"],
    55: ["Cash transportation fees", "cash transportation", "transport de fonds", "transport fonds"],
    56: ["Claims", "claims", "réclamations", "reclamations"],
    57: ["Insurance & Guarantee", "insurance", "assurances et cautionnement", "assurance", "cautionnement"],
    58: ["Tax & license", "tax", "taxes et permis", "taxes", "permis", "license"],
    59: ["Professional services", "accounting", "comptabilité", "comptabilite", "professional services"],
    60: ["Equipment rent", "equipment rent", "location d'équipement", "location d'equipement", "location équipement"],
    61: ["Ad. & Promotion", "advertising", "publicité et promotion", "publicite et promotion", "promotion"],
    62: ["Percent Management fee", "management fee", "honoraires de gestion en pourcentage", "honoraires de gestion en %"],
    63: ["Management Fees (Basic)", "management fees basic", "honoraires de gestion de base", "honoraires de base"],
    64: ["Incentives", "incentives", "incitatif annuel", "incitatif", "incentive"],
    67: ["Depreciation", "depreciation", "amortissement"],
    68: ["Financial fees", "interest", "intérêts sur emprunts", "interets sur emprunts", "emprunts"],
    69: ["Security", "security", "sécurité", "securite"],
    70: ["Co-ownership expenses", "co-ownership", "frais de copropriété", "frais de copropriete", "copropriété"],
    71: ["Shuttle expenses", "shuttle", "frais de navettes", "navettes"],
    72: ["Computer services", "computer", "services informatiques", "informatiques"],
    73: ["Bad debts", "bad debts", "mauvaises créances", "mauvaises creances", "créances"],
    74: ["Dues & Subscription", "dues", "cotisations", "subscription"],
    76: ["Meal & Entertainment", "meal", "représentation repas", "representation repas", "repas", "entertainment"],
}

# ============================================================================
# FICHE STATIONNEMENT MAPPING
# ============================================================================
FICHE_STATIONNEMENT_MAP = [
    ("K17", ["Transient Revenue", "transient revenue"]),
    ("K18", ["Monthly Revenues", "monthly revenues"]),
    ("K19", ["Car-Wash Revenue", "car-wash revenue", "lave-auto"]),
    ("K20", ["Hotel Revenue", "hotel revenue", "revenus hotel"]),
    ("K21", ["Interests", "interests", "intérêts", "interets"]),
    ("K22", ["Miscellaneous", "miscellaneous", "autres revenus", "Other Monthly revenue", "other monthly revenue", "Violation", "violation"]),
    ("K23", ["Discount-Gratuities - Transient", "gratuities transient"]),
    ("K24", ["Discount-Gratuities - Monthly", "rabais", "discount monthly"]),
    ("K25", ["Other Monthly revenue", "other monthly revenue", "Miscellaneous", "miscellaneous"]),
    ("K26", ["TOTAL REVENUE", "Total Revenue", "total revenus", "TOTAL DES REVENUS"]),
]

# ============================================================================
# MONTHLY REPORT LABEL MAPPING
# ============================================================================
MONTHLY_REPORT_MAPPING = {
    "Monthly Revenues": "Monthly Revenues",
    "Gratuities - Monthlies": "Discount-Gratuities - Monthly",
    "Monthlies Collected by the Owner": "Monthly Revenues",
    "Others - Processing Fee": "Other Monthly revenue",
    "Transient Revenue - Day": "Transient Revenue",
    "Coin Box & Meter": "Transient Revenue",
    "Revenues Reimbursement": "Transient Revenue",
    "Gratuities - Transient": "Discount-Gratuities - Transient",
    "Validation": "Transient Revenue",
    "Hotel Revenues": "Hotel Revenue",
    "Shuttle Revenues": "Shuttle expenses",
    "Lave-Auto": "Car-Wash Revenue",
    "Miscellaneous": "Miscellaneous",
    "Revenues Violation": "Violation",
    "Lost card fees": "Miscellaneous",
    "Monthly Processing Fees": "Other Monthly revenue",
    "Evening Tickets": "Transient Revenue",
    "Others Tickets": "Transient Revenue",
    "Special Events": "Transient Revenue",
    "Week end visitors": "Transient Revenue",
    "Online reservation": "Transient Revenue",
    "Salaires stationnement": "Parking wages",
    "Salaires - Supervision": "Other wages",
    "Uniformes": "Uniforms",
    "Fournitures stationnements": "Parking supplies",
    "Billet de stationnement": "Parking supplies",
    "Entretien réparation - général": "R&M - General",
    "Sécurité": "Security",
    "Location d'équipement": "Equipment rent",
    "Assurances": "Insurance & Guarantee",
    "Vehicle Expenses": "Vehicle expenses",
    "Telecommunication": "Telecommunication",
    "Serv. info. - Général": "Computer services",
    "Publicité et promotion": "Ad. & Promotion",
    "Frais de banque & C.C.": "Credit Card fees",
    "Honoraires de gestion (base)": "Management Fees (Basic)",
    "Honoraire de gestion a %": "Percent Management fee",
    "Incitatifs": "Incentives",
}

# ============================================================================
# FILE TYPE HANDLERS
# ============================================================================

def is_excel_file(file_bytes_or_obj):
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.xlsx') or name.endswith('.xls') or name.endswith('.xlsm'):
            return True
    return False

def is_csv_file(file_bytes_or_obj):
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.csv') or name.endswith('.tsv'):
            return True
    return False

def is_pdf_file(file_bytes_or_obj):
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.pdf'):
            return True
    return False

def is_docx_file(file_bytes_or_obj):
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.docx'):
            return True
    return False

def is_txt_file(file_bytes_or_obj):
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith(('.txt', '.tsv', '.text')):
            return True
    return False

def get_file_bytes(uploaded_file):
    if hasattr(uploaded_file, 'read'):
        uploaded_file.seek(0)
        return uploaded_file.read()
    if hasattr(uploaded_file, 'getvalue'):
        return uploaded_file.getvalue()
    return uploaded_file

def read_excel_to_dataframe(uploaded_file):
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
    try:
        file_bytes = get_file_bytes(uploaded_file)
        text = file_bytes.decode('utf-8', errors='ignore')
        lines = text.strip().split('\n')
        if not lines:
            return None
        if '\t' in lines[0]:
            reader = csv.reader(io.StringIO(text), delimiter='\t')
            data = list(reader)
            if data and len(data) >= 2:
                headers = data[0]
                df = pd.DataFrame(data[1:], columns=headers)
                return {"Sheet1": df}
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
    if uploaded_file is None:
        return None, None
    if is_excel_file(uploaded_file):
        result = read_excel_to_dataframe(uploaded_file)
        if result:
            return result, "excel"
    if is_csv_file(uploaded_file):
        result = read_csv_to_dataframe(uploaded_file)
        if result:
            return result, "csv"
    if is_pdf_file(uploaded_file):
        result = read_pdf_to_dataframe(uploaded_file)
        if result:
            return result, "pdf"
    if is_docx_file(uploaded_file):
        result = read_docx_to_dataframe(uploaded_file)
        if result:
            return result, "docx"
    if is_txt_file(uploaded_file):
        result = read_txt_to_dataframe(uploaded_file)
        if result:
            return result, "txt"
    try:
        result = read_excel_to_dataframe(uploaded_file)
        if result:
            return result, "excel"
    except Exception:
        pass
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

def get_parking_codes_from_pnl(pnl_file):
    """
    Extract all parking codes from a P&L file or monthly report.
    Searches sheet names AND file content for codes.
    """
    sheets_dict, file_type = read_any_file_to_dataframes(pnl_file)
    codes = []
    
    if sheets_dict:
        # Search sheet names for codes
        for sheet_name in sheets_dict.keys():
            match = re.search(r'(CMO\d+)', sheet_name, re.IGNORECASE)
            if match:
                codes.append(match.group(1).upper())
            if 'LUNA' in sheet_name.upper():
                codes.append('LUNA')
        
        # Search content for CMO codes and M-pattern codes
        for sheet_name, df in sheets_dict.items():
            if df is None or len(df) == 0:
                continue
            for row_idx in range(min(10, len(df))):
                for col_idx in range(min(10, len(df.columns))):
                    try:
                        cell_text = str(df.iloc[row_idx, col_idx])
                        match = re.search(r'(CMO\d+)', cell_text, re.IGNORECASE)
                        if match:
                            code = match.group(1).upper()
                            if code not in codes:
                                codes.append(code)
                        match2 = re.search(r'\b(M\d{3})\b', cell_text)
                        if match2:
                            code = match2.group(1).upper()
                            if code not in codes:
                                codes.append(code)
                    except Exception:
                        continue
    
    # Remove duplicates, keep order
    seen = set()
    unique_codes = []
    for code in codes:
        if code not in seen:
            seen.add(code)
            unique_codes.append(code)
    
    return unique_codes

def find_sheet_by_pattern(wb, patterns):
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            pattern_clean = pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
            if pattern_clean in sheet_lower:
                return sheet_name
    return None

def find_sheet_in_dict(sheets_dict, parking_code):
    if not sheets_dict:
        return None
    for sheet_name in sheets_dict:
        if sheet_name.upper().strip() == parking_code.upper().strip():
            return sheet_name
    for sheet_name in sheets_dict:
        if parking_code.upper() in sheet_name.upper():
            return sheet_name
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
    if not filename:
        return None
    years = re.findall(r'(20\d{2})', filename)
    if years:
        return int(years[0])
    return None

def safe_float(value, default=0.0):
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '').replace('€', '')
            value = value.replace('(', '').replace(')', '').replace('\xa0', '')
            if value.startswith('-'):
                value = value[1:]
                return -safe_float(value, default)
        return float(value)
    except (ValueError, TypeError):
        return default

def clean_text_for_matching(text):
    if not text:
        return ""
    text = text.lower().strip()
    text = text.replace('é', 'e').replace('è', 'e').replace('ê', 'e').replace('ë', 'e')
    text = text.replace('à', 'a').replace('â', 'a').replace('ä', 'a')
    text = text.replace('î', 'i').replace('ï', 'i')
    text = text.replace('ô', 'o').replace('ö', 'o')
    text = text.replace('û', 'u').replace('ü', 'u')
    text = text.replace('ù', 'u')
    text = text.replace('ç', 'c')
    text = text.replace('œ', 'oe')
    text = re.sub(r'[^a-z0-9 ]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def read_year_mapping_from_template(wb):
    dh_sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
    if not dh_sheet_name:
        current_year = datetime.now().year
        year_map = {}
        for i in range(4):
            year_map[i] = current_year
        for i in range(4, 12):
            year_map[i] = current_year - 1
        return year_map
    ws = wb[dh_sheet_name]
    for row_idx in range(1, 20):
        year_map = {}
        for col_idx in range(2, 14):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is not None and not (isinstance(cell_value, str) and str(cell_value).startswith('=')):
                cell_str = str(cell_value).strip()
                year_match = re.search(r'(20\d{2})', cell_str)
                if year_match:
                    year_map[col_idx - 2] = int(year_match.group(1))
        if len(year_map) >= 6:
            return year_map
    current_year = datetime.now().year
    year_map = {}
    for i in range(4):
        year_map[i] = current_year
    for i in range(4, 12):
        year_map[i] = current_year - 1
    return year_map

def extract_month_year_from_text(text):
    if not text:
        return None, None
    text_lower = text.lower()
    year_match = re.search(r'(20\d{2})', text)
    year = int(year_match.group(1)) if year_match else None
    found_month = None
    for month_name in MONTH_NAMES_MAP:
        if month_name in text_lower:
            found_month = MONTH_NAMES_MAP[month_name]
            break
    return found_month, year

def extract_monthly_data_from_file(uploaded_file):
    result = {}
    month_name = None
    year = None
    file_bytes = get_file_bytes(uploaded_file)
    sheets_dict = None
    if is_excel_file(uploaded_file):
        sheets_dict = read_excel_to_dataframe(uploaded_file)
    elif is_pdf_file(uploaded_file):
        sheets_dict = read_pdf_to_dataframe(uploaded_file)
    elif is_csv_file(uploaded_file):
        sheets_dict = read_csv_to_dataframe(uploaded_file)
    else:
        sheets_dict = read_excel_to_dataframe(uploaded_file)
        if sheets_dict is None:
            sheets_dict = read_pdf_to_dataframe(uploaded_file)
        if sheets_dict is None:
            sheets_dict = read_csv_to_dataframe(uploaded_file)
    if sheets_dict is None:
        return result, (None, None)
    if hasattr(uploaded_file, 'name'):
        month_name, year = extract_month_year_from_text(uploaded_file.name)
    for sheet_name, df in sheets_dict.items():
        if df is None or len(df) == 0:
            continue
        if month_name is None or year is None:
            for row_idx in range(min(5, len(df))):
                for col_idx in range(min(10, len(df.columns))):
                    try:
                        cell_text = str(df.iloc[row_idx, col_idx])
                        m, y = extract_month_year_from_text(cell_text)
                        if m:
                            month_name = m
                        if y:
                            year = y
                    except Exception:
                        continue
        for row_idx in range(len(df)):
            try:
                for col_idx in range(min(5, len(df.columns))):
                    cell_text = str(df.iloc[row_idx, col_idx]).strip()
                    if not cell_text or len(cell_text) < 3:
                        continue
                    cell_clean = clean_text_for_matching(cell_text)
                    for monthly_label, standard_label in MONTHLY_REPORT_MAPPING.items():
                        monthly_clean = clean_text_for_matching(monthly_label)
                        if monthly_clean in cell_clean or cell_clean in monthly_clean:
                            for val_col in range(col_idx + 1, min(col_idx + 4, len(df.columns))):
                                val = safe_float(df.iloc[row_idx, val_col])
                                if val != 0:
                                    if standard_label in result:
                                        result[standard_label] += val
                                    else:
                                        result[standard_label] = val
                                    break
                            break
            except Exception:
                continue
    return result, (month_name, year)

def build_monthly_data_from_files(monthly_files):
    if not monthly_files:
        return None
    monthly_data = {}
    yearly_data = {}
    for uploaded_file in monthly_files:
        file_data, (month_name, year) = extract_monthly_data_from_file(uploaded_file)
        if not file_data or month_name is None:
            continue
        for label, value in file_data.items():
            if label not in monthly_data:
                monthly_data[label] = {}
            monthly_data[label][month_name] = value
            if label not in yearly_data:
                yearly_data[label] = 0
            yearly_data[label] += value
    if not monthly_data:
        return None
    return {'monthly': monthly_data, 'yearly': yearly_data}

# ============================================================================
# P&L DATA EXTRACTION
# ============================================================================

def extract_pnl_data_from_dataframe(df, sheet_name_hint=None):
    result = {'monthly': {}, 'yearly': {}}
    if df is None or len(df) == 0:
        return result
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
        if any(skip in label_lower for skip in ['date filter', 'uc filter', 'currency']):
            continue
        if re.search(r'\d{2}[-/]\d{2}[-/]\d{2}', label):
            continue
        has_any_data = False
        for col_idx in range(1, min(14, len(df.columns))):
            if safe_float(df.iloc[row_idx, col_idx]) != 0:
                has_any_data = True
                break
        if not has_any_data:
            continue
        monthly = {}
        for month_idx in range(12):
            col_idx = month_idx + 1
            if col_idx < len(df.columns):
                val = safe_float(df.iloc[row_idx, col_idx])
                monthly[MONTHS_EN[month_idx]] = val
        yearly_total = 0
        if len(df.columns) > 13:
            yearly_total = safe_float(df.iloc[row_idx, 13])
        elif len(df.columns) > 1:
            yearly_total = safe_float(df.iloc[row_idx, -1])
        clean_label = label.strip().replace('  ', ' ')
        result['monthly'][clean_label] = monthly
        result['yearly'][clean_label] = yearly_total
    return result

def extract_pnl_data(uploaded_file, parking_code):
    sheets_dict, file_type = read_any_file_to_dataframes(uploaded_file)
    if sheets_dict is None:
        return None, None
    sheet_name = find_sheet_in_dict(sheets_dict, parking_code)
    if sheet_name is None:
        for name in sheets_dict:
            if parking_code.upper() in name.upper():
                sheet_name = name
                break
    if sheet_name is None and sheets_dict:
        max_len = 0
        for name, df in sheets_dict.items():
            if len(df) > max_len:
                max_len = len(df)
                sheet_name = name
    if sheet_name is None:
        return None, None
    df = sheets_dict[sheet_name]
    pnl_data = extract_pnl_data_from_dataframe(df, sheet_name)
    return pnl_data, file_type

def find_pnl_value(pnl_data, label_alternatives):
    if pnl_data is None:
        return 0
    yearly = pnl_data.get('yearly', {})
    if not yearly:
        return 0
    clean_alts = [clean_text_for_matching(alt) for alt in label_alternatives]
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in yearly:
            key_clean = clean_text_for_matching(key)
            if alt_clean == key_clean:
                return yearly[key]
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in yearly:
            key_clean = clean_text_for_matching(key)
            if alt_clean in key_clean or key_clean in alt_clean:
                return yearly[key]
    return 0

def find_monthly_pnl_value(monthly_data, label_alternatives):
    if not monthly_data:
        return {}
    clean_alts = [clean_text_for_matching(alt) for alt in label_alternatives]
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in monthly_data:
            key_clean = clean_text_for_matching(key)
            if alt_clean == key_clean:
                return monthly_data[key]
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in monthly_data:
            key_clean = clean_text_for_matching(key)
            if alt_clean in key_clean or key_clean in alt_clean:
                shorter = min(len(alt_clean), len(key_clean))
                longer = max(len(alt_clean), len(key_clean))
                if shorter >= 5 and (shorter / longer) >= 0.5:
                    return monthly_data[key]
    return {}

def merge_monthly_data(current_year_data, previous_year_data, year_map):
    merged = {}
    current_year = datetime.now().year
    for month_idx, year in year_map.items():
        month_name = MONTHS_EN[month_idx]
        if year == current_year and current_year_data:
            source = current_year_data
        elif year == current_year - 1 and previous_year_data:
            source = previous_year_data
        elif year == current_year - 2 and previous_year_data:
            source = previous_year_data
        elif current_year_data:
            source = current_year_data
        else:
            continue
        if source and 'monthly' in source:
            for label, monthly_values in source['monthly'].items():
                if label not in merged:
                    merged[label] = {}
                merged[label][month_name] = monthly_values.get(month_name, 0)
    return merged

# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, previous_year_data, parking_code):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["❌ Budget Initial: Sheet not found"]
        ws = wb[sheet_name]
        total = 0
        if previous_year_data:
            total_alternatives = [
                "TOTAL REVENUE", "Total Revenue", "total revenus", "revenu total",
                "REVENU TOTAL", "Total Revenus", "TOTAL DES REVENUS", "Total des revenus"
            ]
            total = find_pnl_value(previous_year_data, total_alternatives)
        if total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f} (previous year)")
        else:
            backup_alternatives = ["Parking Revenue", "Total Revenus Bruts", "Revenue", "Total"]
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
    except Exception as e:
        updates.append(f"❌ Budget Initial: {str(e)}")
    return updates

def update_fiche_stationnement(wb, year_minus_2_data, parking_code, word_data=None):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["❌ Fiche Stationnement: Sheet not found"]
        ws = wb[sheet_name]
        if year_minus_2_data is None:
            updates.append("⚠️ Fiche Stationnement: Skipped - P&L from 2 years ago not uploaded")
            return updates
        for cell, pnl_labels in FICHE_STATIONNEMENT_MAP:
            yearly_value = find_pnl_value(year_minus_2_data, pnl_labels)
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00'
                updates.append(f"✅ {cell} = ${yearly_value:,.2f}")
            else:
                updates.append(f"⚠️ {cell} = Not found in P&L (tried: {pnl_labels[0]})")
    except Exception as e:
        updates.append(f"❌ Fiche Stationnement: {str(e)}")
    return updates

def update_donnees_historiques(wb, merged_monthly_data, parking_code):
    updates = []
    try:
        dh_sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not dh_sheet_name:
            return ["❌ Donnees Historiques: Sheet not found"]
        ws_dh = wb[dh_sheet_name]
        year_map = read_year_mapping_from_template(wb)
        updates.append(
            f"📅 Year mapping: Jan={year_map.get(0)}, Apr={year_map.get(3)}, "
            f"May={year_map.get(4)}, Dec={year_map.get(11)}"
        )
        if not merged_monthly_data:
            updates.append("⚠️ Donnees Historiques: No merged monthly data available")
            return updates
        cells_updated = 0
        rows_filled = []
        for dh_row, pnl_labels in DH_ROW_MAPPING.items():
            monthly_values = find_monthly_pnl_value(merged_monthly_data, pnl_labels)
            if not monthly_values:
                continue
            if all(v == 0 for v in monthly_values.values()):
                continue
            row_cells = 0
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values:
                    val = monthly_values[month_name]
                    if val != 0:
                        col_letter = get_column_letter(month_idx + 2)
                        cell_ref = f"{col_letter}{dh_row}"
                        ws_dh[cell_ref] = val
                        ws_dh[cell_ref].number_format = '#,##0.00'
                        cells_updated += 1
                        row_cells += 1
            if row_cells > 0:
                rows_filled.append(f"  Row {dh_row}: {pnl_labels[0]} ({row_cells} months)")
        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: {cells_updated} cells in {len(rows_filled)} rows")
            for row_info in rows_filled:
                updates.append(row_info)
        else:
            updates.append("⚠️ Donnees Historiques: No cells updated")
            updates.append(f"   P&L data has {len(merged_monthly_data)} labels available")
            pnl_sample = list(merged_monthly_data.keys())[:20]
            updates.append(f"   Available P&L labels: {pnl_sample}")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {str(e)}")
    return updates


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(
    excel_file,
    pnl_current_year=None,
    pnl_previous_year=None,
    pnl_two_years_ago=None,
    monthly_files_current=None,
    monthly_files_previous=None,
    parking_code=None,
    word_data=None
):
    updates = []
    
    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)
    
    if not parking_code:
        return None, ["❌ Could not determine parking code. Please select a parking code."]
    
    updates.append(f"🔍 Processing: {parking_code}")
    
    current_year_data = None
    previous_year_data = None
    two_years_ago_data = None
    
    if monthly_files_current and len(monthly_files_current) > 0:
        updates.append("📋 Using monthly report files (Current Year)")
        current_year_data = build_monthly_data_from_files(monthly_files_current)
        if current_year_data:
            updates.append(f"📊 Current year from monthlies: {len(current_year_data['yearly'])} labels")
    
    if monthly_files_previous and len(monthly_files_previous) > 0:
        updates.append("📋 Using monthly report files (Previous Year)")
        previous_year_data = build_monthly_data_from_files(monthly_files_previous)
        if previous_year_data:
            updates.append(f"📊 Previous year from monthlies: {len(previous_year_data['yearly'])} labels")
    
    if pnl_current_year and current_year_data is None:
        current_year_data, current_file_type = extract_pnl_data(pnl_current_year, parking_code)
        if current_year_data:
            keys = list(current_year_data['yearly'].keys())[:10]
            updates.append(f"📊 Current year keys ({current_file_type}): {keys}")
    
    if pnl_previous_year and previous_year_data is None:
        previous_year_data, prev_file_type = extract_pnl_data(pnl_previous_year, parking_code)
        if previous_year_data:
            keys = list(previous_year_data['yearly'].keys())[:10]
            updates.append(f"📊 Previous year keys ({prev_file_type}): {keys}")
    
    if pnl_two_years_ago:
        two_years_ago_data, two_ya_file_type = extract_pnl_data(pnl_two_years_ago, parking_code)
        if two_years_ago_data:
            keys = list(two_years_ago_data['yearly'].keys())[:10]
            updates.append(f"📊 2YA keys ({two_ya_file_type}): {keys}")
    
    if current_year_data is None and previous_year_data is None and two_years_ago_data is None:
        return None, [f"❌ Could not find P&L data for {parking_code} in any uploaded file."]
    
    try:
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        file_bytes = excel_file.read()
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        wb_write = load_workbook(io.BytesIO(file_bytes), data_only=False)
        wb_read = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, [f"❌ Error reading template: {str(e)}"]
    
    year_map = read_year_mapping_from_template(wb_read)
    
    merged_monthly = {}
    if year_map and current_year_data and previous_year_data:
        merged_monthly = merge_monthly_data(current_year_data, previous_year_data, year_map)
    
    if not merged_monthly and current_year_data:
        merged_monthly = current_year_data['monthly']
    
    budget_initial_data = previous_year_data if previous_year_data else current_year_data
    
    fiche_data = None
    if pnl_two_years_ago and two_years_ago_data:
        fiche_data = two_years_ago_data
    
    dh_data = merged_monthly if merged_monthly else {}
    if not dh_data and current_year_data:
        dh_data = current_year_data['monthly']
    
    updates.extend(update_budget_initial(wb_write, budget_initial_data, parking_code))
    updates.extend(update_fiche_stationnement(wb_write, fiche_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb_write, dh_data, parking_code))
    
    success_count = sum(1 for u in updates if u.startswith("✅"))
    if success_count == 0:
        updates.append("💡 No updates were made.")
    
    if pnl_current_year and hasattr(pnl_current_year, 'seek'):
        pnl_current_year.seek(0)
    if pnl_previous_year and hasattr(pnl_previous_year, 'seek'):
        pnl_previous_year.seek(0)
    if pnl_two_years_ago and hasattr(pnl_two_years_ago, 'seek'):
        pnl_two_years_ago.seek(0)
    
    output = io.BytesIO()
    wb_write.save(output)
    output.seek(0)
    
    return output, updates
