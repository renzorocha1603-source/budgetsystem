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
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
]

MONTHS_FR = [
    "janvier",
    "février",
    "fevrier",
    "mars",
    "avril",
    "mai",
    "juin",
    "juillet",
    "août",
    "aout",
    "septembre",
    "octobre",
    "novembre",
    "décembre",
    "decembre"
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

# ============================================================================
# DONNEES HISTORIQUES LABEL MAPPING
# ============================================================================

DONNEES_HISTORIQUES_MAP = [
    ("Transient Revenue", [
        "revenus horaires",
        "revenus horaire",
        "transient",
        "transitoire",
        "stationnement horaire"
    ]),
    ("Monthly Revenues", [
        "revenus mensuels",
        "revenus mensuel",
        "monthly",
        "mensuels",
        "stationnement mensuel"
    ]),
    ("Car-Wash Revenue", [
        "revenus lave-auto",
        "revenus lave auto",
        "lave-auto",
        "lave auto",
        "car-wash",
        "car wash",
        "lavage"
    ]),
    ("Hotel Revenue", [
        "revenus hôtel",
        "revenus hotel",
        "revenus d'hôtel",
        "hôtel",
        "hotel",
        "hotelier"
    ]),
    ("Interests", [
        "revenus d'intérêts",
        "revenus d'interets",
        "revenus d'intérêt",
        "intérêts",
        "interets",
        "interests",
        "interest"
    ]),
    ("Miscellaneous", [
        "autres revenus",
        "autre revenus",
        "miscellaneous",
        "misc",
        "divers"
    ]),
    ("Discount-Gratuities - Transient", [
        "gratuités",
        "gratuites",
        "gratuités",
        "gratuities",
        "discount",
        "escomptes",
        "courtoisies"
    ]),
    ("Discount-Gratuities - Monthly", [
        "rabais",
        "discount monthly",
        "escomptes mensuels",
        "rabais mensuel"
    ]),
    ("Parking wages", [
        "salaire stationnement",
        "parking wages",
        "salaires stationnement",
        "main d'oeuvre",
        "main d'œuvre",
        "masse salariale"
    ]),
    ("Total Operation expenses", [
        "total dépenses",
        "total depenses",
        "total operation expenses",
        "operation expenses",
        "dépenses d'opération",
        "depenses d'operation",
        "total des dépenses",
        "total dépenses"
    ]),
    ("Percent Management fee", [
        "honoraires de gestion en pourcentage",
        "percent management fee",
        "frais de gestion en %",
        "% management",
        "management fee",
        "honoraires de gestion"
    ]),
]

DONNEES_HISTORIQUES_EXTRA = [
    ("Other wages", [
        "salaire superviseur",
        "supervisor wages",
        "salaire supervision"
    ]),
    ("Training & Recr.", [
        "formation & recrutement",
        "formation et recrutement",
        "training",
        "recrutement"
    ]),
    ("Uniforms", [
        "uniformes",
        "uniforms",
        "uniformes"
    ]),
    ("R&M - Cleaning", [
        "nettoyage stationnement",
        "nettoyage",
        "cleaning"
    ]),
    ("R&M - Equipement", [
        "entretien stationnement",
        "entretien équipement",
        "entretien equipement",
        "equipment maintenance"
    ]),
    ("R&M - Signs", [
        "signalisation",
        "signs",
        "signage"
    ]),
    ("R&M - Lines", [
        "lignage",
        "lines",
        "line painting"
    ]),
    ("Snow Removal", [
        "déneigement",
        "deneigement",
        "snow removal",
        "snow"
    ]),
    ("Parking supplies", [
        "fournitures stationnement",
        "parking supplies",
        "fournitures"
    ]),
    ("Misc. Re-Billing", [
        "refacturations diverses",
        "refacturations",
        "re-billing",
        "rebilling"
    ]),
    ("R&M - General", [
        "aménagement stationnement",
        "amenagement stationnement",
        "general maintenance"
    ]),
    ("Public services", [
        "services publics",
        "public services",
        "utilities"
    ]),
    ("Office expenses", [
        "fournitures de bureau",
        "office expenses",
        "fournitures bureau"
    ]),
    ("Telecommunication", [
        "telecommunications",
        "télécommunications",
        "telecom"
    ]),
    ("Rent", [
        "loyer",
        "rent",
        "loyer"
    ]),
    ("Travel expenses", [
        "frais de déplacement",
        "frais de deplacement",
        "travel",
        "déplacement"
    ]),
    ("Credit Card fees", [
        "frais de cartes de crédit",
        "frais de cartes de credit",
        "credit card fees",
        "cartes de crédit"
    ]),
    ("Bank fees", [
        "intérêts et frais de banque",
        "interets et frais de banque",
        "bank fees",
        "frais de banque"
    ]),
    ("Cash transportation fees", [
        "transport de fonds",
        "cash transportation",
        "transport fonds"
    ]),
    ("Claims", [
        "réclamations",
        "reclamations",
        "claims"
    ]),
    ("Insurance & Guarantee", [
        "assurances et cautionnement",
        "assurance",
        "insurance",
        "cautionnement"
    ]),
    ("Tax & license", [
        "taxes et permis",
        "taxes",
        "tax",
        "permis"
    ]),
    ("Professional services", [
        "comptabilité",
        "comptabilite",
        "professional services",
        "accounting"
    ]),
    ("Equipment rent", [
        "location d'équipement",
        "location d'equipement",
        "equipment rent",
        "location équipement"
    ]),
    ("Ad. & Promotion", [
        "publicité et promotion",
        "publicite et promotion",
        "advertising",
        "promotion"
    ]),
    ("Management Fees (Basic)", [
        "honoraires de gestion de base",
        "management fees basic",
        "frais de gestion de base",
        "honoraires de base"
    ]),
    ("Incentives", [
        "incitatif annuel",
        "incitatif",
        "incentives",
        "incentive"
    ]),
    ("Depreciation", [
        "amortissement",
        "depreciation",
        "amortissement"
    ]),
    ("Security", [
        "sécurité",
        "securite",
        "security"
    ]),
    ("Co-ownership expenses", [
        "frais de copropriété",
        "frais de copropriete",
        "co-ownership",
        "copropriété"
    ]),
    ("Shuttle expenses", [
        "frais de navettes",
        "navettes",
        "shuttle"
    ]),
    ("Computer services", [
        "services informatiques",
        "computer services",
        "informatiques"
    ]),
    ("Dues & Subscription", [
        "cotisations",
        "cotisations",
        "dues",
        "subscription"
    ]),
    ("Meal & Entertainment", [
        "représentation repas",
        "representation repas",
        "meal",
        "repas",
        "entertainment"
    ]),
]

ALL_DONNEES_MAP = DONNEES_HISTORIQUES_MAP + DONNEES_HISTORIQUES_EXTRA

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
        if name.endswith('.xlsx'):
            return True
        if name.endswith('.xls'):
            return True
        if name.endswith('.xlsm'):
            return True
    return False


def is_csv_file(file_bytes_or_obj):
    """Check if file is CSV format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.csv'):
            return True
        if name.endswith('.tsv'):
            return True
    return False


def is_pdf_file(file_bytes_or_obj):
    """Check if file is PDF format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.pdf'):
            return True
    return False


def is_docx_file(file_bytes_or_obj):
    """Check if file is DOCX format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.docx'):
            return True
    return False


def is_txt_file(file_bytes_or_obj):
    """Check if file is text format"""
    if hasattr(file_bytes_or_obj, 'name'):
        name = file_bytes_or_obj.name.lower()
        if name.endswith('.txt'):
            return True
        if name.endswith('.tsv'):
            return True
        if name.endswith('.text'):
            return True
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
    """
    Universal file reader - handles Excel, PDF, CSV, DOCX, TXT, and more.
    Returns (dict of {sheet_name: DataFrame}, file_type) or (None, None) if unreadable.
    """
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
    """Try to extract year from filename like 'P&L_2025.xlsx' or '2025_P&L.pdf'"""
    if not filename:
        return None
    years = re.findall(r'(20\d{2})', filename)
    if years:
        return int(years[0])
    return None


def safe_float(value, default=0.0):
    """Safely convert a value to float, handling strings, None, and special characters"""
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
    """
    Clean text for label matching: lowercase, remove accents, remove special chars,
    normalize whitespace.
    """
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


def debug_dump_sheet(ws, sheet_name, max_rows=60, max_cols=15):
    """
    Debug function: dump all cell values (including formulas) from a worksheet.
    Returns a list of strings showing what's in each cell.
    """
    lines = []
    lines.append(f"=== DEBUG DUMP: Sheet '{sheet_name}' ===")
    lines.append(f"Max rows: {ws.max_row}, Max cols: {ws.max_column}")
    
    for row_idx in range(1, min(max_rows + 1, ws.max_row + 1)):
        row_parts = []
        for col_idx in range(1, min(max_cols + 1, ws.max_column + 1)):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            
            if val is None:
                row_parts.append("None")
            elif isinstance(val, str) and val.startswith('='):
                row_parts.append(f"[F:{val[:35]}]")
            else:
                str_val = str(val).strip()
                if str_val:
                    row_parts.append(f"'{str_val[:25]}'")
                else:
                    row_parts.append("''")
        
        lines.append(f"R{row_idx}: {' | '.join(row_parts)}")
    
    return lines


def read_year_mapping_from_template(ws):
    """
    Read the year mapping from Donnees Historiques template.
    Looks for the 'Année' row that has years like 2026, 2025 in columns B-M.
    Returns dict: {month_index: year} or fallback if not found.
    """
    # First, find the "Année" row
    annee_row = None
    for row_idx in range(1, 55):
        cell_value = ws.cell(row=row_idx, column=1).value
        if cell_value is not None:
            cell_str = str(cell_value).strip().lower()
            cell_clean = clean_text_for_matching(cell_str)
            if cell_clean == "annee" or cell_clean == "année":
                annee_row = row_idx
                break
            if "annee" in cell_clean or "année" in cell_clean:
                annee_row = row_idx
                break
    
    if annee_row:
        year_map = {}
        for col_idx in range(2, 14):
            cell_value = ws.cell(row=annee_row, column=col_idx).value
            if cell_value is not None:
                cell_str = str(cell_value).strip()
                year_match = re.search(r'(20\d{2})', cell_str)
                if year_match:
                    year_map[col_idx - 2] = int(year_match.group(1))
        if len(year_map) >= 6:
            return year_map
    
    # Fallback: search rows 35-55 for any row with years
    for row_idx in range(35, 55):
        year_map = {}
        for col_idx in range(2, 14):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is not None:
                cell_str = str(cell_value).strip()
                year_match = re.search(r'(20\d{2})', cell_str)
                if year_match:
                    year_map[col_idx - 2] = int(year_match.group(1))
        if len(year_map) >= 6:
            return year_map
    
    # Ultimate fallback
    current_year = datetime.now().year
    year_map = {}
    for i in range(4):
        year_map[i] = current_year
    for i in range(4, 12):
        year_map[i] = current_year - 1
    return year_map


# ============================================================================
# P&L DATA EXTRACTION - Works with ANY file format
# ============================================================================

def extract_pnl_data_from_dataframe(df, sheet_name_hint=None):
    """
    Extract P&L data from a pandas DataFrame.
    Returns dict with 'monthly' and 'yearly' data.
    """
    result = {
        'monthly': {},
        'yearly': {}
    }
    
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
    """
    Extract P&L data from ANY file format.
    Returns (pnl_data_dict, file_type) or (None, None).
    """
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
    """
    Search through P&L data for a matching label.
    Returns the yearly value or 0.
    """
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


def find_monthly_pnl_value(pnl_data, label_alternatives):
    """
    Search through P&L monthly data for a matching label.
    Returns the monthly values dict or empty dict.
    """
    if pnl_data is None:
        return {}
    
    monthly = pnl_data.get('monthly', {})
    
    if not monthly:
        return {}
    
    clean_alts = [clean_text_for_matching(alt) for alt in label_alternatives]
    
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in monthly:
            key_clean = clean_text_for_matching(key)
            if alt_clean == key_clean:
                return monthly[key]
    
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in monthly:
            key_clean = clean_text_for_matching(key)
            if alt_clean in key_clean or key_clean in alt_clean:
                return monthly[key]
    
    return {}


def merge_monthly_data(current_year_data, previous_year_data, year_map):
    """
    Merge monthly data from current and previous years based on year_map.
    Returns a combined dict with structure:
    {pnl_label: {"January": value, "February": value, ...}}
    """
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
                "Total des revenus"
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
    """
    Update Fiche Stationnement sheet.
    Uses data from 2 YEARS AGO.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["❌ Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        
        if year_minus_2_data is None:
            updates.append("⚠️ Fiche Stationnement: No data available")
            return updates
        
        for pnl_label, cell in FICHE_STATIONNEMENT_MAP:
            yearly_value = find_pnl_value(year_minus_2_data, [pnl_label])
            
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00'
                updates.append(f"✅ {cell} = {pnl_label}: ${yearly_value:,.2f}")
            else:
                updates.append(f"⚠️ {cell} = {pnl_label}: Not found in P&L")
        
        total_revenue = find_pnl_value(
            year_minus_2_data,
            ["TOTAL REVENUE", "Total Revenue", "total revenus"]
        )
        ws["K26"] = total_revenue
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"✅ K26 = TOTAL REVENUE: ${total_revenue:,.2f}")
    except Exception as e:
        updates.append(f"❌ Fiche Stationnement: {str(e)}")
    return updates


def update_donnees_historiques(wb, merged_monthly_data, parking_code):
    """
    Update Donnees Historiques sheet with merged monthly data.
    The template has French labels that need to be matched.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["❌ Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # DEBUG: dump the sheet contents to understand its structure
        debug_lines = debug_dump_sheet(ws, sheet_name, max_rows=50, max_cols=15)
        for line in debug_lines[:25]:
            updates.append(f"🔧 {line}")
        
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
        
        if not merged_monthly_data:
            updates.append("⚠️ Donnees Historiques: No merged monthly data available")
            return updates
        
        cells_updated = 0
        rows_filled = []
        
        for pnl_label, french_labels in ALL_DONNEES_MAP:
            # Find the monthly data for this P&L label
            monthly_values = find_monthly_pnl_value(
                {"monthly": merged_monthly_data, "yearly": {}},
                [pnl_label]
            )
            
            if not monthly_values:
                clean_pnl = clean_text_for_matching(pnl_label)
                for key in merged_monthly_data:
                    key_clean = clean_text_for_matching(key)
                    if clean_pnl in key_clean or key_clean in clean_pnl:
                        monthly_values = merged_monthly_data[key]
                        break
            
            if not monthly_values:
                continue
            
            if all(v == 0 for v in monthly_values.values()):
                continue
            
            # Search columns A and B for matching label
            found_row = None
            for row_idx in range(1, 80):
                for col_check in [1, 2]:
                    cell_val = ws.cell(row=row_idx, column=col_check).value
                    if cell_val is not None and not (isinstance(cell_val, str) and str(cell_val).startswith('=')):
                        cell_clean = clean_text_for_matching(str(cell_val))
                        for fl in french_labels:
                            fl_clean = clean_text_for_matching(fl)
                            if fl_clean and len(fl_clean) >= 3:
                                if fl_clean in cell_clean:
                                    found_row = row_idx
                                    break
                        if found_row:
                            break
                if found_row:
                    break
            
            if found_row is None:
                continue
            
            row_cells = 0
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values:
                    val = monthly_values[month_name]
                    if val != 0:
                        col_letter = get_column_letter(month_idx + 2)
                        cell_ref = f"{col_letter}{found_row}"
                        ws[cell_ref] = val
                        ws[cell_ref].number_format = '#,##0.00'
                        cells_updated += 1
                        row_cells += 1
            
            if row_cells > 0:
                rows_filled.append(f"  Row {found_row}: {pnl_label} ({row_cells} months)")
        
        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: {cells_updated} cells updated in {len(rows_filled)} rows")
            for row_info in rows_filled[:20]:
                updates.append(row_info)
            if len(rows_filled) > 20:
                updates.append(f"  ... and {len(rows_filled) - 20} more rows")
        else:
            updates.append("⚠️ Donnees Historiques: No cells updated")
            updates.append(f"   P&L data has {len(merged_monthly_data)} labels available")
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
    
    Year Usage per Sheet:
        - Budget Initial: Previous year total (year - 1)
        - Fiche Stationnement: 2 years ago data (year - 2)
        - Donnees Historiques: Jan-Apr = current year, May-Dec = previous year
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
    
    if current_year_data is None and previous_year_data is None and two_years_ago_data is None:
        return None, [f"❌ Could not find P&L data for {parking_code} in any uploaded file."]
    
    if current_year_data:
        keys = list(current_year_data['yearly'].keys())[:10]
        updates.append(f"📊 Current year keys ({current_file_type}): {keys}")
    
    if previous_year_data:
        keys = list(previous_year_data['yearly'].keys())[:10]
        updates.append(f"📊 Previous year keys ({prev_file_type}): {keys}")
    
    if two_years_ago_data:
        keys = list(two_years_ago_data['yearly'].keys())[:10]
        updates.append(f"📊 2YA keys ({two_ya_file_type}): {keys}")
    
    # ── Read template ───────────────────────────────────────────────────
    try:
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        file_bytes = excel_file.read()
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        
        wb_write = load_workbook(io.BytesIO(file_bytes), data_only=False)
        wb_read = load_workbook(io.BytesIO(file_bytes), data_only=True)
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
        merged_monthly = merge_monthly_data(current_year_data, previous_year_data, year_map)
    
    if not merged_monthly and current_year_data:
        merged_monthly = current_year_data['monthly']
    
    # ── Determine which data to use for each sheet ──────────────────────
    budget_initial_data = previous_year_data if previous_year_data else current_year_data
    
    fiche_data = two_years_ago_data
    if fiche_data is None:
        fiche_data = previous_year_data
    if fiche_data is None:
        fiche_data = current_year_data
    
    dh_data = merged_monthly if merged_monthly else {}
    if not dh_data and current_year_data:
        dh_data = current_year_data['monthly']
    
    # ── Update all sheets ───────────────────────────────────────────────
    updates.extend(update_budget_initial(wb_write, budget_initial_data, parking_code))
    updates.extend(update_fiche_stationnement(wb_write, fiche_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb_write, dh_data, parking_code))
    
    # ── Summary ─────────────────────────────────────────────────────────
    success_count = sum(1 for u in updates if u.startswith("✅"))
    
    if success_count == 0:
        updates.append("💡 No updates were made.")
    
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
