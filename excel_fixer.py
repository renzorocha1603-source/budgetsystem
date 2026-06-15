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

# Month mapping: English month name -> template column index (0=January/Janvier)
MONTH_TO_COL = {
    "January": 0, "February": 1, "March": 2, "April": 3,
    "May": 4, "June": 5, "July": 6, "August": 7,
    "September": 8, "October": 9, "November": 10, "December": 11
}

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
# Maps P&L data labels (English) to template French labels
# Format: (pnl_label_english, [french_labels_to_match_in_template])
DONNEES_HISTORIQUES_MAP = [
    # REVENUS section
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
    # Gratuités / Discounts
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
    # DÉPENSES section
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

# Additional labels for Donnees Historiques template
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

# Combined mapping
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
    # Replace common French accents
    text = text.replace('é', 'e').replace('è', 'e').replace('ê', 'e').replace('ë', 'e')
    text = text.replace('à', 'a').replace('â', 'a').replace('ä', 'a')
    text = text.replace('î', 'i').replace('ï', 'i')
    text = text.replace('ô', 'o').replace('ö', 'o')
    text = text.replace('û', 'u').replace('ü', 'u')
    text = text.replace('ù', 'u')
    text = text.replace('ç', 'c')
    text = text.replace('œ', 'oe')
    # Remove all non-alphanumeric except spaces
    text = re.sub(r'[^a-z0-9 ]', ' ', text)
    # Normalize whitespace
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
                # Show first 40 chars of formula
                row_parts.append(f"[FORMULA:{val[:40]}]")
            else:
                # Show first 30 chars of value
                str_val = str(val).strip()
                if str_val:
                    row_parts.append(f"'{str_val[:30]}'")
                else:
                    row_parts.append("''")
        
        lines.append(f"Row {row_idx}: {' | '.join(row_parts)}")
    
    return lines


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
    
    # Find header row (where month names are)
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
        
        # Skip obvious non-data rows
        if label_lower in ['code', 'profit & loss', '', 'nan', 'none']:
            continue
        
        # Skip rows that are section headers or date filters
        if any(skip in label_lower for skip in ['date filter', 'uc filter', 'currency', 'code']):
            continue
        
        # Skip date rows
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
            yearly_total = safe_float(df.iloc[row_idx, -1])
        
        # Use cleaned label as key
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
    Uses cleaned text matching (removes accents, special chars).
    Returns the yearly value or 0.
    """
    if pnl_data is None:
        return 0
    
    yearly = pnl_data.get('yearly', {})
    
    if not yearly:
        return 0
    
    clean_alts = [clean_text_for_matching(alt) for alt in label_alternatives]
    
    # Try exact match first (cleaned)
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in yearly:
            key_clean = clean_text_for_matching(key)
            if alt_clean == key_clean:
                return yearly[key]
    
    # Try partial match
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
    
    # Try exact match first
    for alt_clean in clean_alts:
        if not alt_clean or len(alt_clean) < 3:
            continue
        for key in monthly:
            key_clean = clean_text_for_matching(key)
            if alt_clean == key_clean:
                return monthly[key]
    
    # Try partial match
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
                "Total des revenus",
                "Revenus Totals"
            ]
            total = find_pnl_value(previous_year_data, total_alternatives)
        
        if total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f} (previous year)")
        else:
            backup_alternatives = [
                "Parking Revenue",
                "Total Revenus Bruts",
                "Revenue",
                "Total"
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
            ["TOTAL REVENUE", "Total Revenue", "total revenus", "Total des revenus"]
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
    
    IMPORTANT: The Donnees Historiques template may have labels in COLUMN B instead of A,
    or labels may be in merged cells, or labels may be formula references to other sheets.
    
    This function tries multiple strategies:
    1. Look for labels in Column A
    2. Look for labels in Column B
    3. Use FIXED ROW MAPPING as fallback based on template structure
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["❌ Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # DEBUG: Dump the sheet contents to understand its structure
        debug_lines = debug_dump_sheet(ws, sheet_name, max_rows=50, max_cols=15)
        # Only show first 20 lines to avoid flooding
        for line in debug_lines[:20]:
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
        
        # ── STRATEGY 1: Try Column A for labels ─────────────────────────
        cells_updated = 0
        rows_filled = []
        
        for pnl_label, french_labels in ALL_DONNEES_MAP:
            # Find the monthly data for this P&L label
            monthly_values = find_monthly_pnl_value(
                {"monthly": merged_monthly_data, "yearly": {}},
                [pnl_label]
            )
            
            if not monthly_values:
                # Try searching in merged data directly
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
            
            # Search Column A for matching label
            found_row = None
            for row_idx in range(1, 80):
                cell_val_a = ws.cell(row=row_idx, column=1).value
                cell_val_b = ws.cell(row=row_idx, column=2).value
                
                # Check Column A
                if cell_val_a is not None and not (isinstance(cell_val_a, str) and cell_val_a.startswith('=')):
                    cell_clean = clean_text_for_matching(str(cell_val_a))
                    for fl in french_labels:
                        fl_clean = clean_text_for_matching(fl)
                        if fl_clean and len(fl_clean) >= 3 and fl_clean in cell_clean:
                            found_row = row_idx
                            break
                
                # Check Column B (some templates put labels in B)
                if found_row is None and cell_val_b is not None and not (isinstance(cell_val_b, str) and cell_val_b.startswith('=')):
                    cell_clean = clean_text_for_matching(str(cell_val_b))
                    for fl in french_labels:
                        fl_clean = clean_text_for_matching(fl)
                        if fl_clean and len(fl_clean) >= 3 and fl_clean in cell_clean:
                            found_row = row_idx
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
                        col_letter = get_column_letter(month_idx + 2)  # B=2
                        cell_ref = f"{col_letter}{found_row}"
                        ws[cell_ref] = val
                        ws[cell_ref].number_format = '#,##0.00'
                        cells_updated += 1
                        row_cells += 1
            
            if row_cells > 0:
                rows_filled.append(f"  Row {found_row}: {pnl_label} ({row_cells} months)")
        
        # ── STRATEGY 2: If no matches, try FIXED ROW MAPPING ────────────
        if cells_updated == 0:
            updates.append("⚠️ Label matching failed, trying fixed row mapping...")
            
            # Based on the template structure you showed, the labels are at these approximate rows:
            # Row 10: Revenus horaires (Transient Revenue)
            # Row 11: Revenus mensuels (Monthly Revenues)
            # Row 12: Revenus Lave-auto (Car-Wash Revenue)
            # Row 13: Revenus hôtel (Hotel Revenue)
            # Row 14: Revenus d'intérêts (Interests)
            # Row 15: Autres revenus (Miscellaneous)
            # Row 16: Total revenus Bruts (Parking Revenue)
            # Row 18: (Gratuités) (Discount-Gratuities - Transient)
            # Row 20: (Rabais) (Discount-Gratuities - Monthly)
            # Row 23: TOTAL REVENUS
            # Row 27: Salaire Stationnement (Parking wages)
            # ... etc
            
            # Instead of guessing rows, let's do a MORE AGGRESSIVE search
            # Check EVERY cell in columns A and B for ANY text that might match
            
            for pnl_label, french_labels in ALL_DONNEES_MAP:
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
                
                # Search ALL rows for ANY match
                found_row = None
                for row_idx in range(1, 80):
                    for col_check in [1, 2]:  # Check columns A and B
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
        else:
            updates.append("⚠️ Donnees Historiques: No cells updated after all strategies")
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
    
    # ── Detect years ────────────────────────────────────────────────────
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
    
    # ── Extract P&L data ────────────────────────────────────────────────
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
        
        # Use data_only=False for writing (preserves formulas)
        wb_write = load_workbook(io.BytesIO(file_bytes), data_only=False)
        # Use data_only=True for reading (gets computed values)
        wb_read = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, [f"❌ Error reading template: {str(e)}"]
    
    # ── Read year mapping ───────────────────────────────────────────────
    dh_sheet_name = find_sheet_by_pattern(wb_read, SHEET_PATTERNS["Donnees Historiques"])
    year_map = None
    if dh_sheet_name:
        year_map = read_year_mapping_from_template(wb_read[dh_sheet_name])
    
    # ── Merge monthly data ──────────────────────────────────────────────
    merged_monthly = {}
    if year_map and current_year_data and previous_year_data:
        merged_monthly = merge_monthly_data(current_year_data, previous_year_data, year_map)
    
    if not merged_monthly and current_year_data:
        merged_monthly = current_year_data['monthly']
    
    # ── Assign data to sheets ───────────────────────────────────────────
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
