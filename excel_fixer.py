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
# VERIFIED ROW MAPPING
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
    37: ["R&M - Equipement", "entretien équipement", "entretien equipement"],
    38: ["R&M - Signs", "signs", "signalisation", "signage"],
    39: ["R&M - Lines", "lines", "lignage", "line painting"],
    40: ["Snow Removal", "snow removal", "déneigement", "deneigement", "snow"],
    41: ["Parking supplies", "parking supplies", "fournitures stationnement", "fournitures"],
    42: ["Misc. Re-Billing", "re-billing", "refacturations diverses", "refacturations", "rebilling"],
    43: ["R&M - General", "amenagement", "aménagement stationnement", "aménagement"],
    46: ["Public services", "public services", "services publics", "utilities"],
    49: ["Office expenses", "office expenses", "fournitures de bureau", "fournitures bureau"],
    50: ["Telecommunication", "telecommunication", "telecommunications", "télécommunications", "telecom"],
    51: ["Rent", "loyer"],
    52: ["Travel expenses", "travel", "frais de déplacement", "frais de deplacement", "déplacement"],
    53: ["Credit Card fees", "credit card", "frais de cartes de crédit", "frais de cartes de credit", "cartes de crédit"],
    54: ["Bank fees", "bank fees", "intérêts et frais de banque", "interets et frais de banque", "frais de banque"],
    55: ["Cash transportation fees", "cash transportation", "transport de fonds", "transport fonds"],
    56: ["Claims", "claims", "réclamations", "reclamations"],
    57: ["Insurance & Guarantee", "insurance", "assurances et cautionnement", "assurance", "cautionnement"],
    58: ["Tax & license", "tax", "taxes et permis", "taxes", "permis", "license"],
    59: ["Professional services", "accounting", "comptabilité", "comptabilite", "professional services"],
    60: ["Equipment rent", "location d'équipement", "location d'equipement", "location équipement"],
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

REVENUE_ROWS = [12, 13, 14, 15, 16, 17, 20, 22]
EXPENSE_ROWS = [r for r in DH_ROW_MAPPING.keys() if r not in REVENUE_ROWS]
REVENUE_CATCH_ALL_ROW = 17
EXPENSE_CATCH_ALL_ROW = 76

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
# EXACT FRENCH LABELS FROM PAGE 10
# ============================================================================
PAGE10_FRENCH_LABELS = {
    # REVENUS
    "Revenus mensuels": "Monthly Revenues",
    "Revenus Journaliers": "Transient Revenue",
    "Revenus Lave-Auto": "Car-Wash Revenue",
    "Divers": "Miscellaneous",
    "Revenus de stationnement": "Parking Revenue",
    "(Gratuités - mensuels)": "Discount-Gratuities - Monthly",
    "TOTAL REVENUS": "TOTAL REVENUE",
    # DÉPENSES
    "Salaires Stationnement": "Parking wages",
    "Uniformes": "Uniforms",
    "Fourn. de stationnement": "Parking supplies",
    "Entretien réparation - Nettoyage": "R&M - Cleaning",
    "Entretien réparation - Equipement": "R&M - Equipement",
    "Entretien réparation - Général": "R&M - General",
    "Taxes et permis": "Tax & license",
    "Assurances Cautionnement": "Insurance & Guarantee",
    "Réclamations": "Claims",
    "Télécommunication": "Telecommunication",
    "Frais de cartes de crédit": "Credit Card fees",
    "Frais de bureau": "Office expenses",
    "Total des frais d'exploitation": "Total Operation expenses",
    "RÉSULTAT D'EXPLOITATION": "OPERATION SURPLUS",
    "Honoraires de gestion": "Percent Management fee",
    "Total des autres frais": "Total other expenses",
    "BÉNÉFICE NET": "NET INCOME",
}

# ============================================================================
# COMBINED ALL LABEL MAPPINGS
# ============================================================================
ALL_LABEL_MAPPINGS = {}
ALL_LABEL_MAPPINGS.update(PAGE10_FRENCH_LABELS)
ALL_LABEL_MAPPINGS.update({
    # English versions
    "Monthly Revenues": "Monthly Revenues",
    "Monthly Revenue": "Monthly Revenues",
    "Daily Revenues": "Transient Revenue",
    "Daily Revenue": "Transient Revenue",
    "Transient Revenue": "Transient Revenue",
    "Car Wash Revenues": "Car-Wash Revenue",
    "Car Wash Revenue": "Car-Wash Revenue",
    "Violation": "Violation",
    "Total Parking Revenue": "Parking Revenue",
    "Gratuities - Monthly": "Discount-Gratuities - Monthly",
    "Discounts - Gratuities (Monthly)": "Discount-Gratuities - Monthly",
    "TOTAL REVENUE": "TOTAL REVENUE",
    "Miscellaneous": "Miscellaneous",
    "Parking Salaries": "Parking wages",
    "Parking Wages": "Parking wages",
    "Parking Supplies": "Parking supplies",
    "Maintenance - Cleaning": "R&M - Cleaning",
    "Maintenance - Equipment": "R&M - Equipement",
    "Maintenance - General": "R&M - General",
    "Repair & Maintenance": "R&M - General",
    "Taxes & Permits": "Tax & license",
    "Insurance & Bonding": "Insurance & Guarantee",
    "Insurance & Guarantee": "Insurance & Guarantee",
    "Claims": "Claims",
    "Telecommunication": "Telecommunication",
    "Credit Card Fees": "Credit Card fees",
    "Bank Fees": "Credit Card fees",
    "Office Expenses": "Office expenses",
    "Ad. & Promotion": "Ad. & Promotion",
    "TOTAL OPERATING EXPENSES": "Total Operation expenses",
    "TOTAL OPERATION EXPENSES": "Total Operation expenses",
    "OPERATING SURPLUS": "OPERATION SURPLUS",
    "OPERATION SURPLUS": "OPERATION SURPLUS",
    "Management Fees": "Percent Management fee",
    "Percent Management Fee": "Percent Management fee",
    "Incentives": "Incentives",
    "TOTAL OTHER EXPENSES": "Total other expenses",
    "NET INCOME": "NET INCOME",
    # Other French
    "Revenus horaires": "Transient Revenue",
    "Revenus quotidiens": "Transient Revenue",
    "Total des revenus": "TOTAL REVENUE",
    "Salaires": "Parking wages",
    "Salaires stationnement": "Parking wages",
    "Fournitures": "Parking supplies",
    "Fournitures stationnements": "Parking supplies",
    "Entretien": "R&M - General",
    "Nettoyage": "R&M - Cleaning",
    "Équipement": "R&M - Equipement",
    "Assurances": "Insurance & Guarantee",
    "Télécommunications": "Telecommunication",
    "Publicité": "Ad. & Promotion",
    "Frais bancaires": "Credit Card fees",
    "Frais de banque & C.C.": "Credit Card fees",
    "Total des dépenses": "Total Operation expenses",
    "Surplus": "OPERATION SURPLUS",
    "Frais de gestion": "Percent Management fee",
    "Incitatifs": "Incentives",
    "Revenu net": "NET INCOME",
    "Location d'équipement": "Equipment rent",
    "Sécurité": "Security",
    "Serv. info. - Général": "Computer services",
    "Honoraires de gestion (base)": "Management Fees (Basic)",
    "Honoraire de gestion a %": "Percent Management fee",
    # Excel monthly
    "Mensuels": "Monthly Revenues",
    "Gratuities - Monthlies": "Discount-Gratuities - Monthly",
    "Gratuites - Mensuels": "Discount-Gratuities - Monthly",
    "Mensuels collectés par le propriétaire": "Monthly Revenues",
    "Monthlies Collected by the Owner": "Monthly Revenues",
    "Autres - Ouverture de dossier": "Other Monthly revenue",
    "Others - Processing Fee": "Other Monthly revenue",
    "Revenus Visiteurs Jours": "Transient Revenue",
    "Transient Revenue - Day": "Transient Revenue",
    "Coin Box & Meter": "Transient Revenue",
    "Revenus Remboursement": "Transient Revenue",
    "Revenues Reimbursement": "Transient Revenue",
    "Gratuities - Transient": "Discount-Gratuities - Transient",
    "Gratuites - Journaliers": "Discount-Gratuities - Transient",
    "Revenus validations": "Transient Revenue",
    "Validation": "Transient Revenue",
    "Revenus hotel": "Hotel Revenue",
    "Hotel Revenues": "Hotel Revenue",
    "Revenus navettes": "Shuttle expenses",
    "Shuttle Revenues": "Shuttle expenses",
    "Lave-Auto": "Car-Wash Revenue",
    "Revenus violation": "Violation",
    "Revenues Violation": "Violation",
    "Frais de cartes perdues": "Miscellaneous",
    "Lost card fees": "Miscellaneous",
    "Intérêts Bancaires": "Other Monthly revenue",
    "Monthly Processing Fees": "Other Monthly revenue",
    "Visiteurs soirs": "Transient Revenue",
    "Evening Tickets": "Transient Revenue",
    "Autres revenus": "Miscellaneous",
    "Others Tickets": "Miscellaneous",
    "Revenus Événement Spécial": "Transient Revenue",
    "Special Events": "Transient Revenue",
    "Fin de semaine": "Transient Revenue",
    "Week end visitors": "Transient Revenue",
    "Réservation en ligne": "Transient Revenue",
    "Online reservation": "Transient Revenue",
    "Salaires - Supervision": "Other wages",
    "Billet de stationnement": "Parking supplies",
    "Entretien réparation - général": "R&M - General",
    "Vehicle Expenses": "Vehicle expenses",
    "Serv,info-Intrnet": "Computer services",
    "Publicité et promotion": "Ad. & Promotion",
})

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

def read_pdf_with_ocr(uploaded_file):
    """Extract text from image-based PDF using OCR."""
    try:
        from pdf2image import convert_from_bytes
        import pytesseract

        file_bytes = get_file_bytes(uploaded_file)
        images = convert_from_bytes(file_bytes, dpi=250)

        sheets = {}
        for page_num, image in enumerate(images):
            try:
                text = pytesseract.image_to_string(image, lang='fra+eng')
            except:
                try:
                    text = pytesseract.image_to_string(image, lang='eng')
                except:
                    continue

            if text and len(text.strip()) > 50:
                lines = text.strip().split('\n')
                lines = [l for l in lines if l.strip()]
                if lines:
                    df = pd.DataFrame(lines, columns=['Text'])
                    sheet_key = f"Page{page_num+1}_OCR"
                    sheets[sheet_key] = df

        return sheets if sheets else None
    except Exception:
        return None

def read_pdf_with_fitz(uploaded_file):
    """
    Extract text from PDF using PyMuPDF (fitz).
    Works well for born-digital PDFs with actual text content.
    Creates sheets named Page1_Fitz, Page2_Fitz, etc.
    """
    try:
        import fitz  # PyMuPDF

        file_bytes = get_file_bytes(uploaded_file)
        doc = fitz.open(stream=file_bytes, filetype="pdf")

        sheets = {}
        total_text_lines = 0

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")

            if text and len(text.strip()) > 50:
                lines = text.strip().split('\n')
                lines = [l.strip() for l in lines if l.strip()]
                if lines:
                    df = pd.DataFrame(lines, columns=['Text'])
                    sheet_key = f"Page{page_num+1}_Fitz"
                    sheets[sheet_key] = df
                    total_text_lines += len(lines)

        doc.close()

        # Only return if we got meaningful data
        if sheets and total_text_lines > 20:
            return sheets

        return None

    except ImportError:
        return None
    except Exception as e:
        return None

def read_pdf_to_dataframe(uploaded_file):
    """
    MAIN PDF READER - FIXED PRIORITY:
    1. Fitz (PyMuPDF) FIRST - best for born-digital PDFs
    2. pdfplumber tables ONLY if Fitz fails
    3. OCR as absolute last resort

    Fitz creates sheets with names like: Page1_Fitz, Page10_Fitz
    These contain clean text that's perfect for Page 10 extraction.
    """

    # STEP 1: Try Fitz first (BEST for born-digital PDFs)
    fitz_sheets = read_pdf_with_fitz(uploaded_file)
    if fitz_sheets:
        return fitz_sheets

    # STEP 2: Fall back to pdfplumber (ONLY if Fitz completely failed)
    try:
        file_bytes = get_file_bytes(uploaded_file)
        sheets = {}

        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if tables:
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

                text = page.extract_text()
                if text and len(text.strip()) > 50:
                    lines = text.strip().split('\n')
                    lines = [l for l in lines if l.strip()]
                    if lines:
                        df = pd.DataFrame(lines, columns=['Text'])
                        sheet_key = f"Page{page_num+1}_Text"
                        sheets[sheet_key] = df

        if sheets:
            return sheets
    except Exception:
        pass

    # STEP 3: Last resort - OCR
    return read_pdf_with_ocr(uploaded_file)

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
    sheets_dict, file_type = read_any_file_to_dataframes(pnl_file)
    codes = []
    if sheets_dict:
        for sheet_name in sheets_dict.keys():
            match = re.search(r'(CMO\d+)', sheet_name, re.IGNORECASE)
            if match:
                codes.append(match.group(1).upper())
            if 'LUNA' in sheet_name.upper():
                codes.append('LUNA')
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

def label_match_score(cell_text, label_text):
    cell_clean = clean_text_for_matching(cell_text)
    label_clean = clean_text_for_matching(label_text)
    if not cell_clean or not label_clean:
        return 0
    if cell_clean == label_clean:
        return 1.0
    if label_clean in cell_clean:
        ratio = len(label_clean) / len(cell_clean)
        if ratio >= 0.5:
            return ratio
        return 0
    if cell_clean in label_clean:
        ratio = len(cell_clean) / len(label_clean)
        if ratio >= 0.5:
            return ratio
        return 0
    return 0

def extract_dollar_amount_from_text(text):
    """Extract a dollar amount from any text. Returns the value or None."""
    if not text:
        return None

    # Try various patterns
    patterns = [
        r'\$([\d,]+\.?\d*)',            # $1,234.56
        r'([\d,]+\.?\d*)\s*\$',         # 1,234.56 $
        r'\$?\s*([\d,]+\.\d{2})\s*\$?', # 1234.56 with 2 decimal places
        r'\(?\$?([\d,]+\.?\d*)\)?',     # (1,234.56)
    ]

    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            try:
                return float(match.group(1).replace(',', ''))
            except:
                continue

    # Try any number that looks like currency (with 2 decimal places)
    match = re.search(r'([\d,]+\.\d{2})', text)
    if match:
        try:
            return float(match.group(1).replace(',', ''))
        except:
            pass

    return None

def parse_text_line_for_data(line):
    """
    Parse a text line for label and dollar amount.
    Uses EXACT French labels from Page 10 for better matching.
    """
    if not line or len(line) < 5:
        return None, None

    line = line.strip()

    # Extract dollar amount
    val = extract_dollar_amount_from_text(line)

    # Try to match EXACT French labels first (Page 10 format)
    for mapping_label, standard_label in PAGE10_FRENCH_LABELS.items():
        if label_match_score(line, mapping_label) >= 0.6:
            return standard_label, val

    # Try all other labels
    for mapping_label, standard_label in ALL_LABEL_MAPPINGS.items():
        if mapping_label in PAGE10_FRENCH_LABELS:
            continue  # Already checked
        if label_match_score(line, mapping_label) >= 0.6:
            return standard_label, val

    # If value found but no label, try to find label by removing the amount
    if val is not None:
        label_text = re.sub(r'\$?[\d,]+\.?\d*\s*\$?', '', line).strip()
        label_text = re.sub(r'\(?[\d,]+\.?\d*\)?', '', label_text).strip()
        label_text = re.sub(r'\s+', ' ', label_text).strip()

        if len(label_text) > 2:
            # Try Page 10 labels first
            for mapping_label, standard_label in PAGE10_FRENCH_LABELS.items():
                if label_match_score(label_text, mapping_label) >= 0.6:
                    return standard_label, val
            # Then all others
            for mapping_label, standard_label in ALL_LABEL_MAPPINGS.items():
                if mapping_label in PAGE10_FRENCH_LABELS:
                    continue
                if label_match_score(label_text, mapping_label) >= 0.6:
                    return standard_label, val

    return None, val

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

def find_best_data_sheet(sheets_dict):
    """
    Find the best sheet for financial data extraction.

    PRIORITY ORDER:
    1. Sheets with "Fitz" in name (best quality text from born-digital PDFs)
    2. Sheets that look like Page 10 (have Page 10 account names)
    3. Sheets with most financial data (dollar amounts + labels)
    """
    if not sheets_dict:
        return None

    # PRIORITY 1: Fitz sheets with Page 10 structure
    fitz_sheets = [name for name in sheets_dict.keys() if 'Fitz' in name]

    if fitz_sheets:
        # Check each Fitz sheet for Page 10 indicators
        for sheet_name in fitz_sheets:
            df = sheets_dict[sheet_name]
            if df is not None and len(df) > 20 and 'Text' in df.columns:
                # Look for Page 10 keywords
                page10_keywords = ['Revenus mensuels', 'TOTAL REVENUS',
                                  'DÉPENSES', 'RÉSULTAT', 'BÉNÉFICE NET']
                matches = 0
                for row_idx in range(min(30, len(df))):
                    try:
                        text = str(df.iloc[row_idx, 0]).upper()
                        for kw in page10_keywords:
                            if clean_text_for_matching(kw) in clean_text_for_matching(text):
                                matches += 1
                    except:
                        continue

                if matches >= 3:  # This is likely Page 10
                    return sheet_name

        # If no Page 10 found, pick the Fitz sheet with most text lines
        best_fitz = None
        max_lines = 0
        for sheet_name in fitz_sheets:
            df = sheets_dict[sheet_name]
            if df is not None and len(df) > max_lines:
                max_lines = len(df)
                best_fitz = sheet_name

        if best_fitz:
            return best_fitz

    # PRIORITY 2: Any sheet with Page 10 structure
    for sheet_name, df in sheets_dict.items():
        if df is not None and len(df) > 20:
            page10_keywords = ['Revenus mensuels', 'TOTAL REVENUS',
                              'DÉPENSES', 'RÉSULTAT', 'BÉNÉFICE NET']
            matches = 0
            for row_idx in range(min(30, len(df))):
                try:
                    text = str(df.iloc[row_idx, 0]).upper()
                    for kw in page10_keywords:
                        if clean_text_for_matching(kw) in clean_text_for_matching(text):
                            matches += 1
                except:
                    continue
            if matches >= 3:
                return sheet_name

    # PRIORITY 3: Fall back to scoring system for other formats
    candidates = []
    for sheet_name, df in sheets_dict.items():
        if df is None or len(df) == 0:
            continue

        text_cells = 0
        numeric_cells = 0
        dollar_cells = 0

        for row_idx in range(min(40, len(df))):
            for col_idx in range(min(10, len(df.columns))):
                try:
                    cell_text = str(df.iloc[row_idx, col_idx]).strip()
                    if cell_text and cell_text.lower() != 'nan' and len(cell_text) > 2:
                        text_cells += 1

                    val = safe_float(df.iloc[row_idx, col_idx])
                    if val != 0:
                        numeric_cells += 1
                        if '$' in cell_text:
                            dollar_cells += 1
                except:
                    pass

        score = numeric_cells * 10 + dollar_cells * 20 + text_cells
        if text_cells > 3:
            candidates.append((sheet_name, score))

    if candidates:
        candidates.sort(key=lambda x: x[1], reverse=True)
        return candidates[0][0]

    # Last resort: any sheet with data
    for sheet_name, df in sheets_dict.items():
        if df is not None and len(df) > 0:
            return sheet_name

    return None

def find_amount_column(df):
    """Find the column with monthly actual values."""
    if df is None or len(df) == 0:
        return None

    if 'Text' in df.columns:
        return None

    for row_idx in range(min(5, len(df))):
        for col_idx in range(min(15, len(df.columns))):
            try:
                cell_text = str(df.iloc[row_idx, col_idx]).lower().strip()
                if any(term in cell_text for term in [
                    'current month actual', 'mois courant', 'courant',
                    'amount', 'navision', 'actual'
                ]):
                    if 'ytd' not in cell_text and 'previous' not in cell_text and 'budget' not in cell_text:
                        return col_idx
            except:
                continue

    best_col = None
    best_count = 0
    for col_idx in range(1, min(8, len(df.columns))):
        numeric_count = 0
        for row_idx in range(min(30, len(df))):
            val = safe_float(df.iloc[row_idx, col_idx])
            if val != 0:
                numeric_count += 1
        if numeric_count > best_count:
            best_count = numeric_count
            best_col = col_idx

    if best_count >= 3:
        return best_col

    return 2

def extract_data_from_text_sheet(df, month_name, year):
    """
    Extract data from OCR/Text sheet.
    Handles labels and values on separate lines.
    Uses exact Page 10 French labels for better matching.
    """
    result = {}

    # Collect all non-empty lines
    all_lines = []
    for row_idx in range(len(df)):
        try:
            if 'Text' in df.columns:
                line = str(df.iloc[row_idx, 0]).strip()
            else:
                parts = []
                for col_idx in range(min(10, len(df.columns))):
                    cell = str(df.iloc[row_idx, col_idx]).strip()
                    if cell and cell.lower() != 'nan' and cell.lower() != 'none':
                        parts.append(cell)
                line = ' '.join(parts)

            if line and len(line) > 2:
                all_lines.append(line)
        except:
            continue

    # Process each line
    for i, line in enumerate(all_lines):
        line_upper = line.upper()

        # Check for TOTAL REVENUS
        if 'TOTAL REVENUS' in line_upper:
            val = extract_dollar_amount_from_text(line)
            if not val and i + 1 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+1])
            if not val and i + 2 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+2])
            if val and val != 0:
                result['_REVENUE_TOTAL_'] = val
            continue

        # Check for TOTAL DES FRAIS / TOTAL OPERATING
        if 'TOTAL DES FRAIS' in line_upper or 'TOTAL OPERATING' in line_upper or 'TOTAL OPERATION' in line_upper:
            val = extract_dollar_amount_from_text(line)
            if not val and i + 1 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+1])
            if not val and i + 2 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+2])
            if val and val != 0:
                result['_EXPENSE_TOTAL_'] = val
            continue

        # Try to match label + value
        std_label, val = parse_text_line_for_data(line)

        # If label found but no value, check next 2 lines
        if std_label and not val:
            if i + 1 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+1])
            if not val and i + 2 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+2])

        # If value found but no label, check previous 2 lines
        if not std_label and val:
            if i > 0:
                std_label, _ = parse_text_line_for_data(all_lines[i-1])
            if not std_label and i > 1:
                std_label, _ = parse_text_line_for_data(all_lines[i-2])

        if std_label and val and val != 0:
            if std_label in result:
                result[std_label] += val
            else:
                result[std_label] = val

    result['_DEBUG_MATCHES_'] = str(len([k for k in result.keys() if not k.startswith('_')]))
    return result

# ============================================================================
# PAGE 10 - PRECISE POSITION-BASED EXTRACTION
# ============================================================================

def extract_page10_from_fitz_text(df):
    """
    Extract data from Page 10 using the known spatial layout.

    Fitz gives us clean text lines. Page 10 has a specific structure:
    - Account names appear in order
    - Numbers follow the account names

    Returns dict with standard English labels mapped to their values.
    """
    result = {}

    # Get all text lines
    text_lines = []
    for row_idx in range(len(df)):
        try:
            if 'Text' in df.columns:
                line = str(df.iloc[row_idx, 0]).strip()
            else:
                parts = [str(df.iloc[row_idx, c]).strip()
                        for c in range(min(10, len(df.columns)))
                        if str(df.iloc[row_idx, c]).strip().lower() not in ['nan', 'none', '']]
                line = ' '.join(parts)

            if line and len(line) > 2:
                text_lines.append(line)
        except:
            continue

    # Page 10 exact account names in order (French with English mapping)
    page10_accounts = [
        ("Monthly Revenues", ["Revenus mensuels", "Mensuels"]),
        ("Transient Revenue", ["Revenus Journaliers", "Journaliers", "Revenus horaires", "Revenus quotidiens"]),
        ("Car-Wash Revenue", ["Revenus Lave-Auto", "Lave-Auto"]),
        ("Miscellaneous", ["Divers"]),
        ("Parking Revenue", ["Revenus de stationnement"]),
        ("Discount-Gratuities - Monthly", ["Gratuités - mensuels", "Gratuites - mensuels", "Gratuités"]),
        ("TOTAL REVENUE", ["TOTAL REVENUS", "TOTAL DES REVENUS"]),
        (None, ["DÉPENSES", "DEPENSES"]),  # Section header - skip
        (None, ["DÉPENSES D'EXPLOITATION", "DEPENSES D'EXPLOITATION"]),  # Sub-header - skip
        ("Parking wages", ["Salaires Stationnement", "Salaires stationnement"]),
        ("Uniforms", ["Uniformes"]),
        ("Parking supplies", ["Fourn. de stationnement", "Fournitures stationnement"]),
        ("R&M - Cleaning", ["Entretien réparation - Nettoyage", "Nettoyage"]),
        ("R&M - Equipement", ["Entretien réparation - Equipement", "Équipement", "Equipement"]),
        ("R&M - General", ["Entretien réparation - Général", "Général", "General"]),
        ("Tax & license", ["Taxes et permis"]),
        ("Insurance & Guarantee", ["Assurances Cautionnement", "Assurances"]),
        ("Claims", ["Réclamations", "Reclamations"]),
        ("Telecommunication", ["Télécommunication", "Telecommunication"]),
        ("Credit Card fees", ["Frais de cartes de crédit", "Frais de cartes de credit"]),
        ("Office expenses", ["Frais de bureau"]),
        ("Total Operation expenses", ["Total des frais d'exploitation"]),
        ("OPERATION SURPLUS", ["RÉSULTAT D'EXPLOITATION", "RESULTAT D'EXPLOITATION"]),
        (None, ["AUTRES FRAIS"]),  # Section header - skip
        ("Percent Management fee", ["Honoraires de gestion"]),
        ("Total other expenses", ["Total des autres frais"]),
        ("NET INCOME", ["BÉNÉFICE NET", "BENEFICE NET"]),
    ]

    account_index = 0
    found_accounts = []

    for line in text_lines:
        line_upper = line.upper().strip()

        if account_index >= len(page10_accounts):
            break

        english_name, french_names = page10_accounts[account_index]

        # Check if this line contains one of the French names
        matched = False
        matched_french = None
        for french_name in french_names:
            if clean_text_for_matching(french_name) in clean_text_for_matching(line_upper):
                matched = True
                matched_french = french_name
                break

        if not matched:
            # Maybe the text is split across lines? Check next few lines
            for offset in [1, 2]:
                if account_index + offset < len(page10_accounts):
                    next_english, next_french_names = page10_accounts[account_index + offset]
                    for next_french in next_french_names:
                        if clean_text_for_matching(next_french) in clean_text_for_matching(line_upper):
                            # We found a later account, mark current as missing
                            account_index += offset
                            matched = True
                            matched_french = next_french
                            english_name = next_english
                            break
                    if matched:
                        break
            if not matched:
                continue

        if english_name is None:
            # Section header - skip
            account_index += 1
            continue

        # Extract the dollar amount from this line
        val = extract_dollar_amount_from_text(line)

        # If no amount on this line, check next line
        if val is None:
            line_idx = text_lines.index(line)
            if line_idx + 1 < len(text_lines):
                val = extract_dollar_amount_from_text(text_lines[line_idx + 1])
            if val is None and line_idx + 2 < len(text_lines):
                val = extract_dollar_amount_from_text(text_lines[line_idx + 2])

        if val is not None and val != 0:
            result[english_name] = val
            found_accounts.append(f"{english_name}: ${val:,.2f}")

        account_index += 1

    result['_DEBUG_MATCHES_'] = str(len(found_accounts))
    result['_DEBUG_FOUND_'] = '; '.join(found_accounts[:10])  # First 10 for debug

    return result

def extract_monthly_data_from_file(uploaded_file):
    """
    Enhanced to detect Page 10 and use position-based extraction.
    Falls back to label matching for other formats.
    """
    result = {}
    month_name = None
    year = None

    sheets_dict, file_type = read_any_file_to_dataframes(uploaded_file)

    result['_DEBUG_TYPE_'] = str(file_type)

    if sheets_dict is None:
        result['_DEBUG_ERROR_'] = "No sheets found"
        return result, (None, None)

    available_sheets = list(sheets_dict.keys())
    result['_DEBUG_SHEETS_'] = str(available_sheets)[:200]

    target_sheet = find_best_data_sheet(sheets_dict)
    result['_DEBUG_TARGET_'] = str(target_sheet) if target_sheet else "None"

    if target_sheet is None:
        result['_DEBUG_ERROR_'] = "No suitable sheet found"
        return result, (None, None)

    df = sheets_dict[target_sheet]
    if df is None or len(df) == 0:
        result['_DEBUG_ERROR_'] = "Sheet has no data"
        return result, (None, None)

    result['_DEBUG_ROWS_'] = str(len(df))
    result['_DEBUG_COLS_'] = str(len(df.columns))

    # Extract filename info for month/year detection
    if hasattr(uploaded_file, 'name'):
        month_name, year = extract_month_year_from_text(uploaded_file.name)

    # DETECT: Is this a Page 10 Fitz sheet?
    is_fitz = 'Fitz' in target_sheet
    is_page10 = False

    if is_fitz and 'Text' in df.columns:
        # Check for Page 10 structure
        page10_indicators = 0
        for row_idx in range(min(30, len(df))):
            try:
                text = str(df.iloc[row_idx, 0]).upper()
                if 'REVENUS MENSUELS' in text:
                    page10_indicators += 1
                if 'TOTAL REVENUS' in text:
                    page10_indicators += 1
                if 'DÉPENSES' in text or 'DEPENSES' in text:
                    page10_indicators += 1
                if 'RÉSULTAT' in text or 'RESULTAT' in text:
                    page10_indicators += 1
                if 'BÉNÉFICE NET' in text or 'BENEFICE NET' in text:
                    page10_indicators += 1
            except:
                continue

        is_page10 = page10_indicators >= 3
        result['_DEBUG_PAGE10_'] = f"Indicators: {page10_indicators}, is_page10: {is_page10}"

    # EXTRACT using the appropriate method
    if is_page10 and 'Text' in df.columns:
        # Use precise Page 10 position-based extraction
        result['_DEBUG_METHOD_'] = "Page10_Fitz_position_based"
        data = extract_page10_from_fitz_text(df)

        for key, value in data.items():
            if not key.startswith('_'):
                result[key] = value

        # Set totals for DH balancing
        if 'TOTAL REVENUE' in data:
            result['_REVENUE_TOTAL_'] = data['TOTAL REVENUE']
        if 'Total Operation expenses' in data:
            result['_EXPENSE_TOTAL_'] = data['Total Operation expenses']

        # Copy debug info
        for key in ['_DEBUG_MATCHES_', '_DEBUG_FOUND_']:
            if key in data:
                result[key] = data[key]

    elif 'Text' in df.columns or target_sheet.endswith('_OCR') or target_sheet.endswith('_Text'):
        # Text-based extraction (existing logic)
        result['_DEBUG_METHOD_'] = "text_based"
        data = extract_data_from_text_sheet(df, month_name, year)
        for key, value in data.items():
            result[key] = value

    else:
        # Table-based extraction (existing logic for Excel/CSV)
        result['_DEBUG_METHOD_'] = "table_based"
        amount_col = find_amount_column(df)
        result['_DEBUG_AMOUNT_COL_'] = str(amount_col) if amount_col is not None else "None"

        if amount_col is None:
            amount_col = 2

        matches_found = 0

        for row_idx in range(len(df)):
            try:
                for col_idx in range(min(10, len(df.columns))):
                    cell_text = str(df.iloc[row_idx, col_idx]).strip().upper()

                    if 'TOTAL REVENUS' in cell_text or 'TOTAL REVENUE' in cell_text:
                        val = safe_float(df.iloc[row_idx, amount_col])
                        if val != 0:
                            result['_REVENUE_TOTAL_'] = val
                        break

                    if 'TOTAL DES FRAIS' in cell_text or 'TOTAL OPERATING' in cell_text:
                        val = safe_float(df.iloc[row_idx, amount_col])
                        if val != 0:
                            result['_EXPENSE_TOTAL_'] = val
                        break

                for col_idx in range(min(10, len(df.columns))):
                    cell_text = str(df.iloc[row_idx, col_idx]).strip()
                    if not cell_text or len(cell_text) < 3:
                        continue

                    for label, standard in ALL_LABEL_MAPPINGS.items():
                        if label_match_score(cell_text, label) >= 0.6:
                            val = safe_float(df.iloc[row_idx, amount_col])
                            if val != 0:
                                matches_found += 1
                                if standard in result:
                                    result[standard] += val
                                else:
                                    result[standard] = val
                            break
            except Exception:
                continue

        result['_DEBUG_MATCHES_'] = str(matches_found)

    return result, (month_name, year)

def build_monthly_data_from_files(monthly_files):
    if not monthly_files:
        return None
    monthly_data = {}
    yearly_data = {}
    monthly_totals = {}

    for uploaded_file in monthly_files:
        file_data, (month_name, year) = extract_monthly_data_from_file(uploaded_file)

        debug_info = {
            'file': getattr(uploaded_file, 'name', 'unknown'),
            'month': month_name,
            'year': year,
            'type': file_data.pop('_DEBUG_TYPE_', '?'),
            'sheets': file_data.pop('_DEBUG_SHEETS_', '?'),
            'target': file_data.pop('_DEBUG_TARGET_', '?'),
            'rows': file_data.pop('_DEBUG_ROWS_', '?'),
            'cols': file_data.pop('_DEBUG_COLS_', '?'),
            'sample': file_data.pop('_DEBUG_SAMPLE_', '?'),
            'matches': file_data.pop('_DEBUG_MATCHES_', '?'),
            'amount_col': file_data.pop('_DEBUG_AMOUNT_COL_', '?'),
            'method': file_data.pop('_DEBUG_METHOD_', '?'),
            'page10': file_data.pop('_DEBUG_PAGE10_', '?'),
            'found': file_data.pop('_DEBUG_FOUND_', '?'),
            'error': file_data.pop('_DEBUG_ERROR_', None),
        }

        if not monthly_data:
            monthly_data['_debug_info'] = []
        if '_debug_info' not in monthly_data:
            monthly_data['_debug_info'] = []
        monthly_data['_debug_info'].append(debug_info)

        if not file_data or month_name is None:
            continue

        revenue_total = file_data.pop('_REVENUE_TOTAL_', None)
        expense_total = file_data.pop('_EXPENSE_TOTAL_', None)

        if revenue_total is not None or expense_total is not None:
            monthly_totals[month_name] = {
                "revenue_total": revenue_total,
                "expense_total": expense_total
            }

        for label, value in file_data.items():
            if label.startswith('_') and label.endswith('_'):
                continue
            if label not in monthly_data:
                monthly_data[label] = {}
            monthly_data[label][month_name] = value
            if label not in yearly_data:
                yearly_data[label] = 0
            yearly_data[label] += value

    debug_list = monthly_data.pop('_debug_info', None)

    if not monthly_data:
        if debug_list:
            return {'monthly': {}, 'yearly': {}, '_debug_info': debug_list}
        return None

    result = {
        'monthly': monthly_data,
        'yearly': yearly_data
    }
    if monthly_totals:
        result['_monthly_totals'] = monthly_totals
    if debug_list:
        result['_debug_info'] = debug_list

    return result

# ============================================================================
# PAGE 3/10 FINANCIAL SUMMARY EXTRACTION (for yearly data)
# ============================================================================

def find_ytd_column(df):
    """Find the YTD Actual column."""
    if df is None or len(df) == 0:
        return None

    if 'Text' in df.columns:
        return None

    for row_idx in range(min(5, len(df))):
        for col_idx in range(min(15, len(df.columns))):
            try:
                cell_text = str(df.iloc[row_idx, col_idx]).lower().strip()
                if any(term in cell_text for term in ['ytd actual', 'cumulatif courant', 'ytd', 'cumulatif', 'year to date']):
                    return col_idx
            except:
                continue

    for col_idx in [5, 6, 7]:
        if col_idx < len(df.columns):
            has_numbers = False
            for row_idx in range(min(20, len(df))):
                if safe_float(df.iloc[row_idx, col_idx]) != 0:
                    has_numbers = True
                    break
            if has_numbers:
                return col_idx

    return None

def extract_page3_data(uploaded_file):
    """Extract YTD Actual data from Page 3/10 Financial Summary."""
    result = {'monthly': {}, 'yearly': {}}

    sheets_dict, file_type = read_any_file_to_dataframes(uploaded_file)
    if sheets_dict is None:
        return None

    target_sheet = find_best_data_sheet(sheets_dict)
    if target_sheet is None:
        return None

    df = sheets_dict[target_sheet]
    if df is None or len(df) == 0:
        return None

    if 'Text' in df.columns:
        data = extract_data_from_text_sheet(df, None, None)
        for key, value in data.items():
            if not key.startswith('_'):
                result['yearly'][key] = value
    else:
        ytd_col = find_ytd_column(df)
        if ytd_col is None:
            return None

        for row_idx in range(len(df)):
            try:
                for col_idx in range(min(10, len(df.columns))):
                    cell_text = str(df.iloc[row_idx, col_idx]).strip()
                    if not cell_text or len(cell_text) < 3:
                        continue

                    for label, standard in ALL_LABEL_MAPPINGS.items():
                        if label_match_score(cell_text, label) >= 0.6:
                            val = safe_float(df.iloc[row_idx, ytd_col])
                            if val != 0:
                                result['yearly'][standard] = val
                            break
            except:
                continue

    if not result['yearly']:
        return None

    return result

# ============================================================================
# P&L DATA EXTRACTION (fallback)
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
                shorter = min(len(alt_clean), len(key_clean))
                longer = max(len(alt_clean), len(key_clean))
                if shorter >= 5 and (shorter / longer) >= 0.6:
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
                if shorter >= 5 and (shorter / longer) >= 0.6:
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

def update_budget_initial(wb, bi_data, parking_code):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["❌ Budget Initial: Sheet not found"]
        ws = wb[sheet_name]

        if bi_data is None:
            updates.append("⚠️ Budget Initial: No data available")
            return updates

        cells_updated = 0
        filled_revenue = 0
        filled_expense = 0

        for dh_row, pnl_labels in DH_ROW_MAPPING.items():
            yearly_value = find_pnl_value(bi_data, pnl_labels)

            if yearly_value != 0:
                ws[f"S{dh_row}"] = yearly_value
                ws[f"S{dh_row}"].number_format = '#,##0.00 $'
                cells_updated += 1
                updates.append(f"✅ BI Row {dh_row}: ${yearly_value:,.2f} ({pnl_labels[0]})")

                if dh_row in REVENUE_ROWS:
                    filled_revenue += yearly_value
                elif dh_row in EXPENSE_ROWS:
                    filled_expense += yearly_value

        expected_net = find_pnl_value(bi_data, [
            "NET INCOME", "net income", "revenus nets", "REVENUS NETS", "BÉNÉFICE NET"
        ])

        if expected_net != 0:
            calculated_net = filled_revenue - filled_expense
            gap = expected_net - calculated_net

            updates.append(f"🔍 BI Validation: Expected NET INCOME = ${expected_net:,.2f}")
            updates.append(f"🔍 BI: Filled Revenue = ${filled_revenue:,.2f}, Filled Expense = ${filled_expense:,.2f}")
            updates.append(f"🔍 BI: Calculated Net = ${calculated_net:,.2f}, Gap = ${gap:,.2f}")

        if cells_updated > 0:
            updates.append(f"✅ Budget Initial: {cells_updated} cells updated")
        else:
            updates.append("⚠️ Budget Initial: No cells updated")

    except Exception as e:
        updates.append(f"❌ Budget Initial: {str(e)}")
    return updates

def update_fiche_stationnement(wb, fs_data, parking_code, word_data=None):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["❌ Fiche Stationnement: Sheet not found"]
        ws = wb[sheet_name]
        if fs_data is None:
            updates.append("⚠️ Fiche Stationnement: No data available")
            return updates
        for cell, pnl_labels in FICHE_STATIONNEMENT_MAP:
            yearly_value = find_pnl_value(fs_data, pnl_labels)
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00 $'
                updates.append(f"✅ {cell} = ${yearly_value:,.2f}")
            else:
                updates.append(f"⚠️ {cell} = Not found (tried: {pnl_labels[0]})")
    except Exception as e:
        updates.append(f"❌ Fiche Stationnement: {str(e)}")
    return updates

def update_donnees_historiques(wb, merged_monthly_data, parking_code, monthly_totals=None):
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

        monthly_filled_revenue = {}
        monthly_filled_expense = {}

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
                        ws_dh[cell_ref].number_format = '#,##0.00 $'
                        cells_updated += 1
                        row_cells += 1

                        if dh_row in REVENUE_ROWS:
                            if month_name not in monthly_filled_revenue:
                                monthly_filled_revenue[month_name] = 0
                            monthly_filled_revenue[month_name] += val
                        elif dh_row in EXPENSE_ROWS:
                            if month_name not in monthly_filled_expense:
                                monthly_filled_expense[month_name] = 0
                            monthly_filled_expense[month_name] += val

            if row_cells > 0:
                rows_filled.append(f"  Row {dh_row}: {pnl_labels[0]} ({row_cells} months)")

        if monthly_totals:
            balancing_updates = []
            for month_name, totals in monthly_totals.items():
                if month_name not in MONTHS_EN:
                    continue

                month_idx = MONTHS_EN.index(month_name)
                col_letter = get_column_letter(month_idx + 2)

                expected_revenue = totals.get("revenue_total")
                expected_expense = totals.get("expense_total")

                if expected_revenue is not None and expected_revenue != 0:
                    actual_revenue = monthly_filled_revenue.get(month_name, 0)
                    revenue_gap = expected_revenue - actual_revenue

                    if abs(revenue_gap) > 0.99:
                        catch_row = REVENUE_CATCH_ALL_ROW
                        cell_ref = f"{col_letter}{catch_row}"
                        current_val = safe_float(ws_dh[cell_ref].value)
                        ws_dh[cell_ref] = current_val + revenue_gap
                        ws_dh[cell_ref].number_format = '#,##0.00 $'
                        balancing_updates.append(
                            f"  ⚖️ {month_name}: Added ${revenue_gap:,.2f} to Row {catch_row}"
                        )
                        cells_updated += 1

                if expected_expense is not None and expected_expense != 0:
                    actual_expense = monthly_filled_expense.get(month_name, 0)
                    expense_gap = expected_expense - actual_expense

                    if abs(expense_gap) > 0.99:
                        catch_row = EXPENSE_CATCH_ALL_ROW
                        cell_ref = f"{col_letter}{catch_row}"
                        current_val = safe_float(ws_dh[cell_ref].value)
                        ws_dh[cell_ref] = current_val + expense_gap
                        ws_dh[cell_ref].number_format = '#,##0.00 $'
                        balancing_updates.append(
                            f"  ⚖️ {month_name}: Added ${expense_gap:,.2f} to Row {catch_row}"
                        )
                        cells_updated += 1

            if balancing_updates:
                updates.append(f"⚖️ Balancing ({len(balancing_updates)} adjustments):")
                for bu in balancing_updates:
                    updates.append(bu)

        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: {cells_updated} cells in {len(rows_filled)} rows")
            for row_info in rows_filled:
                updates.append(row_info)
        else:
            updates.append("⚠️ Donnees Historiques: No cells updated")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {str(e)}")
    return updates


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(
    excel_file,
    monthly_files_current=None,
    monthly_files_previous=None,
    budget_initial_file=None,
    fiche_stationnement_file=None,
    parking_code=None,
    word_data=None
):
    updates = []

    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)

    if not parking_code:
        return None, ["❌ Could not determine parking code. Please select a parking code."]

    updates.append(f"🔍 Processing: {parking_code}")

    dh_current_year_data = None
    dh_previous_year_data = None
    monthly_totals = None

    if monthly_files_current and len(monthly_files_current) > 0:
        updates.append("📋 Processing Current Year monthly files")
        dh_current_year_data = build_monthly_data_from_files(monthly_files_current)
        if dh_current_year_data:
            debug_info = dh_current_year_data.pop('_debug_info', None)
            if debug_info:
                for d in debug_info:
                    if d.get('error'):
                        updates.append(f"🔧 {d['file']}: ERROR - {d['error']}")
                    else:
                        updates.append(f"🔧 {d['file']}: type={d['type']}, target={d['target']}, rows={d['rows']}x{d['cols']}, method={d.get('method','?')}, page10={d.get('page10','?')}, matches={d['matches']}, month={d['month']}")
                        if d.get('found'):
                            updates.append(f"   Found: {d['found']}")

            num_labels = len(dh_current_year_data.get('yearly', {}))
            updates.append(f"📊 Current year: {num_labels} labels")
            if '_monthly_totals' in dh_current_year_data:
                monthly_totals = dh_current_year_data.pop('_monthly_totals')
                updates.append(f"📊 Monthly totals for DH: {len(monthly_totals)} months")

    if monthly_files_previous and len(monthly_files_previous) > 0:
        updates.append("📋 Processing Previous Year monthly files")
        dh_previous_year_data = build_monthly_data_from_files(monthly_files_previous)
        if dh_previous_year_data:
            debug_info = dh_previous_year_data.pop('_debug_info', None)
            if debug_info:
                for d in debug_info:
                    if d.get('error'):
                        updates.append(f"🔧 {d['file']}: ERROR - {d['error']}")
                    else:
                        updates.append(f"🔧 {d['file']}: type={d['type']}, target={d['target']}, method={d.get('method','?')}, matches={d['matches']}")

            num_labels = len(dh_previous_year_data.get('yearly', {}))
            updates.append(f"📊 Previous year: {num_labels} labels")
            if '_monthly_totals' in dh_previous_year_data:
                prev_totals = dh_previous_year_data.pop('_monthly_totals')
                if monthly_totals is None:
                    monthly_totals = {}
                monthly_totals.update(prev_totals)

    bi_data = None
    if budget_initial_file:
        updates.append("📋 Processing Budget Initial source")
        bi_data = extract_page3_data(budget_initial_file)
        if bi_data:
            num_labels = len(bi_data.get('yearly', {}))
            updates.append(f"📊 Budget Initial: {num_labels} labels")
        if bi_data is None or (bi_data and len(bi_data.get('yearly', {})) == 0):
            bi_data, _ = extract_pnl_data(budget_initial_file, parking_code)
            if bi_data:
                updates.append(f"📊 Budget Initial (P&L fallback): {len(bi_data.get('yearly', {}))} labels")

    fs_data = None
    if fiche_stationnement_file:
        updates.append("📋 Processing Fiche Stationnement source")
        fs_data = extract_page3_data(fiche_stationnement_file)
        if fs_data:
            num_labels = len(fs_data.get('yearly', {}))
            updates.append(f"📊 Fiche Stationnement: {num_labels} labels")
        if fs_data is None or (fs_data and len(fs_data.get('yearly', {})) == 0):
            fs_data, _ = extract_pnl_data(fiche_stationnement_file, parking_code)
            if fs_data:
                updates.append(f"📊 Fiche Stationnement (P&L fallback): {len(fs_data.get('yearly', {}))} labels")

    if dh_current_year_data is None and dh_previous_year_data is None:
        updates.append("⚠️ No monthly data available for Donnees Historiques")

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
    if year_map and dh_current_year_data and dh_previous_year_data:
        merged_monthly = merge_monthly_data(dh_current_year_data, dh_previous_year_data, year_map)

    if not merged_monthly and dh_current_year_data:
        merged_monthly = dh_current_year_data['monthly']

    dh_data = merged_monthly if merged_monthly else {}
    if not dh_data and dh_current_year_data:
        dh_data = dh_current_year_data['monthly']

    updates.extend(update_budget_initial(wb_write, bi_data, parking_code))
    updates.extend(update_fiche_stationnement(wb_write, fs_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb_write, dh_data, parking_code, monthly_totals))

    success_count = sum(1 for u in updates if u.startswith("✅"))
    if success_count == 0:
        updates.append("💡 No updates were made.")

    output = io.BytesIO()
    wb_write.save(output)
    output.seek(0)

    return output, updates
