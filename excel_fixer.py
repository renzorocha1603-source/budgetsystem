import io
import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
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
    "Revenus mensuels": "Monthly Revenues",
    "Revenus Journaliers": "Transient Revenue",
    "Revenus Lave-Auto": "Car-Wash Revenue",
    "Divers": "Miscellaneous",
    "Revenus de stationnement": "Parking Revenue",
    "(Gratuités - mensuels)": "Discount-Gratuities - Monthly",
    "TOTAL REVENUS": "TOTAL REVENUE",
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

PAGE10_EXCLUDE_KEYWORDS = [
    'CONCILIATION BI', 'ÉCARTS AU BUDGET', 'ECARTS AU BUDGET',
    'SECTION 1', 'SECTION 2', 'SECTION 3', 'SECTION 4',
    'MANUAL BILLING', 'FAITS SAILLANTS',
    'EXPLICATION DES ÉCARTS', 'AJUSTEMENT DÉPÔT',
    'CF. EXTRAIT BI', 'AVANT TAXES', "FRAIS D'OUVERTURE",
    'COMMENTAIRES', 'ANALYSE', 'SOMMAIRE',
    'MENSUELS TOTALES', 'JOURNALIERS TOTALES', 'DÉPENSES TOTALES',
]

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
    """Extract text from image-based PDF using OCR (fallback for January)."""
    try:
        from pdf2image import convert_from_bytes
        import pytesseract
        file_bytes = get_file_bytes(uploaded_file)
        images = convert_from_bytes(file_bytes, dpi=300)
        sheets = {}
        for page_num, image in enumerate(images):
            try:
                text = pytesseract.image_to_string(image, lang='fra+eng', config='--psm 6')
            except:
                try:
                    text = pytesseract.image_to_string(image, lang='eng', config='--psm 6')
                except:
                    continue
            if text and len(text.strip()) > 30:
                lines = text.strip().split('\n')
                lines = [l for l in lines if l.strip()]
                if lines:
                    df = pd.DataFrame(lines, columns=['Text'])
                    # Prefix with OCR_ so downstream code knows the origin
                    sheet_key = f"OCR_Page{page_num+1}"
                    sheets[sheet_key] = df
        return sheets if sheets else None
    except Exception:
        return None

def read_pdf_with_fitz(uploaded_file):
    """Extract text from PDF using PyMuPDF."""
    try:
        import fitz
        file_bytes = get_file_bytes(uploaded_file)
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        sheets = {}
        total_text_lines = 0
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")
            if text and len(text.strip()) > 30:
                lines = text.strip().split('\n')
                lines = [l.strip() for l in lines if l.strip()]
                if lines:
                    df = pd.DataFrame(lines, columns=['Text'])
                    sheet_key = f"Page{page_num+1}_Fitz"
                    sheets[sheet_key] = df
                    total_text_lines += len(lines)
        doc.close()
        if sheets and total_text_lines > 15:
            return sheets
        return None
    except ImportError:
        return None
    except Exception:
        return None

# ============================================================================
# PDF TO EXCEL CONVERTER
# ============================================================================

def find_page10_in_pdf(doc):
    """Search ALL pages for Page 10 by checking content."""
    for page_num in range(len(doc)):
        page = doc[page_num]
        words = page.get_text("words")
        if not words or len(words) < 20:
            continue
        rows = {}
        for w in words:
            y_key = round(w[1] / 15) * 15
            if y_key not in rows:
                rows[y_key] = []
            rows[y_key].append(w[4])
        sorted_rows = sorted(rows.items())
        line_texts = [' '.join(row_words).upper() for _, row_words in sorted_rows]
        paired_lines = [line_texts[i] + ' ' + line_texts[i+1] for i in range(len(line_texts)-1)]
        combined_text = ' '.join(line_texts) + ' ' + ' '.join(paired_lines)
        excluded = False
        for kw in PAGE10_EXCLUDE_KEYWORDS:
            if kw in combined_text:
                excluded = True
                break
        if excluded:
            continue
        seq = ['REVENUS MENSUELS', 'REVENUS JOURNALIERS', 'REVENUS LAVE-AUTO']
        last_pos = -1
        seq_ok = True
        for term in seq:
            pos = combined_text.find(term)
            if pos == -1 or pos <= last_pos:
                seq_ok = False
                break
            last_pos = pos
        if not seq_ok:
            continue
        expense_found = any(term in combined_text for term in [
            'SALAIRES STATIONNEMENT', 'UNIFORMES', 'ENTRETIEN',
            'TAXES ET PERMIS', 'ASSURANCES', 'TÉLÉCOMMUNICATION'])
        if not expense_found:
            continue
        if 'TOTAL REVENUS' not in combined_text and 'TOTAL DES REVENUS' not in combined_text:
            continue
        if 'BÉNÉFICE NET' not in combined_text and 'BENEFICE NET' not in combined_text:
            continue
        return page_num
    return None

def convert_page10_to_excel(uploaded_file):
    """Convert Page 10 to Excel using coordinate-based extraction."""
    try:
        import fitz
        from openpyxl import Workbook
        from openpyxl.styles import Font
        file_bytes = get_file_bytes(uploaded_file)
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page10_num = find_page10_in_pdf(doc)
        if page10_num is None:
            doc.close()
            return None
        page = doc[page10_num]
        words = page.get_text("words")
        if not words:
            doc.close()
            return None
        rows_dict = {}
        for word in words:
            x0, y0, x1, y1, text, block, line, word_no = word
            y_key = round(y0 / 10) * 10
            if y_key not in rows_dict:
                rows_dict[y_key] = []
            rows_dict[y_key].append((x0, text))
        sorted_y = sorted(rows_dict.keys())
        all_x = [x for y_key in sorted_y for x, _ in rows_dict[y_key]]
        if not all_x:
            doc.close()
            return None
        min_x = min(all_x)
        max_x = max(all_x)
        col_width = (max_x - min_x) / 9
        wb = Workbook()
        ws = wb.active
        ws.title = "Page10"
        headers = ["Account", "Mois Courant", "Budget période", "Écart Budget",
                   "An. Préc.", "Cumulatif courant", "Cumulatif budget",
                   "Écart Budget Cumul.", "An. Préc. Cumul."]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
        excel_row = 2
        for y_key in sorted_y:
            row_items = rows_dict[y_key]
            row_items.sort(key=lambda item: item[0])
            cols = [''] * 9
            for x, text in row_items:
                col_idx = min(8, max(0, int((x - min_x) / col_width)))
                cols[col_idx] = (cols[col_idx] + ' ' + text).strip() if cols[col_idx] else text
            has_content = any(c.strip() for c in cols)
            if has_content:
                for col in range(9):
                    val = cols[col].strip()
                    if val:
                        try:
                            clean = val.replace('$', '').replace(',', '').replace(' ', '')
                            if clean.startswith('(') and clean.endswith(')'):
                                clean = '-' + clean[1:-1]
                            if clean.replace('.', '').replace('-', '').isdigit():
                                ws.cell(row=excel_row, column=col+1, value=float(clean))
                                ws.cell(row=excel_row, column=col+1).number_format = '#,##0.00'
                            else:
                                ws.cell(row=excel_row, column=col+1, value=val)
                        except:
                            ws.cell(row=excel_row, column=col+1, value=val)
                excel_row += 1
        doc.close()
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception:
        return None

def read_pdf_to_dataframe(uploaded_file):
    """
    MAIN PDF READER: Converter -> OCR -> Fitz -> pdfplumber.
    Priority order:
      1. convert_page10_to_excel (coordinate-based, best quality)
      2. read_pdf_with_ocr       (image-based PDFs like January)
      3. read_pdf_with_fitz      (text PDFs, no structure)
      4. pdfplumber              (last resort)
    """
    # ── Path 1: coordinate converter ──────────────────────────────────────
    excel_output = convert_page10_to_excel(uploaded_file)
    if excel_output:
        try:
            xl = pd.ExcelFile(excel_output, engine='openpyxl')
            sheets = {}
            for sheet_name in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                sheets[sheet_name] = df
            if sheets:
                return sheets   # sheet name is "Page10" → picked up by find_best_data_sheet
        except:
            pass

    # ── Path 2: OCR (image-based PDF like January) ─────────────────────────
    ocr_sheets = read_pdf_with_ocr(uploaded_file)
    if ocr_sheets:
        # Keys are already prefixed with "OCR_" inside read_pdf_with_ocr
        return ocr_sheets

    # ── Path 3: fitz plain text ────────────────────────────────────────────
    fitz_sheets = read_pdf_with_fitz(uploaded_file)
    if fitz_sheets:
        return fitz_sheets

    # ── Path 4: pdfplumber ─────────────────────────────────────────────────
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
                            clean_headers = [str(h).strip() if h else "" for h in headers]
                            df = pd.DataFrame(data, columns=clean_headers)
                            sheet_key = f"Page{page_num+1}_Table{table_num+1}"
                            sheets[sheet_key] = df
                text = page.extract_text()
                if text and len(text.strip()) > 50:
                    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
                    if lines:
                        df = pd.DataFrame(lines, columns=['Text'])
                        sheet_key = f"Page{page_num+1}_Text"
                        sheets[sheet_key] = df
        if sheets:
            return sheets
    except Exception:
        pass

    return None

def read_any_file_to_dataframes(uploaded_file):
    if uploaded_file is None:
        return None, None
    if is_excel_file(uploaded_file):
        result = read_excel_to_dataframe(uploaded_file)
        return (result, "excel") if result else (None, None)
    if is_csv_file(uploaded_file):
        result = read_csv_to_dataframe(uploaded_file)
        return (result, "csv") if result else (None, None)
    if is_pdf_file(uploaded_file):
        result = read_pdf_to_dataframe(uploaded_file)
        return (result, "pdf_converted") if result else (None, None)
    try:
        result = read_excel_to_dataframe(uploaded_file)
        return (result, "excel") if result else (None, None)
    except:
        pass
    try:
        result = read_csv_to_dataframe(uploaded_file)
        return (result, "csv") if result else (None, None)
    except:
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
    sheets_dict, _ = read_any_file_to_dataframes(pnl_file)
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
                    except:
                        continue
    return list(dict.fromkeys(codes))

def find_sheet_by_pattern(wb, patterns):
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            if pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ') in sheet_lower:
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
        return max(sheets_dict.keys(), key=lambda n: len(sheets_dict[n]))
    return None

def safe_float(value, default=0.0):
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '').replace('\xa0', '').replace('€', '')
            if value.startswith('(') and value.endswith(')'):
                return -safe_float(value[1:-1], default)
            if value.startswith('-'):
                return -safe_float(value[1:], default)
        return float(value)
    except (ValueError, TypeError):
        return default

def clean_text_for_matching(text):
    if not text:
        return ""
    text = text.lower().strip()
    for c in 'éèêëàâäîïôöûüùçœ':
        text = text.replace(c, 'e')
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
        return ratio if ratio >= 0.5 else 0
    if cell_clean in label_clean:
        ratio = len(cell_clean) / len(label_clean)
        return ratio if ratio >= 0.5 else 0
    return 0

def extract_dollar_amount_from_text(text):
    if not text:
        return None
    for pattern in [r'\$([\d,]+\.?\d*)', r'([\d,]+\.?\d*)\s*\$', r'\$?\s*([\d,]+\.\d{2})\s*\$?', r'\(?\$?([\d,]+\.?\d*)\)?']:
        match = re.search(pattern, text)
        if match:
            try:
                return float(match.group(1).replace(',', ''))
            except:
                continue
    match = re.search(r'([\d,]+\.\d{2})', text)
    if match:
        try:
            return float(match.group(1).replace(',', ''))
        except:
            pass
    return None

def parse_text_line_for_data(line):
    if not line or len(line) < 5:
        return None, None
    line = line.strip()
    val = extract_dollar_amount_from_text(line)
    for mapping_label, standard_label in ALL_LABEL_MAPPINGS.items():
        if label_match_score(line, mapping_label) >= 0.6:
            return standard_label, val
    if val is not None:
        label_text = re.sub(r'\$?[\d,]+\.?\d*\s*\$?', '', line).strip()
        label_text = re.sub(r'\(?[\d,]+\.?\d*\)?', '', label_text).strip()
        label_text = re.sub(r'\s+', ' ', label_text).strip()
        if len(label_text) > 2:
            for mapping_label, standard_label in ALL_LABEL_MAPPINGS.items():
                if label_match_score(label_text, mapping_label) >= 0.6:
                    return standard_label, val
    return None, val

def read_year_mapping_from_template(wb):
    dh_sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
    if not dh_sheet_name:
        current_year = datetime.now().year
        return {i: current_year if i < 4 else current_year - 1 for i in range(12)}
    ws = wb[dh_sheet_name]
    for row_idx in range(1, 20):
        year_map = {}
        for col_idx in range(2, 14):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is not None and not (isinstance(cell_value, str) and str(cell_value).startswith('=')):
                year_match = re.search(r'(20\d{2})', str(cell_value).strip())
                if year_match:
                    year_map[col_idx - 2] = int(year_match.group(1))
        if len(year_map) >= 6:
            return year_map
    current_year = datetime.now().year
    return {i: current_year if i < 4 else current_year - 1 for i in range(12)}

def extract_month_year_from_text(text):
    if not text:
        return None, None
    text_lower = text.lower()
    year_match = re.search(r'(20\d{2})', text)
    year = int(year_match.group(1)) if year_match else None
    for month_name in MONTH_NAMES_MAP:
        if month_name in text_lower:
            return MONTH_NAMES_MAP[month_name], year
    return None, year

# ============================================================================
# FIX 1: find_amount_column — smarter column detection
# ============================================================================

def find_amount_column(df):
    """
    Find the column index that contains the current-month amounts.

    Priority:
      1. Match column header name (converter output has "Mois Courant" at index 1)
      2. Scan columns 1-8, pick the one with the most non-zero numeric values
      3. Default to index 1 (not 2) as last resort

    The old code defaulted to index 2 when fewer than 3 non-zero values were
    found, which caused 0 matches for any file processed through the converter.
    """
    if df is None or len(df) == 0 or 'Text' in df.columns:
        return None

    # ── Priority 1: match by column header name ──────────────────────────
    col_names_lower = [str(c).lower().strip() for c in df.columns]
    for preferred in ['mois courant', 'current month', 'current period', 'période']:
        if preferred in col_names_lower:
            found_idx = col_names_lower.index(preferred)
            return found_idx

    # ── Priority 2: scan all numeric columns, pick the one with most hits ─
    best_col, best_count = None, 0
    # Check up to 50 rows for a more reliable count (old code only checked 30)
    scan_rows = min(50, len(df))
    for col_idx in range(1, min(9, len(df.columns))):
        numeric_count = sum(
            1 for row_idx in range(scan_rows)
            if safe_float(df.iloc[row_idx, col_idx]) != 0
        )
        if numeric_count > best_count:
            best_count = numeric_count
            best_col = col_idx

    # ── Priority 3: default to 1, not 2 ───────────────────────────────────
    # Even a single non-zero value is better than blindly using col 2.
    if best_col is not None and best_count >= 1:
        return best_col

    return 1   # ← was 2; col 1 is always "Mois Courant" in converter output


# ============================================================================
# FIX 2: find_best_data_sheet — add OCR sheet priority
# ============================================================================

def find_best_data_sheet(sheets_dict):
    """
    Sheet priority:
      1. Page10   — converter output (best quality)
      2. OCR_*    — OCR output (image PDFs like January)
      3. *Fitz*   — plain fitz text
      4. Largest sheet by text cell count
    """
    if not sheets_dict:
        return None

    # Priority 1: converter output
    for name in sheets_dict:
        if 'Page10' in name:
            return name

    # Priority 2: OCR output (prefixed by read_pdf_with_ocr)
    ocr_sheets = [n for n in sheets_dict if n.startswith('OCR_')]
    if ocr_sheets:
        return max(ocr_sheets, key=lambda n: len(sheets_dict[n]))

    # Priority 3: fitz text (picks the page with most content)
    fitz_sheets = [n for n in sheets_dict if 'Fitz' in n]
    if fitz_sheets:
        return max(fitz_sheets, key=lambda n: len(sheets_dict[n]))

    # Priority 4: largest sheet by non-empty text cells
    candidates = []
    for name, df in sheets_dict.items():
        if df is None or len(df) == 0:
            continue
        text_cells = sum(
            1 for row_idx in range(min(40, len(df)))
            for col_idx in range(min(10, len(df.columns)))
            if str(df.iloc[row_idx, col_idx]).strip().lower() not in ['nan', 'none', '']
            and len(str(df.iloc[row_idx, col_idx]).strip()) > 2
        )
        if text_cells > 3:
            candidates.append((name, text_cells))
    if candidates:
        return max(candidates, key=lambda x: x[1])[0]

    return next((name for name, df in sheets_dict.items() if df is not None and len(df) > 0), None)


# ============================================================================
# EXTRACTION FUNCTIONS
# ============================================================================

def extract_data_from_text_sheet(df, month_name, year):
    result = {}
    all_lines = []
    for row_idx in range(len(df)):
        try:
            if 'Text' in df.columns:
                line = str(df.iloc[row_idx, 0]).strip()
            else:
                parts = [str(df.iloc[row_idx, c]).strip() for c in range(min(10, len(df.columns)))
                        if str(df.iloc[row_idx, c]).strip().lower() not in ['nan', 'none', '']]
                line = ' '.join(parts)
            if line and len(line) > 2:
                all_lines.append(line)
        except:
            continue
    for i, line in enumerate(all_lines):
        line_upper = line.upper()
        if 'TOTAL REVENUS' in line_upper or 'TOTAL DES REVENUS' in line_upper:
            val = extract_dollar_amount_from_text(line)
            if not val and i + 1 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+1])
            if val:
                result['_REVENUE_TOTAL_'] = val
            continue
        if any(t in line_upper for t in ['TOTAL DES FRAIS', 'TOTAL OPERATING', 'TOTAL DEPENSES']):
            val = extract_dollar_amount_from_text(line)
            if not val and i + 1 < len(all_lines):
                val = extract_dollar_amount_from_text(all_lines[i+1])
            if val:
                result['_EXPENSE_TOTAL_'] = val
            continue
        std_label, val = None, extract_dollar_amount_from_text(line)
        for mapping_label, standard_label in ALL_LABEL_MAPPINGS.items():
            if label_match_score(line, mapping_label) >= 0.6:
                std_label = standard_label
                break
        if std_label and not val:
            for offset in [1, 2]:
                if i + offset < len(all_lines):
                    val = extract_dollar_amount_from_text(all_lines[i+offset])
                    if val:
                        break
        if not std_label and val:
            for offset in [1, 2]:
                if i - offset >= 0:
                    for mapping_label, standard_label in ALL_LABEL_MAPPINGS.items():
                        if label_match_score(all_lines[i-offset], mapping_label) >= 0.6:
                            std_label = standard_label
                            break
                    if std_label:
                        break
        if std_label and val:
            result[std_label] = result.get(std_label, 0) + val
    result['_DEBUG_MATCHES_'] = str(len([k for k in result if not k.startswith('_')]))
    return result


# ============================================================================
# FIX 3: extract_monthly_data_from_file — scan ALL label columns per row
# ============================================================================

def extract_monthly_data_from_file(uploaded_file):
    """
    Main extraction function.

    Key fix: for table-based sheets (converter output), the inner loop now
    checks ALL columns 0-2 for a label match and takes the best-scoring one
    instead of breaking on the first column.  This prevents a spurious match
    in a non-label column from shadowing the real account-name column.
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

    # find_best_data_sheet now handles Page10 → OCR_ → Fitz → fallback
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

    if hasattr(uploaded_file, 'name'):
        month_name, year = extract_month_year_from_text(uploaded_file.name)

    # ── Text-based sheets (OCR / fitz / pdfplumber text) ──────────────────
    if 'Text' in df.columns:
        result['_DEBUG_METHOD_'] = "text_based"
        data = extract_data_from_text_sheet(df, month_name, year)
        for key, value in data.items():
            result[key] = value
        return result, (month_name, year)

    # ── Table-based sheets (converter Excel / native Excel) ────────────────
    result['_DEBUG_METHOD_'] = "table_based"
    amount_col = find_amount_column(df)
    result['_DEBUG_AMOUNT_COL_'] = str(amount_col) if amount_col is not None else "None"
    if amount_col is None:
        amount_col = 1  # safe default for converter output

    matches_found = 0

    for row_idx in range(len(df)):
        try:
            # ── TOTAL markers: check first 3 columns ──────────────────────
            row_text_upper = ' '.join(
                str(df.iloc[row_idx, c]).upper()
                for c in range(min(3, len(df.columns)))
                if str(df.iloc[row_idx, c]).strip().lower() not in ('nan', 'none', '')
            )
            if 'TOTAL REVENUS' in row_text_upper or 'TOTAL REVENUE' in row_text_upper:
                val = safe_float(df.iloc[row_idx, amount_col])
                if val != 0:
                    result['_REVENUE_TOTAL_'] = val
            if 'TOTAL DES FRAIS' in row_text_upper or 'TOTAL OPERATING' in row_text_upper:
                val = safe_float(df.iloc[row_idx, amount_col])
                if val != 0:
                    result['_EXPENSE_TOTAL_'] = val

            # ── Label matching: check ALL label columns, keep best score ──
            # Converter output: account names are in col 0 only.
            # Other formats may spread labels across cols 0-2.
            # Checking all 3 and taking the best match is safe for both cases.
            best_std = None
            best_score = 0.59  # minimum threshold

            for col_idx in range(min(3, len(df.columns))):
                cell_text = str(df.iloc[row_idx, col_idx]).strip()
                if not cell_text or len(cell_text) < 3 or cell_text.lower() in ('nan', 'none'):
                    continue
                for label, standard in ALL_LABEL_MAPPINGS.items():
                    score = label_match_score(cell_text, label)
                    if score > best_score:
                        best_score = score
                        best_std = standard

            if best_std:
                val = safe_float(df.iloc[row_idx, amount_col])
                if val != 0:
                    matches_found += 1
                    result[best_std] = result.get(best_std, 0) + val

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
            'month': month_name, 'year': year,
            'type': file_data.pop('_DEBUG_TYPE_', '?'),
            'target': file_data.pop('_DEBUG_TARGET_', '?'),
            'rows': file_data.pop('_DEBUG_ROWS_', '?'),
            'cols': file_data.pop('_DEBUG_COLS_', '?'),
            'method': file_data.pop('_DEBUG_METHOD_', '?'),
            'matches': file_data.pop('_DEBUG_MATCHES_', '?'),
            'amount_col': file_data.pop('_DEBUG_AMOUNT_COL_', '?'),
            'error': file_data.pop('_DEBUG_ERROR_', None),
        }
        monthly_data.setdefault('_debug_info', []).append(debug_info)
        if not file_data or month_name is None:
            continue
        revenue_total = file_data.pop('_REVENUE_TOTAL_', None)
        expense_total = file_data.pop('_EXPENSE_TOTAL_', None)
        if revenue_total or expense_total:
            monthly_totals[month_name] = {"revenue_total": revenue_total, "expense_total": expense_total}
        for label, value in file_data.items():
            if label.startswith('_'):
                continue
            monthly_data.setdefault(label, {})[month_name] = value
            yearly_data[label] = yearly_data.get(label, 0) + value
    debug_list = monthly_data.pop('_debug_info', None)
    if not monthly_data:
        return {'monthly': {}, 'yearly': {}, '_debug_info': debug_list} if debug_list else None
    result = {'monthly': monthly_data, 'yearly': yearly_data}
    if monthly_totals:
        result['_monthly_totals'] = monthly_totals
    if debug_list:
        result['_debug_info'] = debug_list
    return result

# ============================================================================
# P&L / PAGE 3 EXTRACTION
# ============================================================================

def find_ytd_column(df):
    if df is None or len(df) == 0 or 'Text' in df.columns:
        return None
    for col_idx in [5, 6, 7]:
        if col_idx < len(df.columns) and any(safe_float(df.iloc[r, col_idx]) != 0 for r in range(min(20, len(df)))):
            return col_idx
    return None

def extract_page3_data(uploaded_file):
    result = {'monthly': {}, 'yearly': {}}
    sheets_dict, _ = read_any_file_to_dataframes(uploaded_file)
    if not sheets_dict:
        return None
    target_sheet = find_best_data_sheet(sheets_dict)
    if not target_sheet:
        return None
    df = sheets_dict[target_sheet]
    if df is None or len(df) == 0:
        return None
    if 'Text' in df.columns:
        data = extract_data_from_text_sheet(df, None, None)
        for k, v in data.items():
            if not k.startswith('_'):
                result['yearly'][k] = v
    else:
        ytd_col = find_ytd_column(df)
        if ytd_col is None:
            return None
        for row_idx in range(len(df)):
            try:
                best_std = None
                best_score = 0.59
                for col_idx in range(min(3, len(df.columns))):
                    cell_text = str(df.iloc[row_idx, col_idx]).strip()
                    if len(cell_text) < 3 or cell_text.lower() in ('nan', 'none'):
                        continue
                    for label, standard in ALL_LABEL_MAPPINGS.items():
                        score = label_match_score(cell_text, label)
                        if score > best_score:
                            best_score = score
                            best_std = standard
                if best_std:
                    val = safe_float(df.iloc[row_idx, ytd_col])
                    if val:
                        result['yearly'][best_std] = val
            except:
                continue
    return result if result['yearly'] else None

def extract_pnl_data_from_dataframe(df, sheet_name_hint=None):
    result = {'monthly': {}, 'yearly': {}}
    if df is None or len(df) == 0:
        return result
    for row_idx in range(1, len(df)):
        try:
            label = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
            if not label or label.lower() in ['code', '', 'nan', 'none']:
                continue
            monthly = {MONTHS_EN[m]: safe_float(df.iloc[row_idx, m+1]) for m in range(12) if m+1 < len(df.columns)}
            yearly_total = safe_float(df.iloc[row_idx, -1]) if len(df.columns) > 1 else 0
            clean_label = label.strip().replace('  ', ' ')
            result['monthly'][clean_label] = monthly
            result['yearly'][clean_label] = yearly_total
        except:
            continue
    return result

def extract_pnl_data(uploaded_file, parking_code):
    sheets_dict, _ = read_any_file_to_dataframes(uploaded_file)
    if not sheets_dict:
        return None, None
    sheet_name = find_sheet_in_dict(sheets_dict, parking_code)
    if not sheet_name and sheets_dict:
        sheet_name = max(sheets_dict.keys(), key=lambda n: len(sheets_dict[n]))
    if not sheet_name:
        return None, None
    df = sheets_dict[sheet_name]
    return extract_pnl_data_from_dataframe(df, sheet_name), None

def find_pnl_value(pnl_data, label_alternatives):
    if pnl_data is None:
        return 0
    yearly = pnl_data.get('yearly', {})
    clean_alts = [clean_text_for_matching(a) for a in label_alternatives if len(clean_text_for_matching(a)) >= 3]
    for alt in clean_alts:
        for key, val in yearly.items():
            key_clean = clean_text_for_matching(key)
            if alt == key_clean or (alt in key_clean and len(alt) >= 5 and len(alt)/len(key_clean) >= 0.6) or (key_clean in alt and len(key_clean) >= 5 and len(key_clean)/len(alt) >= 0.6):
                return val
    return 0

def find_monthly_pnl_value(monthly_data, label_alternatives):
    if not monthly_data:
        return {}
    clean_alts = [clean_text_for_matching(a) for a in label_alternatives if len(clean_text_for_matching(a)) >= 3]
    for alt in clean_alts:
        for key, monthly in monthly_data.items():
            key_clean = clean_text_for_matching(key)
            if alt == key_clean or (alt in key_clean and len(alt) >= 5 and len(alt)/len(key_clean) >= 0.6) or (key_clean in alt and len(key_clean) >= 5 and len(key_clean)/len(alt) >= 0.6):
                return monthly
    return {}

def merge_monthly_data(current_year_data, previous_year_data, year_map):
    merged = {}
    current_year = datetime.now().year
    for month_idx, year in year_map.items():
        month_name = MONTHS_EN[month_idx]
        source = current_year_data if year == current_year else previous_year_data if year < current_year else current_year_data
        if source and 'monthly' in source:
            for label, monthly_values in source['monthly'].items():
                merged.setdefault(label, {})[month_name] = monthly_values.get(month_name, 0)
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
            return ["⚠️ Budget Initial: No data available"]
        cells_updated, filled_revenue, filled_expense = 0, 0, 0
        for dh_row, pnl_labels in DH_ROW_MAPPING.items():
            yearly_value = find_pnl_value(bi_data, pnl_labels)
            if yearly_value != 0:
                ws[f"S{dh_row}"] = yearly_value
                ws[f"S{dh_row}"].number_format = '#,##0.00'
                cells_updated += 1
                updates.append(f"✅ BI Row {dh_row}: ${yearly_value:,.2f} ({pnl_labels[0]})")
                if dh_row in REVENUE_ROWS:
                    filled_revenue += yearly_value
                elif dh_row in EXPENSE_ROWS:
                    filled_expense += yearly_value
        updates.append(f"✅ Budget Initial: {cells_updated} cells updated" if cells_updated else "⚠️ Budget Initial: No cells updated")
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
            return ["⚠️ Fiche Stationnement: No data available"]
        for cell, pnl_labels in FICHE_STATIONNEMENT_MAP:
            yearly_value = find_pnl_value(fs_data, pnl_labels)
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00'
                updates.append(f"✅ {cell} = ${yearly_value:,.2f}")
            else:
                updates.append(f"⚠️ {cell} = Not found")
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
        updates.append(f"📅 Year mapping: Jan={year_map.get(0)}, Apr={year_map.get(3)}, May={year_map.get(4)}, Dec={year_map.get(11)}")
        if not merged_monthly_data:
            return ["⚠️ Donnees Historiques: No merged monthly data available"]
        monthly_filled_revenue, monthly_filled_expense = {}, {}
        cells_updated, rows_filled = 0, []
        for dh_row, pnl_labels in DH_ROW_MAPPING.items():
            monthly_values = find_monthly_pnl_value(merged_monthly_data, pnl_labels)
            if not monthly_values or all(v == 0 for v in monthly_values.values()):
                continue
            row_cells = 0
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values:
                    val = monthly_values[month_name]
                    if val != 0:
                        col_letter = get_column_letter(month_idx + 2)
                        ws_dh[f"{col_letter}{dh_row}"] = val
                        ws_dh[f"{col_letter}{dh_row}"].number_format = '#,##0.00'
                        cells_updated += 1
                        row_cells += 1
                        if dh_row in REVENUE_ROWS:
                            monthly_filled_revenue[month_name] = monthly_filled_revenue.get(month_name, 0) + val
                        elif dh_row in EXPENSE_ROWS:
                            monthly_filled_expense[month_name] = monthly_filled_expense.get(month_name, 0) + val
            if row_cells > 0:
                rows_filled.append(f"  Row {dh_row}: {pnl_labels[0]} ({row_cells} months)")
        if monthly_totals:
            balancing_updates = []
            for month_name, totals in monthly_totals.items():
                if month_name not in MONTHS_EN:
                    continue
                month_idx = MONTHS_EN.index(month_name)
                col_letter = get_column_letter(month_idx + 2)
                for total_type, catch_row, filled_dict in [
                    ("revenue_total", REVENUE_CATCH_ALL_ROW, monthly_filled_revenue),
                    ("expense_total", EXPENSE_CATCH_ALL_ROW, monthly_filled_expense)
                ]:
                    expected = totals.get(total_type)
                    if expected and expected != 0:
                        actual = filled_dict.get(month_name, 0)
                        gap = expected - actual
                        if abs(gap) > 0.99:
                            cell_ref = f"{col_letter}{catch_row}"
                            current_val = safe_float(ws_dh[cell_ref].value)
                            ws_dh[cell_ref] = current_val + gap
                            ws_dh[cell_ref].number_format = '#,##0.00'
                            balancing_updates.append(f"  ⚖️ {month_name}: Added ${gap:,.2f} to Row {catch_row}")
                            cells_updated += 1
            if balancing_updates:
                updates.append(f"⚖️ Balancing ({len(balancing_updates)} adjustments):")
                updates.extend(balancing_updates)
        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: {cells_updated} cells in {len(rows_filled)} rows")
            updates.extend(rows_filled)
        else:
            updates.append("⚠️ Donnees Historiques: No cells updated")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {str(e)}")
    return updates

# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    updates = []
    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)
    if not parking_code:
        return None, ["❌ Could not determine parking code."]
    updates.append(f"🔍 Processing: {parking_code}")
    dh_current_year_data = dh_previous_year_data = None
    monthly_totals = None

    if monthly_files_current:
        updates.append("📋 Processing Current Year monthly files")
        dh_current_year_data = build_monthly_data_from_files(monthly_files_current)
        if dh_current_year_data:
            debug_info = dh_current_year_data.pop('_debug_info', None)
            if debug_info:
                for d in debug_info:
                    if d.get('error'):
                        updates.append(f"🔧 {d['file']}: ERROR - {d['error']}")
                    else:
                        updates.append(f"🔧 {d['file']}: type={d['type']}, target={d['target']}, rows={d['rows']}x{d['cols']}, method={d.get('method','?')}, matches={d['matches']}, amount_col={d.get('amount_col','?')}, month={d['month']}")
            updates.append(f"📊 Current year: {len(dh_current_year_data.get('yearly', {}))} labels")
            if '_monthly_totals' in dh_current_year_data:
                monthly_totals = dh_current_year_data.pop('_monthly_totals')

    if monthly_files_previous:
        updates.append("📋 Processing Previous Year monthly files")
        dh_previous_year_data = build_monthly_data_from_files(monthly_files_previous)
        if dh_previous_year_data:
            debug_info = dh_previous_year_data.pop('_debug_info', None)
            if debug_info:
                for d in debug_info:
                    updates.append(f"🔧 {d['file']}: {'ERROR - '+d['error'] if d.get('error') else f'type={d[\"type\"]}, target={d[\"target\"]}, matches={d[\"matches\"]}'}")
            updates.append(f"📊 Previous year: {len(dh_previous_year_data.get('yearly', {}))} labels")
            if '_monthly_totals' in dh_previous_year_data:
                prev_totals = dh_previous_year_data.pop('_monthly_totals')
                monthly_totals = {**(monthly_totals or {}), **prev_totals}

    bi_data = fs_data = None
    if budget_initial_file:
        updates.append("📋 Processing Budget Initial source")
        bi_data = extract_page3_data(budget_initial_file)
        if not bi_data or not bi_data.get('yearly'):
            bi_data, _ = extract_pnl_data(budget_initial_file, parking_code)
        updates.append(f"📊 Budget Initial: {len(bi_data.get('yearly', {})) if bi_data else 0} labels")
    if fiche_stationnement_file:
        updates.append("📋 Processing Fiche Stationnement source")
        fs_data = extract_page3_data(fiche_stationnement_file)
        if not fs_data or not fs_data.get('yearly'):
            fs_data, _ = extract_pnl_data(fiche_stationnement_file, parking_code)
        updates.append(f"📊 Fiche Stationnement: {len(fs_data.get('yearly', {})) if fs_data else 0} labels")

    if not dh_current_year_data and not dh_previous_year_data:
        updates.append("⚠️ No monthly data available")

    try:
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        file_bytes = excel_file.read()
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
        wb_write = load_workbook(io.BytesIO(file_bytes))
        wb_read = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, [f"❌ Error reading template: {str(e)}"]

    year_map = read_year_mapping_from_template(wb_read)
    merged_monthly = {}
    if year_map and dh_current_year_data and dh_previous_year_data:
        merged_monthly = merge_monthly_data(dh_current_year_data, dh_previous_year_data, year_map)
    if not merged_monthly and dh_current_year_data:
        merged_monthly = dh_current_year_data['monthly']
    dh_data = merged_monthly or (dh_current_year_data['monthly'] if dh_current_year_data else {})

    updates.extend(update_budget_initial(wb_write, bi_data, parking_code))
    updates.extend(update_fiche_stationnement(wb_write, fs_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb_write, dh_data, parking_code, monthly_totals))

    if not any(u.startswith("✅") for u in updates):
        updates.append("💡 No updates were made.")

    output = io.BytesIO()
    wb_write.save(output)
    output.seek(0)
    return output, updates
