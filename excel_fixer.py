# excel_fixer.py - FINAL COMPLETE VERSION
import fitz
import os
import re
from io import BytesIO
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pytesseract
from PIL import Image

# ============================================
# PDF P&L PAGE - FIXED 28-ROW LAYOUT (SAME FOR ALL MONTHS)
# ============================================
NUM_ROWS = 28
AMOUNT_COL = 1
TABLE_TOP = 0.06
TABLE_BOTTOM = 0.95
TABLE_LEFT = 0.03
TABLE_RIGHT = 0.97
COLUMN_WIDTHS = [0.25, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10]

PAGE_MARKERS = ["1981 McGill College", "Revenus mensuels", "BÉNÉFICE NET", "Mois Courant"]

# PDF Row -> Account Name (IDENTICAL for ALL months)
PDF_ROW_MAP = {
    2: 'Revenus mensuels',
    3: 'Revenus Journaliers', 
    4: 'Revenus Lave-Auto',
    5: 'Divers',
    7: 'Gratuités - mensuels',
    8: 'TOTAL REVENUS',
    11: 'Salaires Stationnement',
    12: 'Uniformes',
    13: 'Fourn. de stationnement',
    14: 'Entretien réparation - Nettoyage',
    15: 'Entretien réparation - Equipement',
    16: 'Entretien réparation - Général',
    17: 'Taxes et permis',
    18: 'Assurances Cautionnement',
    19: 'Réclamations',
    20: 'Télécommunication',
    21: 'Frais de cartes de crédit',
    22: 'Frais de bureau',
    23: "Total des frais d'exploitation",
    24: "RÉSULTAT D'EXPLOITATION",
    26: 'Honoraires de gestion',
    28: 'BÉNÉFICE NET'
}

# ============================================
# TEMPLATE "2-Données historiques" MAPPING
# ============================================
MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

# EXACT mapping: PDF account -> Template row
PDF_TO_TEMPLATE = {
    'Revenus mensuels': 13,
    'Revenus Journaliers': 12,
    'Revenus Lave-Auto': 14,
    'Divers': 17,
    'Gratuités - mensuels': 20,
    'TOTAL REVENUS': 26,
    'Salaires Stationnement': 29,
    'Uniformes': 32,
    'Fourn. de stationnement': 41,
    'Entretien réparation - Nettoyage': 35,
    'Entretien réparation - Equipement': 37,
    'Entretien réparation - Général': 36,
    'Frais de cartes de crédit': 53,
    'Frais de bureau': 49,
    'Télécommunication': 50,
    'Taxes et permis': 58,
    'Assurances Cautionnement': 57,
    'Réclamations': 56,
    'Honoraires de gestion': 63,
    "Total des frais d'exploitation": 84,
    'BÉNÉFICE NET': 86,
}

def safe_float(value):
    """Convert European numbers: '7 106 417,00' -> 7106417.00"""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip()
    is_negative = value.startswith('(') and value.endswith(')')
    if is_negative:
        value = value[1:-1]
    value = value.replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        result = float(value)
        return -result if is_negative else result
    except:
        return None

def extract_month_from_filename(file_obj):
    """Extract French month name from filename"""
    months = {
        'janvier': 'Janvier', 'février': 'Février', 'mars': 'Mars',
        'avril': 'Avril', 'mai': 'Mai', 'juin': 'Juin',
        'juillet': 'Juillet', 'août': 'Août', 'septembre': 'Septembre',
        'octobre': 'Octobre', 'novembre': 'Novembre', 'décembre': 'Décembre'
    }
    name = file_obj.name.lower() if hasattr(file_obj, 'name') else str(file_obj).lower()
    for key, value in months.items():
        if key in name:
            return value
    return None

# ============================================
# DIGITAL PDF EXTRACTION (Feb-Dec)
# ============================================

def find_pl_page_digital(pdf_path):
    """Find P&L page in digital PDF by content markers"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        if all(m.lower() in text.lower() for m in PAGE_MARKERS):
            return doc, i
    doc.close()
    return None, None

def extract_table_digital(page):
    """Extract 28 rows x 9 columns from digital P&L page"""
    w, h = page.rect.width, page.rect.height
    top = h * TABLE_TOP
    bottom = h * TABLE_BOTTOM
    left = w * TABLE_LEFT
    right = w * TABLE_RIGHT
    table_w = right - left
    row_h = (bottom - top) / NUM_ROWS
    
    data = []
    for row_idx in range(NUM_ROWS):
        row_data = []
        y = top + (row_idx * row_h)
        x = left
        for col_w in COLUMN_WIDTHS:
            cell_w = col_w * table_w
            rect = fitz.Rect(x + 1, y + 1, x + cell_w - 1, y + row_h - 1)
            text = page.get_text("text", clip=rect)
            row_data.append(' '.join(text.strip().split()))
            x += cell_w
        data.append(row_data)
    return data

def extract_from_digital_pdf(pdf_path):
    """Extract data from digital PDF"""
    doc, page_num = find_pl_page_digital(pdf_path)
    if doc is None:
        return None
    
    page = doc[page_num]
    table = extract_table_digital(page)
    doc.close()
    
    return parse_table_data(table)

# ============================================
# IMAGE PDF EXTRACTION (January - OCR)
# ============================================

def find_pl_page_ocr(pdf_path):
    """Find P&L page in image PDF using OCR"""
    doc = fitz.open(pdf_path)
    
    for i in range(len(doc)):
        page = doc[i]
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # OCR with French language
        try:
            text = pytesseract.image_to_string(img, lang='fra')
        except:
            text = pytesseract.image_to_string(img)
        
        # Check for key markers
        matches = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if matches >= 3:
            return doc, i, text
    
    doc.close()
    return None, None, None

def extract_table_ocr(ocr_text):
    """Parse OCR text into 28-row table structure"""
    lines = ocr_text.split('\n')
    
    # Find table start (look for "Revenus mensuels" or "REVENUS DE STATIONNEMENT")
    table_start = 0
    for i, line in enumerate(lines):
        if any(term in line.lower() for term in ['revenus mensuels', 'revenus de stationnement']):
            table_start = i
            break
    
    # Extract 28 lines from table start
    table_lines = lines[table_start:table_start + 30]  # Get extra lines to be safe
    
    data = []
    for i in range(NUM_ROWS):
        if i < len(table_lines):
            # Split line into columns by whitespace
            parts = table_lines[i].split()
            
            # First part is account name (may be multiple words)
            # Last parts are numbers
            account_parts = []
            number_parts = []
            
            for part in parts:
                # Check if part looks like a number
                if re.match(r'^-?[\d\s,\.]+$', part.replace('(', '').replace(')', '')):
                    number_parts.append(part)
                else:
                    account_parts.append(part)
            
            account = ' '.join(account_parts)
            row_data = [account] + number_parts[:8]  # Account + 8 number columns
            
            # Pad to 9 columns
            while len(row_data) < 9:
                row_data.append('')
        else:
            row_data = [''] * 9
        
        data.append(row_data)
    
    return data

def extract_from_image_pdf(pdf_path):
    """Extract data from image-based PDF (January) using OCR"""
    doc, page_num, ocr_text = find_pl_page_ocr(pdf_path)
    if doc is None:
        return None
    
    table = extract_table_ocr(ocr_text)
    doc.close()
    
    return parse_table_data(table)

# ============================================
# COMMON TABLE PARSING
# ============================================

def parse_table_data(table):
    """Parse extracted table into {account: amount} dictionary"""
    data = {}
    
    for row_num, account_name in PDF_ROW_MAP.items():
        if row_num <= len(table):
            raw = table[row_num - 1][AMOUNT_COL] if len(table[row_num - 1]) > AMOUNT_COL else ''
            amount = safe_float(raw)
            if amount is not None:
                data[account_name] = amount
    
    return data

# ============================================
# TEMPLATE FILLING
# ============================================

def fill_template(wb, all_data):
    """Fill '2-Données historiques' sheet"""
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
    
    ws = wb[sheet_name]
    updates = []
    total_cells = 0
    
    for month_name, pdf_data in all_data.items():
        col = MONTH_COLUMN.get(month_name)
        if col is None:
            continue
        
        month_cells = 0
        for pdf_account, amount in pdf_data.items():
            template_row = PDF_TO_TEMPLATE.get(pdf_account)
            if template_row is None:
                continue
            
            cell = ws.cell(row=template_row, column=col)
            cell.value = amount
            cell.number_format = '#,##0.00'
            month_cells += 1
            total_cells += 1
            
            col_letter = get_column_letter(col)
            updates.append(f"   ✅ {month_name} ({col_letter}{template_row}): "
                          f"${amount:,.2f} - {pdf_account}")
        
        if month_cells > 0:
            updates.append(f"📊 {month_name}: {month_cells} cells filled")
    
    updates.append(f"\n📊 TOTAL: {total_cells} cells filled in template")
    return updates

def validate_template(wb, all_data):
    """Validate BÉNÉFICE NET = REVENUS NETS (Template Row 86)"""
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
    
    ws = wb[sheet_name]
    results = []
    
    for month_name, pdf_data in all_data.items():
        col = MONTH_COLUMN.get(month_name)
        if col is None:
            continue
        
        benefice_net = pdf_data.get('BÉNÉFICE NET')
        template_val = ws.cell(row=86, column=col).value
        
        if benefice_net is not None and template_val is not None:
            diff = abs(benefice_net - template_val)
            if diff > 0.01:
                results.append(f"⚠️ {month_name}: BÉNÉFICE NET ${benefice_net:,.2f} ≠ REVENUS NETS ${template_val:,.2f}")
            else:
                results.append(f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net:,.2f}")
        elif benefice_net is None:
            results.append(f"⚠️ {month_name}: BÉNÉFICE NET not extracted from PDF")
    
    return results

# ============================================
# MAIN FUNCTIONS
# ============================================

def get_parking_codes_from_pnl(file_obj):
    """Extract parking codes from uploaded file"""
    codes = []
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """
    MAIN FUNCTION - Extract P&L data from PDFs and fill template.
    Handles both digital PDFs (Feb-Dec) and image PDFs (January).
    """
    updates = []
    all_data = {}
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append("=" * 60)
        
        # Load template
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
        # Process all monthly files
        all_files = []
        if monthly_files_current:
            all_files.extend(monthly_files_current)
        if monthly_files_previous:
            all_files.extend(monthly_files_previous)
        
        if all_files:
            updates.append(f"\n📁 Processing {len(all_files)} files...")
            
            for file_obj in all_files:
                month = extract_month_from_filename(file_obj)
                file_obj.seek(0)
                
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(file_obj.read())
                    tmp_path = tmp.name
                
                try:
                    # Try digital extraction first
                    data = extract_from_digital_pdf(tmp_path)
                    
                    # If digital fails, try OCR (for January/image PDFs)
                    if data is None or len(data) == 0:
                        updates.append(f"   🔍 {month}: Digital extraction failed, trying OCR...")
                        data = extract_from_image_pdf(tmp_path)
                    
                    if data and len(data) > 0:
                        all_data[month] = data
                        rev = data.get('TOTAL REVENUS', 'N/A')
                        net = data.get('BÉNÉFICE NET', 'N/A')
                        updates.append(f"   ✅ {month}: {len(data)} accounts | Revenue: ${rev} | Net Income: ${net}")
                    else:
                        updates.append(f"   ❌ {month}: No data extracted (tried both digital and OCR)")
                finally:
                    os.unlink(tmp_path)
        
        # Fill template
        if all_data:
            updates.append(f"\n📝 Filling template with {len(all_data)} months of data...")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            
            # Validate
            updates.append(f"\n🔍 Validation:")
            validations = validate_template(wb, all_data)
            updates.extend(validations)
        else:
            updates.append("\n⚠️ No data extracted from any file!")
        
        # Save
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("\n✅ WORKFLOW COMPLETE")
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"\n❌ ERROR: {str(e)}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates
