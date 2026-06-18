# excel_fixer.py - FIXED VERSION (match by text, not row number)
import fitz
import os
import re
from io import BytesIO
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ============================================
# PDF P&L PAGE - LAYOUT (same for all months)
# ============================================
NUM_ROWS = 30  # Extract more rows to be safe
AMOUNT_COL = 1
TABLE_TOP = 0.05
TABLE_BOTTOM = 0.95
TABLE_LEFT = 0.03
TABLE_RIGHT = 0.97
COLUMN_WIDTHS = [0.28, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09]

PAGE_MARKERS = ["1981 McGill College", "Revenus mensuels", "BÉNÉFICE NET", "Mois Courant"]

# Match PDF account names to our standard labels (case-insensitive, partial match)
PDF_ACCOUNT_PATTERNS = {
    'Revenus mensuels': ['revenus mensuels'],
    'Revenus Journaliers': ['revenus journaliers', 'revenus horaires'],
    'Revenus Lave-Auto': ['revenus lave-auto', 'lave-auto'],
    'Divers': ['divers'],
    'Gratuités - mensuels': ['gratuités', 'gratuite'],
    'TOTAL REVENUS': ['total revenus', 'total des revenus'],
    'Salaires Stationnement': ['salaires stationnement', 'salaire stationnement'],
    'Uniformes': ['uniformes', 'uniforme'],
    'Fourn. de stationnement': ['fourn. de stationnement', 'fournitures stationnement'],
    'Entretien réparation - Nettoyage': ['nettoyage'],
    'Entretien réparation - Equipement': ['equipement', 'équipement'],
    'Entretien réparation - Général': ['général', 'general'],
    'Taxes et permis': ['taxes et permis', 'taxe'],
    'Assurances Cautionnement': ['assurances', 'cautionnement'],
    'Réclamations': ['réclamations', 'reclamation'],
    'Télécommunication': ['télécommunication', 'telecommunication'],
    'Frais de cartes de crédit': ['cartes de crédit', 'credit card'],
    'Frais de bureau': ['frais de bureau', 'bureau'],
    "Total des frais d'exploitation": ['total des frais d\'exploitation'],
    "Honoraires de gestion": ['honoraires de gestion'],
    'BÉNÉFICE NET': ['bénéfice net', 'benefice net'],
}

# ============================================
# TEMPLATE MAPPING
# ============================================
MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

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

def match_account(text):
    """Match extracted text to a standard account name"""
    text_lower = text.lower().strip()
    for standard_name, patterns in PDF_ACCOUNT_PATTERNS.items():
        for pattern in patterns:
            if pattern in text_lower:
                return standard_name
    return None

def find_pl_page_digital(pdf_path):
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        if all(m.lower() in text.lower() for m in PAGE_MARKERS):
            return doc, i
    doc.close()
    return None, None

def extract_table_digital(page):
    """Extract rows and match accounts by name"""
    w, h = page.rect.width, page.rect.height
    top = h * TABLE_TOP
    bottom = h * TABLE_BOTTOM
    left = w * TABLE_LEFT
    right = w * TABLE_RIGHT
    table_w = right - left
    row_h = (bottom - top) / NUM_ROWS
    
    data = {}
    
    for row_idx in range(NUM_ROWS):
        y = top + (row_idx * row_h)
        x = left
        
        # Get account name (column 0)
        cell_w = COLUMN_WIDTHS[0] * table_w
        rect = fitz.Rect(x + 1, y + 1, x + cell_w - 1, y + row_h - 1)
        account_text = page.get_text("text", clip=rect).strip()
        account_text = ' '.join(account_text.split())
        
        # Get amount (column 1)
        x += cell_w
        cell_w = COLUMN_WIDTHS[1] * table_w
        rect = fitz.Rect(x + 1, y + 1, x + cell_w - 1, y + row_h - 1)
        amount_text = page.get_text("text", clip=rect).strip()
        amount_text = ' '.join(amount_text.split())
        
        # Match account name
        standard_name = match_account(account_text)
        if standard_name:
            amount = safe_float(amount_text)
            if amount is not None:
                data[standard_name] = amount
    
    return data

def extract_from_digital_pdf(pdf_path):
    doc, page_num = find_pl_page_digital(pdf_path)
    if doc is None:
        return None
    page = doc[page_num]
    data = extract_table_digital(page)
    doc.close()
    return data

# OCR for January
def find_pl_page_ocr(pdf_path):
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return None, None, None
    
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        page = doc[i]
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        try:
            text = pytesseract.image_to_string(img, lang='fra')
        except:
            text = pytesseract.image_to_string(img)
        
        if "revenus mensuels" in text.lower() or "total revenus" in text.lower():
            return doc, i, text
    doc.close()
    return None, None, None

def extract_table_ocr(ocr_text):
    """Parse OCR text by matching account names"""
    lines = ocr_text.split('\n')
    data = {}
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Try to find an account name in the line
        standard_name = match_account(line)
        if standard_name:
            # Extract numbers from the line
            numbers = re.findall(r'-?[\d\s,\.]+', line)
            if numbers:
                # First number after account name is usually the amount
                amount = safe_float(numbers[0])
                if amount is not None:
                    data[standard_name] = amount
    
    return data

def extract_from_image_pdf(pdf_path):
    doc, page_num, ocr_text = find_pl_page_ocr(pdf_path)
    if doc is None:
        return None
    data = extract_table_ocr(ocr_text)
    doc.close()
    return data

def fill_template(wb, all_data):
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

def get_parking_codes_from_pnl(file_obj):
    codes = []
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    updates = []
    all_data = {}
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append("=" * 60)
        
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template: {parking_code or 'Unknown'}")
        
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
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(file_obj.read())
                    tmp_path = tmp.name
                
                try:
                    data = extract_from_digital_pdf(tmp_path)
                    
                    if data is None or len(data) < 3:
                        updates.append(f"   🔍 {month}: Digital failed, trying OCR...")
                        data = extract_from_image_pdf(tmp_path)
                    
                    if data and len(data) > 0:
                        all_data[month] = data
                        rev = data.get('TOTAL REVENUS', 'N/A')
                        net = data.get('BÉNÉFICE NET', 'N/A')
                        updates.append(f"   ✅ {month}: {len(data)} accts | Rev: ${rev} | Net: ${net}")
                    else:
                        updates.append(f"   ❌ {month}: No data")
                finally:
                    os.unlink(tmp_path)
        
        if all_data:
            updates.append(f"\n📝 Filling template...")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            
            updates.append(f"\n🔍 Validation:")
            validations = validate_template(wb, all_data)
            updates.extend(validations)
        else:
            updates.append("\n⚠️ No data extracted!")
        
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
