# excel_fixer.py - WITH DEBUG + FIXED DIGITAL EXTRACTION
import fitz
import os
import re
from io import BytesIO
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ============================================
# CONFIGURATION
# ============================================
DEBUG = True  # Set to False after fixing

TABLE_TOP = 0.06
TABLE_BOTTOM = 0.94

PAGE_MARKERS = ["1981 McGill College", "Revenus mensuels", "BÉNÉFICE NET", "Mois Courant"]

ACCOUNT_PATTERNS = [
    ('TOTAL REVENUS', ['total revenus']),
    ('BÉNÉFICE NET', ['bénéfice net', 'benefice net']),
    ('Revenus mensuels', ['revenus mensuels']),
    ('Revenus Journaliers', ['revenus journaliers', 'revenus horaires']),
    ('Revenus Lave-Auto', ['revenus lave-auto', 'lave-auto']),
    ('Divers', ['divers']),
    ('Gratuités - mensuels', ['gratuités - mensuels', 'gratuités', 'gratuite']),
    ('Salaires Stationnement', ['salaires stationnement', 'salaire stationnement']),
    ('Uniformes', ['uniformes', 'uniforme']),
    ('Fourn. de stationnement', ['fourn. de stationnement', 'fournitures stationnement']),
    ('Entretien réparation - Nettoyage', ['nettoyage']),
    ('Entretien réparation - Equipement', ['equipement', 'équipement']),
    ('Entretien réparation - Général', ['général', 'general']),
    ('Taxes et permis', ['taxes et permis', 'taxe']),
    ('Assurances Cautionnement', ['assurances', 'cautionnement']),
    ('Réclamations', ['réclamations', 'reclamation']),
    ('Télécommunication', ['télécommunication', 'telecommunication']),
    ('Frais de cartes de crédit', ['cartes de crédit', 'credit card']),
    ('Frais de bureau', ['frais de bureau']),
    ('Honoraires de gestion', ['honoraires de gestion']),
]

MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

# ONLY YELLOW CELLS - no formula rows
PDF_TO_TEMPLATE = {
    'Revenus Journaliers': 12,
    'Revenus mensuels': 13,
    'Revenus Lave-Auto': 14,
    'Divers': 17,
    'Gratuités - mensuels': 20,
    'Salaires Stationnement': 29,
    'Uniformes': 32,
    'Entretien réparation - Nettoyage': 35,
    'Entretien réparation - Général': 36,
    'Entretien réparation - Equipement': 37,
    'Fourn. de stationnement': 41,
    'Frais de bureau': 49,
    'Télécommunication': 50,
    'Frais de cartes de crédit': 53,
    'Réclamations': 56,
    'Assurances Cautionnement': 57,
    'Taxes et permis': 58,
    'Honoraires de gestion': 63,
}

def safe_float(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip()
    is_negative = False
    if value.startswith('(') and value.endswith(')'):
        is_negative = True
        value = value[1:-1].strip()
    value = value.replace('$', '').strip()
    value = value.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    if value.startswith('-'):
        is_negative = True
        value = value[1:]
    elif value.endswith('-'):
        is_negative = True
        value = value[:-1]
    if ',' in value and '.' not in value:
        value = value.replace(',', '.')
    elif ',' in value and '.' in value:
        if value.rfind('.') > value.rfind(','):
            value = value.replace(',', '')
        else:
            value = value.replace('.', '').replace(',', '.')
    value = re.sub(r'[^\d\.\-]', '', value)
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
    text_lower = text.lower().strip()
    text_lower = re.sub(r'[|\[\]{}()]', '', text_lower)
    text_lower = re.sub(r'\s+', ' ', text_lower)
    for standard_name, patterns in ACCOUNT_PATTERNS:
        for pattern in patterns:
            pattern_words = pattern.split()
            if all(word in text_lower for word in pattern_words):
                return standard_name
    return None

def find_pl_page(pdf_path):
    """Find P&L page - check ALL pages and show what we find"""
    doc = fitz.open(pdf_path)
    
    if DEBUG:
        print(f"   DEBUG: Searching {len(doc)} pages for P&L...")
    
    for i in range(len(doc)):
        text = doc[i].get_text()
        
        # Check each marker individually
        found = []
        missing = []
        for m in PAGE_MARKERS:
            if m.lower() in text.lower():
                found.append(m)
            else:
                missing.append(m)
        
        if DEBUG and len(found) >= 2:
            print(f"   DEBUG: Page {i+1}: {len(found)}/{len(PAGE_MARKERS)} markers found: {found}")
        
        if len(found) >= 3:
            if DEBUG:
                print(f"   DEBUG: ✅ P&L found on page {i+1}")
            return doc, i
    
    # If not found, show what each page has
    if DEBUG:
        print(f"   DEBUG: P&L NOT FOUND! Checking all pages:")
        for i in range(len(doc)):
            text = doc[i].get_text()[:200]
            if 'revenu' in text.lower() or 'dépense' in text.lower() or 'stationnement' in text.lower():
                print(f"   DEBUG: Page {i+1} has financial content: {text[:100]}...")
    
    doc.close()
    return None, None

def extract_from_digital_pdf(pdf_path):
    """Digital extraction using text blocks"""
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        return None
    
    page = doc[page_num]
    
    # METHOD 1: Try text blocks first
    blocks = page.get_text("blocks")
    blocks.sort(key=lambda b: (round(b[1], 0), b[0]))
    
    h = page.rect.height
    table_top = h * TABLE_TOP
    table_bottom = h * TABLE_BOTTOM
    table_blocks = [b for b in blocks if table_top <= b[1] <= table_bottom]
    
    if DEBUG:
        print(f"   DEBUG: Found {len(table_blocks)} blocks in table area")
    
    # Group by row
    rows = {}
    for block in table_blocks:
        y_key = round(block[1] / 5) * 5
        if y_key not in rows:
            rows[y_key] = []
        rows[y_key].append(block)
    
    sorted_rows = sorted(rows.items())
    data = {}
    
    for y_key, row_blocks in sorted_rows:
        row_blocks.sort(key=lambda b: b[0])
        if not row_blocks:
            continue
        
        account_text = row_blocks[0][4].strip()
        account_text = ' '.join(account_text.split())
        
        standard_name = match_account(account_text)
        if not standard_name and len(row_blocks) >= 2:
            second_text = row_blocks[1][4].strip()
            if not any(c.isdigit() for c in second_text.replace(' ', '').replace(',', '').replace('.', '')):
                combined = account_text + ' ' + ' '.join(second_text.split())
                standard_name = match_account(combined)
        
        if not standard_name:
            continue
        
        amount = None
        for block in row_blocks[1:]:
            block_text = block[4].strip()
            if match_account(block_text):
                continue
            amount = safe_float(block_text)
            if amount is not None:
                break
        
        if amount is not None and standard_name not in data:
            data[standard_name] = amount
    
    # METHOD 2: If blocks didn't get enough, try extracting all text and parsing
    if len(data) < 5:
        if DEBUG:
            print(f"   DEBUG: Blocks only got {len(data)} accounts, trying full text parse...")
        
        text = page.get_text("text")
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            if len(line) < 5:
                continue
            
            standard_name = match_account(line)
            if not standard_name or standard_name in data:
                continue
            
            # Find number in this line
            parts = line.split()
            for part in parts:
                amount = safe_float(part)
                if amount is not None:
                    data[standard_name] = amount
                    break
    
    doc.close()
    
    if DEBUG:
        print(f"   DEBUG: Extracted {len(data)} accounts: {list(data.keys())}")
    
    return data if len(data) >= 3 else None

def extract_from_image_pdf(pdf_path):
    """OCR for January"""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        if DEBUG:
            print(f"   DEBUG: Tesseract not installed")
        return None
    
    doc = fitz.open(pdf_path)
    
    for i in range(len(doc)):
        page = doc[i]
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        try:
            text = pytesseract.image_to_string(img, lang='fra')
        except:
            try:
                text = pytesseract.image_to_string(img, lang='eng')
            except:
                text = pytesseract.image_to_string(img)
        
        if DEBUG:
            print(f"   DEBUG: OCR page {i+1} - found 'revenus': {'revenus' in text.lower()}")
        
        if "revenus" in text.lower() or "stationnement" in text.lower():
            data = {}
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                if len(line) < 5:
                    continue
                
                standard_name = match_account(line)
                if not standard_name or standard_name in data:
                    continue
                
                number_patterns = [
                    r'\(?\s*(\d[\d\s]{0,15},\d{2})\s*\)?\s*\$?',
                    r'\$?\s*(\d[\d\s]{0,15},\d{2})\s*\$?',
                    r'(\d{1,3}(?:\s+\d{3})*,\d{2})',
                ]
                
                for pattern in number_patterns:
                    matches = re.findall(pattern, line)
                    for match in matches:
                        if isinstance(match, tuple):
                            match = match[0]
                        amount = safe_float(match)
                        if amount is not None and abs(amount) > 0.01:
                            data[standard_name] = amount
                            break
                    if standard_name in data:
                        break
            
            doc.close()
            
            if DEBUG:
                print(f"   DEBUG: OCR extracted {len(data)} accounts")
            
            return data if len(data) >= 3 else None
    
    doc.close()
    return None

def fill_template(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        for sn in wb.sheetnames:
            if wb[sn].max_row >= 86:
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
            updates.append(f"   ✅ {month_name} ({col_letter}{template_row}): ${amount:,.2f} - {pdf_account}")
        
        if month_cells > 0:
            updates.append(f"📊 {month_name}: {month_cells} cells filled")
    
    updates.append(f"\n📊 TOTAL: {total_cells} yellow cells filled")
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
        
        benefice_net_pdf = pdf_data.get('BÉNÉFICE NET')
        revenus_nets_template = ws.cell(row=86, column=col).value
        
        if benefice_net_pdf is not None and revenus_nets_template is not None:
            diff = abs(benefice_net_pdf - revenus_nets_template)
            if diff > 0.01:
                results.append(f"⚠️ {month_name}: BÉNÉFICE NET PDF=${benefice_net_pdf:,.2f} ≠ REVENUS NETS Template=${revenus_nets_template:,.2f}")
            else:
                results.append(f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net_pdf:,.2f}")
        elif benefice_net_pdf is None:
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
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
        all_files = []
        if monthly_files_current:
            all_files.extend(monthly_files_current)
        if monthly_files_previous:
            all_files.extend(monthly_files_previous)
        
        if not all_files:
            updates.append("\n⚠️ No files!")
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue(), updates
        
        updates.append(f"\n📁 Processing {len(all_files)} files...")
        
        for file_obj in all_files:
            month = extract_month_from_filename(file_obj)
            file_obj.seek(0)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            try:
                if DEBUG:
                    print(f"\n{'='*60}")
                    print(f"PROCESSING: {month} - {file_obj.name}")
                    print(f"{'='*60}")
                
                # Try digital first
                data = extract_from_digital_pdf(tmp_path)
                
                # If digital fails, try OCR
                if data is None or len(data) < 3:
                    updates.append(f"   🔍 {month}: Digital failed ({len(data or {})} accts), trying OCR...")
                    data = extract_from_image_pdf(tmp_path)
                
                if data and len(data) > 0:
                    all_data[month] = data
                    rev = data.get('TOTAL REVENUS', 'N/A')
                    net = data.get('BÉNÉFICE NET', 'N/A')
                    updates.append(f"   ✅ {month}: {len(data)} accounts | Rev: ${rev} | Net: ${net}")
                else:
                    updates.append(f"   ❌ {month}: No data")
            finally:
                os.unlink(tmp_path)
        
        if all_data:
            updates.append(f"\n📝 Filling yellow cells...")
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
