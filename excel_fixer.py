# excel_fixer.py - FINAL VERSION (YELLOW CELLS ONLY)
"""
Extracts P&L data from PDF monthly reports and fills CMO111.xlsx template.
ONLY writes to YELLOW (user input) cells - formulas calculate everything else.
Validates BÉNÉFICE NET (PDF) = REVENUS NETS (Template Row 86 formula).
"""
import fitz
import os
import re
from io import BytesIO
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ============================================
# PDF P&L PAGE - LAYOUT (IDENTICAL FOR ALL MONTHS)
# ============================================
TABLE_TOP = 0.06
TABLE_BOTTOM = 0.94

PAGE_MARKERS = ["1981 McGill College", "Revenus mensuels", "BÉNÉFICE NET", "Mois Courant"]

# Patterns to match account names in PDF
ACCOUNT_PATTERNS = [
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
    # Keep these for validation only - NOT written to template
    ('TOTAL REVENUS', ['total revenus']),
    ('BÉNÉFICE NET', ['bénéfice net', 'benefice net']),
]

# ============================================
# TEMPLATE MAPPING - YELLOW CELLS ONLY
# ============================================
MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

# ONLY map to YELLOW (user input) cells - NO formula rows
PDF_TO_TEMPLATE = {
    # REVENUES - Yellow rows 12-17, 20, 22, 24
    'Revenus Journaliers': 12,    # Revenus horaires
    'Revenus mensuels': 13,       # Revenus mensuels
    'Revenus Lave-Auto': 14,      # Revenus Lave-auto
    'Divers': 17,                 # Autres revenus
    'Gratuités - mensuels': 20,   # (Gratuités)
    
    # LABOUR - Yellow rows 29-32
    'Salaires Stationnement': 29, # Salaire Stationnement
    'Uniformes': 32,              # Uniformes
    
    # MAINTENANCE - Yellow rows 35-43
    'Entretien réparation - Nettoyage': 35,   # Nettoyage stationnement
    'Entretien réparation - Général': 36,     # Entretien stationnement
    'Entretien réparation - Equipement': 37,  # Entretien équipement
    'Fourn. de stationnement': 41,            # Fournitures stationnement
    
    # PUBLIC SERVICES - Yellow row 46
    # (none mapped directly from PDF)
    
    # OVERHEAD - Yellow rows 49-64
    'Frais de bureau': 49,                    # Fournitures de bureau
    'Télécommunication': 50,                  # Télécommunications
    'Frais de cartes de crédit': 53,          # Frais de cartes de crédit
    'Réclamations': 56,                       # Réclamations
    'Assurances Cautionnement': 57,           # Assurances et cautionnement
    'Taxes et permis': 58,                    # Taxes et permis
    'Honoraires de gestion': 63,              # Honoraires de gestion de base
    
    # OTHER EXPENSES - Yellow rows 67-80
    # (none mapped directly from PDF)
}

# ============================================
# UTILITY FUNCTIONS
# ============================================

def safe_float(value):
    """
    Handle Canadian French number formats:
    "113 648,89 $" -> 113648.89
    "(1 206,86) $" -> -1206.86
    "-1 206,86 $" -> -1206.86
    """
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    
    value = str(value).strip()
    
    # Handle parentheses (negative)
    is_negative = False
    if value.startswith('(') and value.endswith(')'):
        is_negative = True
        value = value[1:-1].strip()
    
    # Remove $ signs
    value = value.replace('$', '').strip()
    
    # Remove ALL spaces (thousand separators)
    value = value.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    
    # Handle minus sign
    if value.startswith('-'):
        is_negative = True
        value = value[1:]
    elif value.endswith('-'):
        is_negative = True
        value = value[:-1]
    
    # Handle comma as decimal
    if ',' in value and '.' not in value:
        value = value.replace(',', '.')
    elif ',' in value and '.' in value:
        if value.rfind('.') > value.rfind(','):
            value = value.replace(',', '')
        else:
            value = value.replace('.', '').replace(',', '.')
    
    # Remove non-numeric except minus and decimal
    value = re.sub(r'[^\d\.\-]', '', value)
    
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

def match_account(text):
    """Match extracted text to standard account name"""
    text_lower = text.lower().strip()
    text_lower = re.sub(r'[|\[\]{}()]', '', text_lower)
    text_lower = re.sub(r'\s+', ' ', text_lower)
    
    for standard_name, patterns in ACCOUNT_PATTERNS:
        for pattern in patterns:
            pattern_words = pattern.split()
            if all(word in text_lower for word in pattern_words):
                return standard_name
    return None

# ============================================
# PDF EXTRACTION
# ============================================

def find_pl_page(pdf_path):
    """Find P&L page in PDF"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        matches = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if matches >= 3:
            return doc, i
    doc.close()
    return None, None

def extract_from_digital_pdf(pdf_path):
    """Extract using text blocks - numbers stay intact"""
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        return None
    
    page = doc[page_num]
    blocks = page.get_text("blocks")
    blocks.sort(key=lambda b: (round(b[1], 0), b[0]))
    
    h = page.rect.height
    table_top = h * TABLE_TOP
    table_bottom = h * TABLE_BOTTOM
    table_blocks = [b for b in blocks if table_top <= b[1] <= table_bottom]
    
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
        
        # Get account name from first block(s)
        account_text = row_blocks[0][4].strip()
        account_text = ' '.join(account_text.split())
        
        # Try combining first two blocks for long names
        standard_name = match_account(account_text)
        if not standard_name and len(row_blocks) >= 2:
            second_text = row_blocks[1][4].strip()
            if not any(c.isdigit() for c in second_text):  # Not a number
                combined = account_text + ' ' + ' '.join(second_text.split())
                standard_name = match_account(combined)
        
        if not standard_name:
            continue
        
        # Find amount in subsequent blocks
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
    
    doc.close()
    return data

def extract_from_image_pdf(pdf_path):
    """OCR for January (image PDF)"""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
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
        
        if "revenus" in text.lower() or "stationnement" in text.lower():
            data = {}
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                if len(line) < 5:
                    continue
                
                standard_name = match_account(line)
                if not standard_name:
                    continue
                
                # Find numbers in line
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
                            if standard_name not in data:
                                data[standard_name] = amount
                            break
                    if standard_name in data:
                        break
            
            doc.close()
            return data if data else None
    
    doc.close()
    return None

# ============================================
# TEMPLATE FILLING (YELLOW CELLS ONLY)
# ============================================

def fill_template(wb, all_data):
    """
    Fill ONLY yellow (user input) cells in '2-Données historiques'.
    Formulas will auto-calculate totals.
    """
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
            updates.append(f"⚠️ Unknown month: {month_name}")
            continue
        
        month_cells = 0
        for pdf_account, amount in pdf_data.items():
            template_row = PDF_TO_TEMPLATE.get(pdf_account)
            if template_row is None:
                continue  # Skip formula rows and validation-only accounts
            
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
    
    updates.append(f"\n📊 TOTAL: {total_cells} cells filled (formulas auto-calculate)")
    return updates

def validate_template(wb, all_data):
    """
    Validate BÉNÉFICE NET (PDF) = REVENUS NETS (Template Row 86 formula).
    If mismatch, identify which accounts might be missing.
    """
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
                results.append(f"⚠️ {month_name}: BÉNÉFICE NET ${benefice_net_pdf:,.2f} ≠ REVENUS NETS ${revenus_nets_template:,.2f} (diff: ${diff:,.2f})")
                
                # Check what might be missing
                total_revenus_pdf = pdf_data.get('TOTAL REVENUS')
                total_revenus_template = ws.cell(row=26, column=col).value
                
                if total_revenus_pdf and total_revenus_template:
                    if abs(total_revenus_pdf - total_revenus_template) > 0.01:
                        results.append(f"   🔍 TOTAL REVENUS mismatch: PDF=${total_revenus_pdf:,.2f} vs Template=${total_revenus_template:,.2f}")
                        
                        # Check individual revenue accounts
                        for acc, row in [('Revenus mensuels', 13), ('Revenus Journaliers', 12), ('Revenus Lave-Auto', 14), ('Divers', 17), ('Gratuités - mensuels', 20)]:
                            pdf_val = pdf_data.get(acc)
                            template_val = ws.cell(row=row, column=col).value
                            if pdf_val is not None:
                                results.append(f"      {acc}: PDF=${pdf_val:,.2f} → Template Row {row}=${template_val}")
            else:
                results.append(f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net_pdf:,.2f}")
        elif benefice_net_pdf is None:
            results.append(f"⚠️ {month_name}: BÉNÉFICE NET not in PDF data")
        elif revenus_nets_template is None:
            results.append(f"⚠️ {month_name}: REVENUS NETS formula returned empty")
    
    return results

# ============================================
# MAIN FUNCTIONS (called by app.py)
# ============================================

def get_parking_codes_from_pnl(file_obj):
    """Extract parking codes from uploaded file"""
    codes = []
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
        file_obj.seek(0)
        if file_obj.name.lower().endswith('.pdf'):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(file_obj.read())
                    tmp_path = tmp.name
                doc = fitz.open(tmp_path)
                for page in doc:
                    text = page.get_text()
                    found = re.findall(r'(CMO\d+)', text, re.IGNORECASE)
                    codes.extend(found)
                doc.close()
                os.unlink(tmp_path)
            except:
                pass
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """
    MAIN FUNCTION
    1. Extract P&L data from PDFs
    2. Fill ONLY yellow cells in template
    3. Formulas auto-calculate totals
    4. Validate BÉNÉFICE NET = REVENUS NETS
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
        
        # Collect files
        all_files = []
        if monthly_files_current:
            all_files.extend(monthly_files_current)
        if monthly_files_previous:
            all_files.extend(monthly_files_previous)
        
        if not all_files:
            updates.append("\n⚠️ No monthly files provided!")
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue(), updates
        
        # Process each file
        updates.append(f"\n📁 Processing {len(all_files)} files...")
        
        for file_obj in all_files:
            month = extract_month_from_filename(file_obj)
            file_obj.seek(0)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            try:
                # Try digital first
                data = extract_from_digital_pdf(tmp_path)
                
                # Fall back to OCR if needed
                if data is None or len(data) < 3:
                    updates.append(f"   🔍 {month}: Trying OCR...")
                    data = extract_from_image_pdf(tmp_path)
                
                if data and len(data) > 0:
                    all_data[month] = data
                    rev = data.get('TOTAL REVENUS', 'N/A')
                    net = data.get('BÉNÉFICE NET', 'N/A')
                    updates.append(f"   ✅ {month}: {len(data)} accounts | PDF Revenue: ${rev} | PDF Net: ${net}")
                else:
                    updates.append(f"   ❌ {month}: No data extracted")
            finally:
                os.unlink(tmp_path)
        
        # Fill template (yellow cells only)
        if all_data:
            updates.append(f"\n📝 Filling YELLOW cells with {len(all_data)} months of data...")
            updates.append("   (Formula rows auto-calculate)")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            
            # Validate
            updates.append(f"\n🔍 Validating BÉNÉFICE NET = REVENUS NETS...")
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
