# excel_fixer.py - FINAL COMPLETE VERSION
"""
Extracts P&L data from PDF monthly reports and fills CMO111.xlsx template
Handles Canadian French number formats (113 648,89 $)
Works with both digital PDFs and image-based PDFs (OCR for January)
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

# Patterns to match account names (more specific first)
ACCOUNT_PATTERNS = [
    ('TOTAL REVENUS', ['total revenus', 'total des revenus']),
    ('BÉNÉFICE NET', ['bénéfice net', 'benefice net']),
    ("Total des frais d'exploitation", ['total des frais d\'exploitation', 'total des frais d\'exploit']),
    ("RÉSULTAT D'EXPLOITATION", ['résultat d\'exploitation', 'resultat d\'exploitation']),
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

# ============================================
# UTILITY FUNCTIONS
# ============================================

def safe_float(value):
    """
    Handle Canadian French number formats:
    "113 648,89 $" -> 113648.89
    "7 106 417,00" -> 7106417.00
    "(1 206,86) $" -> -1206.86
    "-1 206,86 $" -> -1206.86
    """
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    
    value = str(value).strip()
    
    # Handle parentheses (negative numbers)
    is_negative = False
    if value.startswith('(') and value.endswith(')'):
        is_negative = True
        value = value[1:-1].strip()
    
    # Remove leading/trailing $ signs
    value = value.replace('$', '').strip()
    
    # Remove ALL spaces (thousand separators in French Canadian)
    value = value.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    
    # Handle minus sign (before or after)
    if value.startswith('-'):
        is_negative = True
        value = value[1:]
    elif value.endswith('-'):
        is_negative = True
        value = value[:-1]
    
    # Handle comma as decimal separator (French)
    if ',' in value and '.' not in value:
        value = value.replace(',', '.')
    elif ',' in value and '.' in value:
        # Both present - comma is thousand, dot is decimal (or vice versa)
        # If dot is last and only 2 digits after, it's decimal
        if value.rfind('.') > value.rfind(','):
            # Dot is decimal, comma is thousand
            value = value.replace(',', '')
        else:
            # Comma is decimal, dot is thousand
            value = value.replace('.', '').replace(',', '.')
    
    # Remove any remaining non-numeric characters except minus and decimal
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
    """Match extracted text to a standard account name"""
    text_lower = text.lower().strip()
    # Clean up common OCR artifacts
    text_lower = text_lower.replace('|', '').replace('  ', ' ')
    
    for standard_name, patterns in ACCOUNT_PATTERNS:
        for pattern in patterns:
            if pattern in text_lower:
                return standard_name
    return None

# ============================================
# DIGITAL PDF EXTRACTION (Feb-Dec)
# ============================================

def find_pl_page_digital(pdf_path):
    """Find P&L page in digital PDF"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        matches = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if matches >= 3:
            return doc, i
    doc.close()
    return None, None

def extract_from_digital_pdf(pdf_path):
    """
    Extract using PyMuPDF text blocks.
    Text blocks keep numbers intact (e.g., "113 648,89 $" stays as one block).
    """
    doc, page_num = find_pl_page_digital(pdf_path)
    if doc is None:
        return None
    
    page = doc[page_num]
    
    # Get all text blocks with their positions
    blocks = page.get_text("blocks")
    
    # Sort blocks by vertical position (y), then horizontal (x)
    blocks.sort(key=lambda b: (round(b[1], 0), b[0]))
    
    # Filter to table area
    h = page.rect.height
    table_top = h * TABLE_TOP
    table_bottom = h * TABLE_BOTTOM
    table_blocks = [b for b in blocks if table_top <= b[1] <= table_bottom]
    
    # Group blocks by row (similar y position, rounded to nearest 5 points)
    rows = {}
    for block in table_blocks:
        y_key = round(block[1] / 5) * 5  # Group by 5-point intervals
        if y_key not in rows:
            rows[y_key] = []
        rows[y_key].append(block)
    
    # Sort rows by y position
    sorted_rows = sorted(rows.items())
    
    data = {}
    
    for y_key, row_blocks in sorted_rows:
        # Sort blocks in this row by x position (left to right)
        row_blocks.sort(key=lambda b: b[0])
        
        # First block is usually the account name
        if not row_blocks:
            continue
        
        account_block = row_blocks[0]
        account_text = account_block[4].strip()
        account_text = ' '.join(account_text.split())
        
        # Match to standard account name
        standard_name = match_account(account_text)
        if not standard_name:
            # Try combining first two blocks (sometimes account names span blocks)
            if len(row_blocks) >= 2:
                account_text2 = row_blocks[1][4].strip()
                account_text2 = ' '.join(account_text2.split())
                combined = account_text + ' ' + account_text2
                standard_name = match_account(combined)
            
            if not standard_name:
                continue
        
        # Second block is usually "Mois Courant" (the amount we want)
        # But skip if the second block is also text (like section headers)
        amount = None
        
        for block in row_blocks[1:]:
            block_text = block[4].strip()
            
            # Skip if this looks like a section header or text
            if match_account(block_text):
                continue
            
            # Try to extract a number
            amount = safe_float(block_text)
            if amount is not None:
                break
        
        if amount is not None:
            # Only keep the first occurrence of each account (avoid duplicates)
            if standard_name not in data:
                data[standard_name] = amount
    
    doc.close()
    return data

# ============================================
# IMAGE PDF EXTRACTION (January - OCR)
# ============================================

def find_pl_page_ocr(pdf_path):
    """Find P&L page in image PDF using OCR"""
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
        
        # Try French first, fallback to English
        try:
            text = pytesseract.image_to_string(img, lang='fra')
        except:
            try:
                text = pytesseract.image_to_string(img, lang='eng')
            except:
                text = pytesseract.image_to_string(img)
        
        # Check for P&L markers
        if "revenus" in text.lower() and ("mensuels" in text.lower() or "stationnement" in text.lower()):
            return doc, i, text
    
    doc.close()
    return None, None, None

def extract_from_image_pdf(pdf_path):
    """Extract from image-based PDF using OCR"""
    doc, page_num, ocr_text = find_pl_page_ocr(pdf_path)
    if doc is None:
        return None
    
    data = {}
    lines = ocr_text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Try to match an account name in this line
        standard_name = match_account(line)
        if not standard_name:
            continue
        
        # Look for dollar amounts in the line
        # Pattern: number with spaces, optional comma, optional $ sign
        amount_patterns = [
            r'(\d[\d\s]*,\d{2})\s*\$',  # "113 648,89 $"
            r'\$\s*(\d[\d\s]*,\d{2})',   # "$ 113 648,89"
            r'(\d[\d\s]*,\d{2})',        # "113 648,89"
            r'\(\s*(\d[\d\s]*,\d{2})\s*\)\s*\$',  # "(1 206,86) $"
        ]
        
        for pattern in amount_patterns:
            matches = re.findall(pattern, line)
            for match in matches:
                amount = safe_float(match)
                if amount is not None:
                    if standard_name not in data:
                        data[standard_name] = amount
                    break
            if standard_name in data:
                break
    
    doc.close()
    return data if data else None

# ============================================
# TEMPLATE FILLING
# ============================================

def fill_template(wb, all_data):
    """Fill '2-Données historiques' sheet with extracted data"""
    # Find the correct sheet
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
                continue
            
            # Write value to cell
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
    """Validate BÉNÉFICE NET (PDF) = REVENUS NETS (Template Row 86)"""
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
                results.append(
                    f"⚠️ {month_name}: BÉNÉFICE NET ${benefice_net:,.2f} ≠ "
                    f"REVENUS NETS ${template_val:,.2f} (diff: ${diff:,.2f})"
                )
            else:
                results.append(f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net:,.2f}")
        elif benefice_net is None:
            results.append(f"⚠️ {month_name}: BÉNÉFICE NET not extracted from PDF")
        elif template_val is None:
            results.append(f"⚠️ {month_name}: REVENUS NETS not found in template (Row 86)")
    
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
        
        # Also try to read from PDF content
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
    MAIN FUNCTION - Extract P&L data from PDFs and fill template.
    
    Process:
    1. Load Excel template
    2. For each PDF monthly report:
       a. Try digital extraction (text blocks)
       b. If fails, try OCR (for January/image PDFs)
    3. Fill "2-Données historiques" sheet with extracted data
    4. Validate BÉNÉFICE NET = REVENUS NETS
    5. Return updated Excel file
    
    Returns:
        (excel_bytes, updates_list)
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
        updates.append(f"   Sheets: {wb.sheetnames}")
        
        # Collect all files
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
            
            # Save to temp file for processing
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            try:
                # Try digital extraction first
                data = extract_from_digital_pdf(tmp_path)
                
                # If digital fails or returns too few accounts, try OCR
                if data is None or len(data) < 3:
                    updates.append(f"   🔍 {month}: Digital extraction insufficient, trying OCR...")
                    ocr_data = extract_from_image_pdf(tmp_path)
                    if ocr_data and len(ocr_data) > len(data or {}):
                        data = ocr_data
                
                if data and len(data) > 0:
                    all_data[month] = data
                    rev = data.get('TOTAL REVENUS', 'N/A')
                    net = data.get('BÉNÉFICE NET', 'N/A')
                    updates.append(f"   ✅ {month}: {len(data)} accounts | Revenue: ${rev} | Net Income: ${net}")
                    
                    # Show what accounts were found
                    for acc in ['Revenus mensuels', 'Revenus Journaliers', 'TOTAL REVENUS', 'BÉNÉFICE NET']:
                        if acc in data:
                            updates.append(f"      {acc}: ${data[acc]:,.2f}")
                else:
                    updates.append(f"   ❌ {month}: No data extracted")
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
        
        # Save to bytes
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("\n" + "=" * 60)
        updates.append("✅ WORKFLOW COMPLETE")
        updates.append("=" * 60)
        
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"\n❌ ERROR: {str(e)}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates


# ============================================
# STANDALONE TESTING
# ============================================
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        print(f"Testing extraction on: {pdf_path}")
        
        # Test digital extraction
        data = extract_from_digital_pdf(pdf_path)
        if data:
            print(f"\nDigital extraction: {len(data)} accounts")
            for k, v in data.items():
                print(f"  {k}: ${v:,.2f}")
        else:
            print("Digital extraction failed, trying OCR...")
            data = extract_from_image_pdf(pdf_path)
            if data:
                print(f"\nOCR extraction: {len(data)} accounts")
                for k, v in data.items():
                    print(f"  {k}: ${v:,.2f}")
            else:
                print("All extraction methods failed")
    else:
        print("Usage: python excel_fixer.py <pdf_file>")
