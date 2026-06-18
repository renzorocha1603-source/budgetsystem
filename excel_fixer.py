# excel_fixer.py - DYNAMIC GRID (WORKS FOR ALL MONTHS)
import io
import re
import fitz
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
import os

# ============================================================================
# GRID CONFIGURATION
# ============================================================================

# Mois Courant column x-range (consistent across all PDFs)
MOIS_COURANT_X_MIN = 165
MOIS_COURANT_X_MAX = 195

# Account names to find (text search -> template row)
ACCOUNT_NAMES = {
    "revenus mensuels": 13,
    "revenus journaliers": 12,
    "revenus horaires": 12,
    "revenus lave-auto": 14,
    "divers": 17,
    "gratuités": 20,
    "salaires stationnement": 29,
    "salaire stationnement": 29,
    "uniformes": 32,
    "nettoyage": 35,
    "entretien réparation - nettoyage": 35,
    "général": 36,
    "general": 36,
    "entretien réparation - général": 36,
    "entretien stationnement": 36,
    "equipement": 37,
    "équipement": 37,
    "entretien réparation - equipement": 37,
    "fourn. de stationnement": 41,
    "fournitures stationnement": 41,
    "frais de bureau": 49,
    "télécommunication": 50,
    "telecommunication": 50,
    "frais de cartes de crédit": 53,
    "frais de cartes de credit": 53,
    "réclamations": 56,
    "reclamations": 56,
    "assurances": 57,
    "assurances cautionnement": 57,
    "taxes et permis": 58,
    "honoraires de gestion": 63,
}

VALIDATION_NAMES = {
    "total revenus": "_TOTAL_REVENUS_",
    "total des frais d'exploitation": "_TOTAL_EXPENSES_",
    "bénéfice net": "_BENEFICE_NET_",
    "benefice net": "_BENEFICE_NET_",
}

# ============================================================================
# CANADIAN FRENCH NUMBER PARSING
# ============================================================================

def parse_amount(text):
    if not text:
        return None
    text = str(text).strip()
    is_negative = False
    if text.startswith('(') and text.endswith(')'):
        is_negative = True
        text = text[1:-1].strip()
    text = text.replace('$', '').strip()
    text = text.replace(' ', '').replace('\xa0', '').replace('\u202f', '')
    if text.startswith('-'):
        is_negative = True
        text = text[1:]
    if ',' in text:
        text = text.replace(',', '.')
    text = re.sub(r'[^\d\.\-]', '', text)
    try:
        value = float(text)
        return -value if is_negative else value
    except:
        return None

# ============================================================================
# MONTH DETECTION
# ============================================================================

MONTH_MAP = {
    'janvier': 'January', 'février': 'February', 'fevrier': 'February',
    'mars': 'March', 'avril': 'April', 'mai': 'May', 'juin': 'June',
    'juillet': 'July', 'août': 'August', 'aout': 'August',
    'septembre': 'September', 'octobre': 'October',
    'novembre': 'November', 'décembre': 'December', 'decembre': 'December'
}

MONTH_COLUMN = {
    'January': 2, 'February': 3, 'March': 4, 'April': 5,
    'May': 6, 'June': 7, 'July': 8, 'August': 9,
    'September': 10, 'October': 11, 'November': 12, 'December': 13
}

# ============================================================================
# DYNAMIC GRID EXTRACTION
# ============================================================================

def find_pl_page(doc):
    for page_num in range(len(doc)):
        text = doc[page_num].get_text()
        if "revenus" in text.lower() and "bénéfice net" in text.lower():
            return page_num
    return None

def extract_by_dynamic_grid(pdf_path, debug_updates=None):
    """Find accounts by text, extract Mois Courant by x-position."""
    doc = fitz.open(pdf_path)
    page_num = find_pl_page(doc)
    
    if page_num is None:
        doc.close()
        if debug_updates is not None:
            debug_updates.append("❌ P&L page not found")
        return {}
    
    page = doc[page_num]
    words = page.get_text("words")
    doc.close()
    
    if not words or len(words) < 50:
        if debug_updates is not None:
            debug_updates.append(f"⚠️ Only {len(words)} words - trying OCR")
        return {}
    
    if debug_updates is not None:
        debug_updates.append(f"📐 Page {page_num+1}, {len(words)} words")
    
    # Build map: account_name -> y_position
    account_y = {}
    for w in words:
        x0, y0, x1, y1, text, block, line, word_no = w
        text_lower = text.lower().strip()
        
        # Only look at first column (x < 200)
        if x0 > 200:
            continue
        
        for search_term in ACCOUNT_NAMES:
            if search_term in text_lower and search_term not in account_y:
                account_y[search_term] = int(y0)
                break
        
        for search_term in VALIDATION_NAMES:
            if search_term in text_lower and search_term not in account_y:
                account_y[search_term] = int(y0)
                break
    
    if debug_updates is not None:
        debug_updates.append(f"🔍 Found {len(account_y)} account positions")
    
    # Group words by y-position
    rows = {}
    for w in words:
        x0, y0, x1, y1, text, block, line, word_no = w
        y_key = int(y0)
        if y_key not in rows:
            rows[y_key] = {}
        if x0 not in rows[y_key]:
            rows[y_key][x0] = []
        rows[y_key][x0].append(text)
    
    data = {}
    
    # Extract template accounts
    for search_term, template_row in ACCOUNT_NAMES.items():
        if template_row in data:
            continue
        
        y_pos = account_y.get(search_term)
        if y_pos is None:
            data[template_row] = 0.0
            continue
        
        # Find closest y-key in rows
        closest_y = None
        for y_key in rows.keys():
            if abs(y_key - y_pos) <= 2:
                closest_y = y_key
                break
        
        if closest_y is None:
            data[template_row] = 0.0
            if debug_updates is not None:
                debug_updates.append(f"  ❌ {search_term}: no row at y={y_pos}")
            continue
        
        # Get Mois Courant value from x-range 165-195
        mois_courant_words = []
        for x_pos, texts in rows[closest_y].items():
            if MOIS_COURANT_X_MIN <= x_pos <= MOIS_COURANT_X_MAX:
                mois_courant_words.extend(texts)
        
        if mois_courant_words:
            combined = ' '.join(mois_courant_words)
            is_neg = combined.startswith('-') or combined.startswith('(')
            amount = parse_amount(combined)
            if amount is not None:
                if is_neg:
                    amount = -abs(amount)
                data[template_row] = amount
                if debug_updates is not None:
                    debug_updates.append(f"  ✅ {search_term}: {amount:,.2f} $ -> Row {template_row}")
            else:
                data[template_row] = 0.0
        else:
            data[template_row] = 0.0
            if debug_updates is not None:
                debug_updates.append(f"  ⚠️ {search_term}: $0.00 -> Row {template_row}")
    
    # Extract validation accounts
    for search_term, validation_key in VALIDATION_NAMES.items():
        if validation_key in data:
            continue
        
        y_pos = account_y.get(search_term)
        if y_pos is None:
            continue
        
        closest_y = None
        for y_key in rows.keys():
            if abs(y_key - y_pos) <= 2:
                closest_y = y_key
                break
        
        if closest_y is None:
            continue
        
        mois_courant_words = []
        for x_pos, texts in rows[closest_y].items():
            if MOIS_COURANT_X_MIN <= x_pos <= MOIS_COURANT_X_MAX:
                mois_courant_words.extend(texts)
        
        if mois_courant_words:
            combined = ' '.join(mois_courant_words)
            amount = parse_amount(combined)
            if amount is not None:
                data[validation_key] = amount
                if debug_updates is not None:
                    debug_updates.append(f"  📊 {validation_key} = {amount:,.2f} $")
    
    if debug_updates is not None:
        template_count = len([k for k in data if not str(k).startswith('_')])
        validation_count = len([k for k in data if str(k).startswith('_')])
        debug_updates.append(f"  📊 {template_count} template + {validation_count} validation")
    
    return data

# ============================================================================
# OCR EXTRACTION (January fallback)
# ============================================================================

def is_number_line(text):
    text = text.strip()
    if not text:
        return False
    return bool(re.match(r'^[\d\s,.\-()$]+$', text))

def extract_by_ocr(pdf_path, debug_updates=None):
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        if debug_updates is not None:
            debug_updates.append("❌ Tesseract not installed")
        return {}
    
    doc = fitz.open(pdf_path)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        try:
            text = pytesseract.image_to_string(img, lang='fra')
        except:
            try:
                text = pytesseract.image_to_string(img, lang='eng')
            except:
                text = pytesseract.image_to_string(img)
        
        if "revenus" in text.lower() and len(text) > 200:
            if debug_updates is not None:
                debug_updates.append(f"📝 OCR: {len(text)} chars from page {page_num+1}")
            
            lines = text.split('\n')
            clean_lines = [l.strip() for l in lines]
            data = {}
            
            for search_term, template_row in ACCOUNT_NAMES.items():
                if template_row in data:
                    continue
                
                for i, line in enumerate(clean_lines):
                    if search_term in line.lower():
                        if i + 1 < len(clean_lines):
                            next_line = clean_lines[i + 1]
                            if next_line and is_number_line(next_line):
                                amount = parse_amount(next_line)
                                if amount is not None:
                                    data[template_row] = amount
                        break
                
                if template_row not in data:
                    data[template_row] = 0.0
            
            doc.close()
            
            if len([k for k in data if not str(k).startswith('_')]) >= 5:
                return data
    
    doc.close()
    return {}

# ============================================================================
# MAIN EXTRACTION
# ============================================================================

def extract_from_pdf(pdf_path, debug_updates=None):
    # Try dynamic grid first
    data = extract_by_dynamic_grid(pdf_path, debug_updates)
    template_count = len([k for k in data if not str(k).startswith('_')])
    
    if template_count >= 10:
        return data
    
    # Fall back to OCR
    if debug_updates is not None:
        debug_updates.append(f"🔍 Grid got {template_count}, trying OCR...")
    
    ocr_data = extract_by_ocr(pdf_path, debug_updates)
    ocr_count = len([k for k in ocr_data if not str(k).startswith('_')])
    
    if ocr_count > template_count:
        return ocr_data
    
    return data if template_count >= 5 else ocr_data

# ============================================================================
# TEMPLATE FILLING
# ============================================================================

def fill_template(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        return ["❌ Sheet not found"]
    
    ws = wb[sheet_name]
    updates = []
    total = 0
    
    for month_en, month_data in all_data.items():
        col = MONTH_COLUMN.get(month_en)
        if col is None:
            continue
        
        col_letter = get_column_letter(col)
        month_cells = 0
        
        for key, amount in month_data.items():
            if str(key).startswith('_'):
                continue
            ws.cell(row=key, column=col).value = amount
            ws.cell(row=key, column=col).number_format = '#,##0.00 $'
            month_cells += 1
            total += 1
            updates.append(f"   ✅ {month_en} ({col_letter}{key}): {amount:,.2f} $")
        
        if month_cells > 0:
            updates.append(f"📊 {month_en}: {month_cells} cells")
    
    updates.append(f"\n📊 TOTAL: {total} yellow cells (formulas auto-calculate)")
    return updates

# ============================================================================
# VALIDATION
# ============================================================================

def validate_results(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        return ["⚠️ Cannot validate"]
    
    results = []
    
    for month_en, month_data in all_data.items():
        col = MONTH_COLUMN.get(month_en)
        if col is None:
            continue
        
        pdf_ben = month_data.get('_BENEFICE_NET_')
        pdf_rev = month_data.get('_TOTAL_REVENUS_')
        pdf_exp = month_data.get('_TOTAL_EXPENSES_')
        
        results.append(f"\n📊 {month_en}:")
        
        if pdf_ben is not None and pdf_rev is not None and pdf_exp is not None:
            expected = pdf_rev - pdf_exp
            diff = abs(pdf_ben - expected)
            if diff > 0.01:
                results.append(f"   ⚠️ PDF: {pdf_ben:,.2f} $ | Expected: {expected:,.2f} $ (diff: {diff:,.2f} $)")
            else:
                results.append(f"   ✅ BÉNÉFICE NET = {pdf_ben:,.2f} $")
        elif pdf_ben is not None:
            results.append(f"   ⚠️ PDF Net: {pdf_ben:,.2f} $")
        else:
            results.append(f"   ⚠️ Not found")
    
    return results

# ============================================================================
# APP.PY INTERFACE
# ============================================================================

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
    debug = []
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append("=" * 60)
        
        if isinstance(excel_file, bytes):
            wb = load_workbook(io.BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(io.BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template: {parking_code or 'Unknown'}")
        
        all_files = []
        if monthly_files_current:
            all_files.extend(monthly_files_current)
        if monthly_files_previous:
            all_files.extend(monthly_files_previous)
        
        if not all_files:
            updates.append("⚠️ No files!")
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue(), updates
        
        updates.append(f"\n📁 {len(all_files)} files...")
        
        for file_obj in all_files:
            month_en = None
            if hasattr(file_obj, 'name'):
                for fr, en in MONTH_MAP.items():
                    if fr in file_obj.name.lower():
                        month_en = en
                        break
            
            if month_en is None:
                updates.append(f"⚠️ Unknown month")
                continue
            
            file_obj.seek(0)
            file_bytes = file_obj.read()
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            
            try:
                debug.append(f"--- {month_en} ---")
                data = extract_from_pdf(tmp_path, debug)
                
                if data:
                    cnt = len([k for k in data if not str(k).startswith('_')])
                    all_data[month_en] = data
                    ben = data.get('_BENEFICE_NET_', 'N/A')
                    updates.append(f"   ✅ {month_en}: {cnt} accounts | PDF Net: {ben} $")
                else:
                    updates.append(f"   ❌ {month_en}: No data")
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
        
        if debug:
            updates.append("\n🔍 Debug:")
            updates.extend(debug)
        
        if all_data:
            updates.append(f"\n📝 Filling {len(all_data)} months...")
            updates.extend(fill_template(wb, all_data))
            updates.append(f"\n🔍 Validation:")
            updates.extend(validate_results(wb, all_data))
        else:
            updates.append("\n⚠️ No data!")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        updates.append("\n✅ DONE")
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"\n❌ {e}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates
