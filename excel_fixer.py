# excel_fixer.py - GRID-BASED EXTRACTION (FINAL FINAL)
import io
import re
import fitz
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
import os

# ============================================================================
# GRID CONFIGURATION (from February PDF analysis)
# ============================================================================

# Mois Courant column x-range (where the first number after account name appears)
MOIS_COURANT_X_MIN = 165
MOIS_COURANT_X_MAX = 195

# Row y-positions for each account (from PDF word positions)
ACCOUNT_ROWS = {
    164: 13,   # Revenus mensuels
    175: 12,   # Revenus Journaliers
    185: 14,   # Revenus Lave-Auto
    196: 17,   # Divers
    238: 20,   # Gratuités - mensuels
    302: 29,   # Salaires Stationnement
    313: 32,   # Uniformes
    323: 41,   # Fourn. de stationnement
    334: 35,   # Nettoyage
    345: 37,   # Equipement
    355: 36,   # Général
    366: 58,   # Taxes et permis
    376: 57,   # Assurances Cautionnement
    387: 56,   # Réclamations
    398: 50,   # Télécommunication
    408: 53,   # Frais de cartes de crédit
    419: 49,   # Frais de bureau
    482: 63,   # Honoraires de gestion
}

# Validation rows
VALIDATION_ROWS = {
    249: "_TOTAL_REVENUS_",
    430: "_TOTAL_EXPENSES_",
    514: "_BENEFICE_NET_",
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
# GRID-BASED EXTRACTION
# ============================================================================

def find_pl_page(doc):
    """Find the P&L page."""
    for page_num in range(len(doc)):
        text = doc[page_num].get_text()
        if "revenus mensuels" in text.lower() and "bénéfice net" in text.lower():
            return page_num
    return None

def extract_from_pdf(pdf_path, debug_updates=None):
    """Extract P&L data using fixed grid coordinates."""
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
    
    if debug_updates is not None:
        debug_updates.append(f"📐 Page {page_num+1}, {len(words)} words")
    
    # Group words by y-position (rounded to nearest integer)
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
    for y_pos, template_row in ACCOUNT_ROWS.items():
        # Find the y-key closest to our target
        closest_y = None
        for y_key in rows.keys():
            if abs(y_key - y_pos) <= 2:  # Within 2 points
                closest_y = y_key
                break
        
        if closest_y is None:
            if debug_updates is not None:
                debug_updates.append(f"  ❌ y={y_pos}: NO WORDS -> Row {template_row}")
            data[template_row] = 0.0
            continue
        
        # Get all words in Mois Courant x-range
        mois_courant_words = []
        for x_pos, texts in rows[closest_y].items():
            if MOIS_COURANT_X_MIN <= x_pos <= MOIS_COURANT_X_MAX:
                mois_courant_words.extend(texts)
        
        if mois_courant_words:
            # Combine words into a single number string
            combined = ' '.join(mois_courant_words)
            # Check for negative
            is_neg = combined.startswith('-') or combined.startswith('(')
            amount = parse_amount(combined)
            if amount is not None:
                if is_neg:
                    amount = -abs(amount)
                data[template_row] = amount
                if debug_updates is not None:
                    debug_updates.append(f"  ✅ y={y_pos}: ${amount:,.2f} -> Row {template_row}")
            else:
                data[template_row] = 0.0
                if debug_updates is not None:
                    debug_updates.append(f"  ⚠️ y={y_pos}: parse fail '{combined}' -> Row {template_row}")
        else:
            # No words in Mois Courant column = empty cell
            data[template_row] = 0.0
            if debug_updates is not None:
                debug_updates.append(f"  ⚠️ y={y_pos}: $0.00 (empty) -> Row {template_row}")
    
    # Extract validation accounts
    for y_pos, validation_key in VALIDATION_ROWS.items():
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
                    debug_updates.append(f"  📊 {validation_key} = ${amount:,.2f}")
    
    if debug_updates is not None:
        template_count = len([k for k in data if not str(k).startswith('_')])
        validation_count = len([k for k in data if str(k).startswith('_')])
        debug_updates.append(f"  📊 {template_count} template + {validation_count} validation")
    
    return data

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
            ws.cell(row=key, column=col).number_format = '#,##0.00'
            month_cells += 1
            total += 1
            updates.append(f"   ✅ {month_en} ({col_letter}{key}): ${amount:,.2f}")
        
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
                results.append(f"   ⚠️ PDF: ${pdf_ben:,.2f} | Expected: ${expected:,.2f} (diff: ${diff:,.2f})")
            else:
                results.append(f"   ✅ BÉNÉFICE NET = ${pdf_ben:,.2f}")
        elif pdf_ben is not None:
            results.append(f"   ⚠️ PDF Net: ${pdf_ben:,.2f}")
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
                    updates.append(f"   ✅ {month_en}: {cnt} accounts | PDF Net: ${ben}")
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
