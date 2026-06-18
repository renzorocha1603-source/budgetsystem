# excel_fixer.py - COORDINATE-BASED + ALLISON FALLBACK (FINAL)
import io
import re
import json
import requests
import pandas as pd
import fitz
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import tempfile
import os

# ============================================================================
# MISTRAL CONFIGURATION (for Allison fallback)
# ============================================================================
MISTRAL_API_KEY = "em5oqjSdA1Nus9iUpa1MNAJtQA4YfCtK"
MISTRAL_URL = "https://api.mistral.ai/v1/chat/completions"
MISTRAL_MODEL = "mistral-small-latest"

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
# ACCOUNT MAPPING
# ============================================================================

# Maps French account names found in PDF to template rows
ACCOUNT_MAP = {
    "revenus mensuels": 13,
    "revenus journaliers": 12,
    "revenus horaires": 12,
    "revenus lave-auto": 14,
    "divers": 17,
    "gratuités - mensuels": 20,
    "gratuités": 20,
    "salaires stationnement": 29,
    "salaire stationnement": 29,
    "uniformes": 32,
    "nettoyage": 35,
    "entretien réparation - nettoyage": 35,
    "entretien réparation - général": 36,
    "entretien réparation - general": 36,
    "entretien stationnement": 36,
    "entretien réparation - equipement": 37,
    "entretien réparation - équipement": 37,
    "fourn. de stationnement": 41,
    "fournitures stationnement": 41,
    "frais de bureau": 49,
    "télécommunication": 50,
    "telecommunication": 50,
    "frais de cartes de crédit": 53,
    "frais de cartes de credit": 53,
    "réclamations": 56,
    "reclamations": 56,
    "assurances cautionnement": 57,
    "assurances": 57,
    "taxes et permis": 58,
    "honoraires de gestion": 63,
}

# Validation accounts to capture for comparison
VALIDATION_ACCOUNTS = {
    "total revenus": "_TOTAL_REVENUS_",
    "total des revenus": "_TOTAL_REVENUS_",
    "total des frais d'exploitation": "_TOTAL_EXPENSES_",
    "bénéfice net": "_BENEFICE_NET_",
    "benefice net": "_BENEFICE_NET_",
}

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
# COORDINATE-BASED EXTRACTION (PRIMARY METHOD)
# ============================================================================

def find_pl_page(doc):
    """Find the P&L page by content markers."""
    for page_num in range(len(doc)):
        text = doc[page_num].get_text()
        if "revenus mensuels" in text.lower() and "bénéfice net" in text.lower():
            return page_num
    return None

def extract_by_coordinates(pdf_path, debug_updates=None):
    """Extract P&L data using x-positions to identify Mois Courant column."""
    doc = fitz.open(pdf_path)
    page_num = find_pl_page(doc)
    
    if page_num is None:
        doc.close()
        if debug_updates is not None:
            debug_updates.append("❌ P&L page not found")
        return None
    
    page = doc[page_num]
    words = page.get_text("words")
    
    if not words:
        doc.close()
        return None
    
    # Calculate column boundaries from x-positions
    all_x = [w[0] for w in words]
    min_x = min(all_x)
    max_x = max(all_x)
    col_width = (max_x - min_x) / 9
    
    if debug_updates is not None:
        debug_updates.append(f"📐 X range: {min_x:.0f}-{max_x:.0f}, col width: {col_width:.0f}")
    
    # Group words by row (y-position rounded to nearest 10)
    rows = {}
    for w in words:
        x0, y0, x1, y1, text, block, line, word_no = w
        y_key = round(y0 / 10) * 10
        if y_key not in rows:
            rows[y_key] = []
        rows[y_key].append((x0, text))
    
    data = {}
    
    for y_key in sorted(rows.keys()):
        row_words = sorted(rows[y_key], key=lambda item: item[0])
        
        # Get account name from column 0
        account_text = ""
        for x, text in row_words:
            col_idx = int((x - min_x) / col_width)
            if col_idx == 0:
                account_text += " " + text
        
        account_text = account_text.strip().lower()
        if not account_text:
            continue
        
        # Find matching template row
        template_row = None
        for search_term, row_num in ACCOUNT_MAP.items():
            if search_term in account_text:
                template_row = row_num
                break
        
        # Check for validation accounts
        validation_key = None
        for search_term, vk in VALIDATION_ACCOUNTS.items():
            if search_term in account_text:
                validation_key = vk
                break
        
        if template_row is None and validation_key is None:
            continue
        
        # Get Mois Courant value from column 1
        mois_courant_value = None
        for x, text in row_words:
            col_idx = int((x - min_x) / col_width)
            if col_idx == 1:  # Mois Courant column
                mois_courant_value = parse_amount(text)
                break
        
        # If no word in column 1, cell is empty → $0.00
        if mois_courant_value is None:
            mois_courant_value = 0.0
        
        if template_row is not None:
            data[template_row] = mois_courant_value
            if debug_updates is not None:
                debug_updates.append(f"  ✅ {account_text[:40]}: ${mois_courant_value:,.2f} -> Row {template_row}")
        
        if validation_key is not None:
            data[validation_key] = mois_courant_value
            if debug_updates is not None:
                debug_updates.append(f"  📊 {validation_key} = ${mois_courant_value:,.2f}")
    
    doc.close()
    
    if debug_updates is not None:
        template_count = len([k for k in data if not str(k).startswith('_')])
        validation_count = len([k for k in data if str(k).startswith('_')])
        debug_updates.append(f"  📊 Total: {template_count} template + {validation_count} validation")
    
    return data if len(data) >= 10 else None

# ============================================================================
# ALLISON AI FALLBACK
# ============================================================================

def extract_with_allison(pdf_path, debug_updates=None):
    """Use Allison (Mistral) to extract P&L data as fallback."""
    if debug_updates is not None:
        debug_updates.append("🤖 Trying Allison AI extraction...")
    
    try:
        doc = fitz.open(pdf_path)
        pdf_text = ""
        for page_num in range(len(doc)):
            pdf_text += doc[page_num].get_text("text") + "\n"
        doc.close()
        
        prompt = f"""Extract P&L data from this PDF text. Return ONLY valid JSON.

The table has 9 columns: Account | Mois Courant | Budget | Écart | An Préc | Cumulatif | Cumul budget | Écart cumul | An Préc cumul

We ONLY want Mois Courant (Column 1). The PDF text has numbers on separate lines after each account name. The FIRST number after the account name is Mois Courant. If there's no number on the next line, Mois Courant = 0.

Account mapping (template_row: account_name):
12: Revenus Journaliers/Horaires
13: Revenus mensuels
14: Revenus Lave-Auto
17: Divers
20: Gratuités - mensuels
29: Salaires Stationnement
32: Uniformes
35: Nettoyage
36: Général/Entretien stationnement
37: Equipement
41: Fourn. de stationnement
49: Frais de bureau
50: Télécommunication
53: Frais de cartes de crédit
56: Réclamations
57: Assurances Cautionnement
58: Taxes et permis
63: Honoraires de gestion

Return: {{"12": 71064.17, "13": 43585.46, ...}}

PDF TEXT:
{pdf_text[:8000]}"""
        
        resp = requests.post(
            MISTRAL_URL,
            headers={"Authorization": f"Bearer {MISTRAL_API_KEY}", "Content-Type": "application/json"},
            json={"model": MISTRAL_MODEL, "messages": [{"role": "user", "content": prompt}], "temperature": 0.1},
            timeout=45
        )
        
        if resp.status_code == 200:
            response_text = resp.json()["choices"][0]["message"]["content"]
            # Parse JSON
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                allison_data = json.loads(json_match.group())
                data = {}
                for key_str, amount in allison_data.items():
                    data[int(key_str)] = float(amount)
                if debug_updates is not None:
                    debug_updates.append(f"✅ Allison extracted {len(data)} accounts")
                return data if len(data) >= 10 else None
    except Exception as e:
        if debug_updates is not None:
            debug_updates.append(f"❌ Allison error: {e}")
    
    return None

# ============================================================================
# MAIN EXTRACTION
# ============================================================================

def extract_from_pdf(pdf_path, debug_updates=None):
    """Extract P&L data: coordinates first, Allison fallback."""
    
    # Method 1: Coordinate-based extraction
    if debug_updates is not None:
        debug_updates.append("🔍 Using coordinate-based extraction...")
    
    data = extract_by_coordinates(pdf_path, debug_updates)
    if data and len(data) >= 10:
        return data
    
    # Method 2: Allison AI
    data = extract_with_allison(pdf_path, debug_updates)
    if data and len(data) >= 10:
        return data
    
    # Method 3: Return empty
    if debug_updates is not None:
        debug_updates.append("❌ All extraction methods failed")
    
    return {}

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
        return ["❌ Sheet '2-Données historiques' not found"]
    
    ws = wb[sheet_name]
    updates = []
    total_cells = 0
    
    for month_en, month_data in all_data.items():
        col = MONTH_COLUMN.get(month_en)
        if col is None:
            updates.append(f"⚠️ Unknown month: {month_en}")
            continue
        
        col_letter = get_column_letter(col)
        month_cells = 0
        
        for key, amount in month_data.items():
            if str(key).startswith('_'):
                continue
            cell = ws.cell(row=key, column=col)
            cell.value = amount
            cell.number_format = '#,##0.00'
            month_cells += 1
            total_cells += 1
            updates.append(f"   ✅ {month_en} ({col_letter}{key}): ${amount:,.2f}")
        
        if month_cells > 0:
            updates.append(f"📊 {month_en}: {month_cells} cells filled")
    
    updates.append(f"\n📊 TOTAL: {total_cells} yellow cells filled (formulas auto-calculate)")
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
        return ["⚠️ Cannot validate - sheet not found"]
    
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
            expected_net = pdf_rev - pdf_exp
            diff = abs(pdf_ben - expected_net)
            if diff > 0.01:
                results.append(f"   ⚠️ PDF BÉNÉFICE NET: ${pdf_ben:,.2f}")
                results.append(f"   📈 PDF Revenue: ${pdf_rev:,.2f} - PDF Expenses: ${pdf_exp:,.2f}")
                results.append(f"   📊 Expected Net: ${expected_net:,.2f} (diff: ${diff:,.2f})")
            else:
                results.append(f"   ✅ BÉNÉFICE NET = ${pdf_ben:,.2f}")
                results.append(f"   📈 Revenue: ${pdf_rev:,.2f} | Expenses: ${pdf_exp:,.2f} | Net: ${pdf_ben:,.2f}")
        elif pdf_ben is not None:
            results.append(f"   ⚠️ PDF BÉNÉFICE NET: ${pdf_ben:,.2f} (missing totals)")
        else:
            results.append(f"   ⚠️ BÉNÉFICE NET not found in PDF")
    
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
    debug_updates = []
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append("=" * 60)
        
        if isinstance(excel_file, bytes):
            wb = load_workbook(io.BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(io.BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
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
        
        updates.append(f"\n📁 Processing {len(all_files)} files...")
        
        for file_obj in all_files:
            month_en = None
            if hasattr(file_obj, 'name'):
                name_lower = file_obj.name.lower()
                for fr, en in MONTH_MAP.items():
                    if fr in name_lower:
                        month_en = en
                        break
            
            if month_en is None:
                updates.append(f"⚠️ Could not determine month for {getattr(file_obj, 'name', 'unknown')}")
                continue
            
            file_obj.seek(0)
            file_bytes = file_obj.read()
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            
            try:
                debug_updates.append(f"--- {month_en} ---")
                data = extract_from_pdf(tmp_path, debug_updates)
                
                if data:
                    display_count = len([k for k in data if not str(k).startswith('_')])
                    all_data[month_en] = data
                    pdf_benefice = data.get('_BENEFICE_NET_', 'N/A')
                    updates.append(f"   ✅ {month_en}: {display_count} accounts | PDF BÉNÉFICE NET: ${pdf_benefice}")
                else:
                    updates.append(f"   ❌ {month_en}: No data extracted")
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
        
        if debug_updates:
            updates.append("\n🔍 DEBUG:")
            updates.extend(debug_updates)
        
        if all_data:
            updates.append(f"\n📝 Filling YELLOW cells for {len(all_data)} months...")
            updates.append("   (Formula rows auto-calculate)")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            
            updates.append(f"\n🔍 Validation (PDF BÉNÉFICE NET vs Expected):")
            validations = validate_results(wb, all_data)
            updates.extend(validations)
        else:
            updates.append("\n⚠️ No data extracted from any file!")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("\n✅ WORKFLOW COMPLETE")
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"\n❌ ERROR: {str(e)}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates
