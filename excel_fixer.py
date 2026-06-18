# excel_fixer.py - COMPLETE WITH DIAGNOSTIC
"""
Excel Fixer Module - PDF to Excel conversion + Template filling
Only Solutions Inc. - Budget System
Includes built-in diagnostic mode
"""
import pandas as pd
import os
import re
from datetime import datetime
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO
import tempfile
from pathlib import Path

# ============================================
# HARDCODED SPATIAL LAYOUT - P&L PAGE
# ============================================
COLUMN_WIDTHS = [0.22, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10]
AMOUNT_COL = 1  # Mois Courant - ALWAYS column 1
NUM_ROWS = 27
TABLE_TOP = 0.08
TABLE_BOTTOM = 0.95
TABLE_LEFT = 0.03
TABLE_RIGHT = 0.97

# Content markers to identify P&L page
PAGE_MARKERS = [
    "1981 McGill College",
    "Revenus mensuels", 
    "BÉNÉFICE NET",
    "Mois Courant"
]

# ============================================
# P&L ROW MAPPING (Row number -> French account name)
# ============================================
PAGE10_ROW_MAPPING = {
    1: 'Revenus mensuels',
    2: 'Revenus Journaliers',
    3: 'Revenus Lave-Auto',
    4: 'Divers',
    5: 'Revenus de stationnement',
    6: 'Gratuités - mensuels',
    7: 'TOTAL REVENUS',
    8: 'DÉPENSES',
    9: "DÉPENSES D'EXPLOITATION",
    10: 'Salaires Stationnement',
    11: 'Uniformes',
    12: 'Fourn. de stationnement',
    13: 'Entretien réparation - Nettoyage',
    14: 'Entretien réparation - Equipement',
    15: 'Entretien réparation - Général',
    16: 'Taxes et permis',
    17: 'Assurances Cautionnement',
    18: 'Réclamations',
    19: 'Télécommunication',
    20: 'Frais de cartes de crédit',
    21: 'Frais de bureau',
    22: "Total des frais d'exploitation",
    23: "RÉSULTAT D'EXPLOITATION",
    24: 'AUTRES FRAIS',
    25: 'Honoraires de gestion',
    26: 'Total des autres frais',
    27: 'BÉNÉFICE NET'
}

# ============================================
# TEMPLATE MAPPING
# ============================================
MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

TEMPLATE_ROW_MAP = {
    'Revenus mensuels': 11,
    'Revenus Journaliers': 9,
    'Revenus Lave-Auto': 12,
    'Divers': 15,
    'Gratuités - mensuels': 18,
    'TOTAL REVENUS': 24,
    'Salaires Stationnement': 27,
    'Uniformes': 30,
    'Entretien réparation - Nettoyage': 34,
    'Entretien réparation - Equipement': 36,
    'Entretien réparation - Général': 35,
    'Fourn. de stationnement': 40,
    'Taxes et permis': 54,
    'Assurances Cautionnement': 53,
    'Réclamations': 52,
    'Télécommunication': 47,
    'Frais de cartes de crédit': 49,
    'Frais de bureau': 46,
    "Total des frais d'exploitation": 66,
    'Honoraires de gestion': 56,
    'BÉNÉFICE NET': 68,
}

# ============================================
# UTILITY FUNCTIONS
# ============================================

def safe_float(value, debug=False):
    """Convert European number format to float."""
    if value is None or value == "":
        if debug: print(f"        safe_float: empty input")
        return None
    if isinstance(value, (int, float)):
        return float(value)
    
    original = str(value).strip()
    if debug: print(f"        safe_float input: '{original}'")
    
    is_negative = original.startswith('(') and original.endswith(')')
    cleaned = original
    if is_negative:
        cleaned = cleaned[1:-1]
    
    cleaned = cleaned.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    if ',' in cleaned:
        cleaned = cleaned.replace(',', '.')
    cleaned = re.sub(r'[^\d\.\-]', '', cleaned)
    
    try:
        result = float(cleaned)
        if is_negative:
            result = -result
        if debug: print(f"        safe_float output: {result}")
        return result
    except:
        if debug: print(f"        safe_float FAILED")
        return None

def extract_month_from_filename(filename):
    """Extract month name from filename"""
    months = {
        'janvier': 'Janvier', 'février': 'Février', 'mars': 'Mars',
        'avril': 'Avril', 'mai': 'Mai', 'juin': 'Juin',
        'juillet': 'Juillet', 'août': 'Août', 'septembre': 'Septembre',
        'octobre': 'Octobre', 'novembre': 'Novembre', 'décembre': 'Décembre'
    }
    
    if hasattr(filename, 'name'):
        name = filename.name.lower()
    else:
        name = str(filename).lower()
    
    for key, value in months.items():
        if key in name:
            return value
    
    match = re.search(
        r'(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)',
        name, re.IGNORECASE
    )
    if match:
        return match.group(1).capitalize()
    
    return None

# ============================================
# DIAGNOSTIC FUNCTION (built-in)
# ============================================

def run_diagnostic(pdf_path):
    """
    DIAGNOSTIC MODE: Test multiple Y offsets and show ALL rows.
    Finds the correct offset for table extraction.
    """
    print("\n" + "="*80)
    print("🔍 DIAGNOSTIC MODE")
    print("="*80)
    print(f"File: {os.path.basename(pdf_path)}")
    
    doc = fitz.open(pdf_path)
    print(f"Total pages: {len(doc)}")
    
    # Find P&L page
    pl_page = None
    pl_page_num = None
    
    for i in range(len(doc)):
        text = doc[i].get_text()
        # Check for markers
        markers_found = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if markers_found >= 3:
            pl_page = doc[i]
            pl_page_num = i
            print(f"✅ P&L found on page {i+1} ({markers_found}/{len(PAGE_MARKERS)} markers)")
            break
    
    if pl_page is None:
        print("❌ P&L page not found!")
        print("\nShowing text from each page:")
        for i in range(len(doc)):
            text = doc[i].get_text()
            print(f"\n--- Page {i+1} ---")
            print(text[:300])
        doc.close()
        return None
    
    w = pl_page.rect.width
    h = pl_page.rect.height
    print(f"Page size: {w:.0f} x {h:.0f} points")
    
    # Test different Y offsets
    best_offset = None
    best_score = 0
    
    for y_offset in [0.05, 0.06, 0.07, 0.08, 0.09, 0.10, 0.12, 0.15, 0.20]:
        print(f"\n{'='*60}")
        print(f"Testing Y offset: {y_offset} (top = {h*y_offset:.0f} pts)")
        print(f"{'='*60}")
        
        top = h * y_offset
        bottom = h * TABLE_BOTTOM
        left = w * TABLE_LEFT
        right = w * TABLE_RIGHT
        table_w = right - left
        row_h = (bottom - top) / NUM_ROWS
        
        found_accounts = 0
        found_total_revenus = False
        found_benefice_net = False
        found_total_expenses = False
        
        for row_idx in range(NUM_ROWS):
            y = top + (row_idx * row_h)
            x = left
            
            # Column 0 - Account name
            cw = COLUMN_WIDTHS[0] * table_w
            account = pl_page.get_text("text", clip=fitz.Rect(x+1, y+1, x+cw-1, y+row_h-1))
            account = ' '.join(account.strip().split())
            
            # Column 1 - Amount
            x += cw
            cw = COLUMN_WIDTHS[1] * table_w
            amount = pl_page.get_text("text", clip=fitz.Rect(x+1, y+1, x+cw-1, y+row_h-1))
            amount = ' '.join(amount.strip().split())
            
            # Check for key accounts
            if "TOTAL REVENUS" in account.upper():
                found_total_revenus = True
                found_accounts += 1
                print(f"  ⭐ Row {row_idx+1}: TOTAL REVENUS = '{amount}'")
            elif "BÉNÉFICE NET" in account.upper():
                found_benefice_net = True
                found_accounts += 1
                print(f"  ⭐ Row {row_idx+1}: BÉNÉFICE NET = '{amount}'")
            elif "Total des frais" in account or "TOTAL DÉPENSES" in account.upper():
                found_total_expenses = True
                found_accounts += 1
                print(f"  ⭐ Row {row_idx+1}: TOTAL EXPENSES = '{amount}'")
            elif account and row_idx < 10:
                print(f"  Row {row_idx+1}: '{account[:50]}' = '{amount}'")
        
        score = found_accounts
        if found_total_revenus and found_benefice_net and found_total_expenses:
            score += 10
        
        print(f"  Score: {score} (Total Rev: {found_total_revenus}, Bén Net: {found_benefice_net}, Total Exp: {found_total_expenses})")
        
        if score > best_score:
            best_score = score
            best_offset = y_offset
    
    print(f"\n{'='*60}")
    print(f"🏆 BEST Y OFFSET: {best_offset} (score: {best_score})")
    print(f"   Update TABLE_TOP = {best_offset} in excel_fixer.py")
    print(f"{'='*60}")
    
    # Now show full extraction with best offset
    if best_offset:
        print(f"\n{'='*60}")
        print(f"FULL EXTRACTION WITH BEST OFFSET ({best_offset})")
        print(f"{'='*60}")
        
        top = h * best_offset
        bottom = h * TABLE_BOTTOM
        left = w * TABLE_LEFT
        right = w * TABLE_RIGHT
        table_w = right - left
        row_h = (bottom - top) / NUM_ROWS
        
        for row_idx in range(NUM_ROWS):
            y = top + (row_idx * row_h)
            x = left
            
            cw = COLUMN_WIDTHS[0] * table_w
            account = pl_page.get_text("text", clip=fitz.Rect(x+1, y+1, x+cw-1, y+row_h-1))
            account = ' '.join(account.strip().split())
            
            x += cw
            cw = COLUMN_WIDTHS[1] * table_w
            amount = pl_page.get_text("text", clip=fitz.Rect(x+1, y+1, x+cw-1, y+row_h-1))
            amount = ' '.join(amount.strip().split())
            
            amount_val = safe_float(amount)
            
            mapped = PAGE10_ROW_MAPPING.get(row_idx + 1, "")
            if mapped:
                print(f"Row {row_idx+1:2d}: [{mapped[:40]:40s}] = {amount_val}")
            elif account:
                print(f"Row {row_idx+1:2d}: [{account[:40]:40s}] = {amount_val}")
    
    doc.close()
    return best_offset

# ============================================
# PDF EXTRACTION FUNCTIONS
# ============================================

def find_pl_page(pdf_path):
    """Find P&L page by content markers"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        markers_found = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if markers_found >= 3:
            return doc, i
    doc.close()
    return None, None

def extract_table_from_page(page):
    """Extract 27 rows × 9 columns using hardcoded layout"""
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
            text = page.get_text("text", clip=fitz.Rect(x+1, y+1, x+cell_w-1, y+row_h-1))
            row_data.append(' '.join(text.strip().split()))
            x += cell_w
        
        data.append(row_data)
    
    return data

def extract_from_pdf(pdf_path, debug=False):
    """Extract financial data from a PDF monthly report"""
    if debug:
        # Run diagnostic first to find best offset
        best_offset = run_diagnostic(pdf_path)
        if best_offset:
            global TABLE_TOP
            TABLE_TOP = best_offset
    
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        print(f"❌ No P&L page found in {pdf_path}")
        return {}
    
    page = doc[page_num]
    table = extract_table_from_page(page)
    doc.close()
    
    financial_data = {}
    print(f"\nExtracting from {os.path.basename(pdf_path)}:")
    
    for row_num, account_name in PAGE10_ROW_MAPPING.items():
        if account_name in ['DÉPENSES', "DÉPENSES D'EXPLOITATION", 'AUTRES FRAIS']:
            continue
        
        if row_num <= len(table):
            raw_amount = table[row_num - 1][AMOUNT_COL]
            amount = safe_float(raw_amount, debug=debug)
            
            if amount is not None:
                financial_data[account_name] = amount
                if debug:
                    print(f"  ✅ Row {row_num:2d}: {account_name[:40]} = ${amount:,.2f}")
            elif debug:
                print(f"  ❌ Row {row_num:2d}: {account_name[:40]} = '{raw_amount}' (FAILED)")
    
    print(f"  Total: {len(financial_data)} accounts extracted")
    return financial_data

# ============================================
# TEMPLATE FILLING FUNCTIONS
# ============================================

def fill_template_sheet(wb, all_monthly_data):
    """Fill the DONNÉES HISTORIQUES sheet with extracted data."""
    # Find the correct sheet
    sheet_name = None
    for sn in wb.sheetnames:
        if 'DONNÉE' in sn.upper() or 'HISTORIQUE' in sn.upper():
            sheet_name = sn
            break
    
    if sheet_name is None:
        for sn in wb.sheetnames:
            ws = wb[sn]
            if ws.max_row > 50:
                sheet_name = sn
                break
    
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
    
    ws = wb[sheet_name]
    updates = []
    total_cells_updated = 0
    
    print(f"\n📝 Filling sheet: '{sheet_name}'")
    
    for month_name, account_data in all_monthly_data.items():
        col = MONTH_COLUMN.get(month_name)
        if col is None:
            updates.append(f"⚠️ Unknown month: {month_name}")
            continue
        
        month_updates = 0
        
        for french_account, amount in account_data.items():
            row_num = TEMPLATE_ROW_MAP.get(french_account)
            if row_num is None:
                continue
            
            cell = ws.cell(row=row_num, column=col)
            cell.value = amount
            cell.number_format = '#,##0.00'
            
            month_updates += 1
            total_cells_updated += 1
        
        if month_updates > 0:
            updates.append(f"📊 {month_name}: {month_updates} cells filled (column {get_column_letter(col)})")
    
    updates.append(f"📊 TOTAL: {total_cells_updated} cells filled")
    return updates

def validate_template(wb, all_monthly_data):
    """Validate BÉNÉFICE NET = REVENUS NETS"""
    sheet_name = None
    for sn in wb.sheetnames:
        if 'DONNÉE' in sn.upper() or 'HISTORIQUE' in sn.upper():
            sheet_name = sn
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
    
    ws = wb[sheet_name]
    validation_results = []
    
    for month_name, account_data in all_monthly_data.items():
        col = MONTH_COLUMN.get(month_name)
        if col is None:
            continue
        
        benefice_net = account_data.get('BÉNÉFICE NET')
        revenus_nets_cell = ws.cell(row=68, column=col).value
        
        if benefice_net is not None and revenus_nets_cell is not None:
            difference = abs(benefice_net - revenus_nets_cell)
            if difference > 0.01:
                validation_results.append(
                    f"⚠️ {month_name}: BÉNÉFICE NET (${benefice_net:,.2f}) ≠ "
                    f"REVENUS NETS (${revenus_nets_cell:,.2f}) - Diff: ${difference:,.2f}"
                )
            else:
                validation_results.append(
                    f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net:,.2f}"
                )
        elif benefice_net is not None and revenus_nets_cell is None:
            validation_results.append(
                f"⚠️ {month_name}: BÉNÉFICE NET extracted (${benefice_net:,.2f}) "
                f"but REVENUS NETS cell is empty"
            )
    
    return validation_results

# ============================================
# MAIN FUNCTIONS (called by app.py)
# ============================================

def get_parking_codes_from_pnl(file_obj):
    """Extract parking codes from uploaded file."""
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

def process_pdf_file(file_obj, debug=False):
    """Process a single PDF file object. Returns (month_name, financial_data_dict)"""
    month = extract_month_from_filename(file_obj)
    
    file_obj.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(file_obj.read())
        tmp_path = tmp.name
    
    try:
        if debug:
            financial_data = extract_from_pdf(tmp_path, debug=True)
        else:
            doc, page_num = find_pl_page(tmp_path)
            if doc is None:
                return month, {}
            
            page = doc[page_num]
            table = extract_table_from_page(page)
            doc.close()
            
            financial_data = {}
            for row_num, account_name in PAGE10_ROW_MAPPING.items():
                if account_name in ['DÉPENSES', "DÉPENSES D'EXPLOITATION", 'AUTRES FRAIS']:
                    continue
                if row_num <= len(table):
                    amount = safe_float(table[row_num - 1][AMOUNT_COL])
                    if amount is not None:
                        financial_data[account_name] = amount
        
        return month, financial_data
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """
    MAIN FUNCTION - Called by app.py when user clicks "Run Workflow".
    """
    updates = []
    all_monthly_data = {}
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW STARTED")
        updates.append("=" * 60)
        
        # Load template
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        updates.append(f"   Sheets: {wb.sheetnames}")
        
        # Process current year files
        if monthly_files_current:
            updates.append(f"\n📁 Processing {len(monthly_files_current)} current year files...")
            
            for i, file_obj in enumerate(monthly_files_current):
                # Use debug mode for first file to find correct offset
                use_debug = (i == 0)
                month, data = process_pdf_file(file_obj, debug=use_debug)
                
                if month and data:
                    all_monthly_data[month] = data
                    updates.append(f"   ✅ {month}: {len(data)} accounts extracted")
                    if 'TOTAL REVENUS' in data:
                        updates.append(f"      Revenue: ${data['TOTAL REVENUS']:,.2f}")
                    if 'BÉNÉFICE NET' in data:
                        updates.append(f"      Net Income: ${data['BÉNÉFICE NET']:,.2f}")
                else:
                    updates.append(f"   ❌ Failed: {file_obj.name if hasattr(file_obj, 'name') else 'Unknown file'}")
        
        # Process previous year files
        if monthly_files_previous:
            updates.append(f"\n📁 Processing {len(monthly_files_previous)} previous year files...")
            
            for file_obj in monthly_files_previous:
                month, data = process_pdf_file(file_obj)
                if month and data:
                    all_monthly_data[month] = data
                    updates.append(f"   ✅ {month}: {len(data)} accounts extracted")
                else:
                    updates.append(f"   ❌ Failed")
        
        # Fill template
        if all_monthly_data:
            updates.append(f"\n📝 Filling template with {len(all_monthly_data)} months...")
            fill_updates = fill_template_sheet(wb, all_monthly_data)
            updates.extend(fill_updates)
            
            # Validate
            updates.append(f"\n🔍 Validation:")
            validation_results = validate_template(wb, all_monthly_data)
            updates.extend(validation_results)
        else:
            updates.append("\n⚠️ No data extracted!")
        
        # Save
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("\n" + "=" * 60)
        updates.append("✅ WORKFLOW COMPLETED")
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
        if pdf_path == "--diagnostic" and len(sys.argv) > 2:
            # Run diagnostic on specific file
            run_diagnostic(sys.argv[2])
        else:
            # Run extraction with diagnostic
            data = extract_from_pdf(pdf_path, debug=True)
            print(f"\n{'='*60}")
            print("FINAL EXTRACTED DATA:")
            print(f"{'='*60}")
            for account, amount in data.items():
                print(f"  {account}: ${amount:,.2f}")
    else:
        print("Usage:")
        print("  python excel_fixer.py --diagnostic <pdf_file>")
        print("  python excel_fixer.py <pdf_file>")
