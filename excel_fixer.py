# excel_fixer.py - WITH DIAGNOSTIC MODE
"""
Extracts P&L data from PDFs and fills CMO111.xlsx template
Set DIAGNOSTIC_MODE = True to debug extraction
"""
import fitz
import pandas as pd
import os
import re
from io import BytesIO
import tempfile
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ============================================
# CONFIGURATION
# ============================================
DIAGNOSTIC_MODE = True  # Set to True to see detailed extraction debug

# PDF P&L PAGE - SPATIAL MAPPING
COLUMN_WIDTHS = [0.22, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10]
AMOUNT_COL = 1
NUM_ROWS = 27
TABLE_TOP = 0.08
TABLE_BOTTOM = 0.95
TABLE_LEFT = 0.03
TABLE_RIGHT = 0.97

PAGE_MARKERS = [
    "1981 McGill College",
    "Revenus mensuels",
    "BÉNÉFICE NET",
    "Mois Courant"
]

# PDF Row -> French Account Name
PDF_ROW_MAP = {
    1: 'Revenus mensuels',
    2: 'Revenus Journaliers',
    3: 'Revenus Lave-Auto',
    4: 'Divers',
    5: 'Revenus de stationnement',
    6: 'Gratuités - mensuels',
    7: 'TOTAL REVENUS',
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
    25: 'Honoraires de gestion',
    26: 'Total des autres frais',
    27: 'BÉNÉFICE NET'
}

# TEMPLATE MAPPING
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

def find_pl_page(pdf_path):
    """Find P&L page by content markers"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        matches = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if matches >= 3:
            if DIAGNOSTIC_MODE:
                print(f"   📄 P&L found on page {i+1} ({matches}/4 markers)")
            return doc, i
    
    if DIAGNOSTIC_MODE:
        print(f"   ❌ P&L NOT FOUND - checking all pages...")
        for i in range(len(doc)):
            text = doc[i].get_text()
            if "Revenus" in text or "DÉPENSES" in text:
                print(f"   Page {i+1} has financial content but missing markers")
                print(f"   First 300 chars: {text[:300]}")
    
    doc.close()
    return None, None

def extract_table_from_page(page, pdf_name=""):
    """Extract 27 rows x 9 columns from P&L page"""
    w, h = page.rect.width, page.rect.height
    
    if DIAGNOSTIC_MODE:
        print(f"\n   {'='*50}")
        print(f"   DIAGNOSTIC: {pdf_name}")
        print(f"   Page size: {w:.0f} x {h:.0f}")
        print(f"   Table Y: {h*TABLE_TOP:.0f} to {h*TABLE_BOTTOM:.0f}")
        print(f"   Row height: {(h*TABLE_BOTTOM - h*TABLE_TOP)/27:.1f}")
        print(f"   {'='*50}")
        
        # Show full page text
        full_text = page.get_text()
        print(f"\n   FULL PAGE TEXT (first 2000 chars):")
        print(f"   {full_text[:2000]}")
        print(f"\n   {'='*50}")
        
        # Test multiple Y offsets
        for y_off in [0.05, 0.06, 0.07, 0.08, 0.09, 0.10, 0.12]:
            top = h * y_off
            row_h = (h * TABLE_BOTTOM - top) / NUM_ROWS
            left = w * TABLE_LEFT
            right = w * TABLE_RIGHT
            tw = right - left
            cw = [0.22, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10]
            
            print(f"\n   --- Testing Y offset: {y_off} (top={top:.0f}, row_h={row_h:.1f}) ---")
            
            for row_idx in range(NUM_ROWS):
                y = top + row_idx * row_h
                x = left
                
                w0 = cw[0] * tw
                acc = page.get_text("text", clip=fitz.Rect(x, y, x+w0, y+row_h)).strip()
                acc = ' '.join(acc.split())
                
                x += w0
                w1 = cw[1] * tw
                amt = page.get_text("text", clip=fitz.Rect(x, y, x+w1, y+row_h)).strip()
                amt = ' '.join(amt.split())
                
                if acc:  # Only show rows with content
                    marker = ""
                    if "TOTAL REVENUS" in acc.upper():
                        marker = " ⭐ TOTAL REVENUS"
                    elif "BÉNÉFICE NET" in acc.upper():
                        marker = " ⭐ BÉNÉFICE NET"
                    elif "Total des frais" in acc:
                        marker = " ⭐ TOTAL EXPENSES"
                    elif "RÉSULTAT" in acc.upper():
                        marker = " ⭐ OPERATING SURPLUS"
                    
                    print(f"   Row {row_idx+1:2d}: [{acc[:45]:45s}] [${amt[:20]:20s}]{marker}")
            
            # Check if we found key accounts at this offset
            print()
    
    # Use the default coordinates
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
    
    if DIAGNOSTIC_MODE:
        print(f"\n   {'='*50}")
        print(f"   EXTRACTED DATA (using Y offset {TABLE_TOP}):")
        print(f"   {'='*50}")
        for row_idx, row in enumerate(data):
            if row[0] or row[1]:
                print(f"   Row {row_idx+1:2d}: [{row[0][:45]:45s}] [${row[1][:20]:20s}]")
    
    return data

def extract_from_pdf(pdf_path):
    """Extract financial data from PDF"""
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        return {}
    
    page = doc[page_num]
    pdf_name = os.path.basename(pdf_path)
    table = extract_table_from_page(page, pdf_name)
    doc.close()
    
    data = {}
    for row_num, account_name in PDF_ROW_MAP.items():
        if row_num <= len(table):
            raw = table[row_num - 1][AMOUNT_COL]
            amount = safe_float(raw)
            if amount is not None:
                data[account_name] = amount
                if DIAGNOSTIC_MODE:
                    print(f"   ✅ Extracted: {account_name} = ${amount:,.2f}")
            else:
                if DIAGNOSTIC_MODE:
                    print(f"   ❌ Failed: {account_name} - raw='{raw}'")
    
    return data

def fill_template(wb, all_data):
    """Fill '2-Données historiques' sheet"""
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
    """Validate BÉNÉFICE NET = REVENUS NETS"""
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
                results.append(
                    f"⚠️ {month_name}: BÉNÉFICE NET (${benefice_net_pdf:,.2f}) ≠ "
                    f"REVENUS NETS (${revenus_nets_template:,.2f}) - Diff: ${diff:,.2f}"
                )
            else:
                results.append(f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net_pdf:,.2f}")
        elif benefice_net_pdf is None:
            results.append(f"⚠️ {month_name}: BÉNÉFICE NET not extracted from PDF")
        elif revenus_nets_template is None:
            results.append(f"⚠️ {month_name}: REVENUS NETS not found in template")
    
    return results

def get_parking_codes_from_pnl(file_obj):
    """Extract parking codes"""
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
    """MAIN FUNCTION"""
    updates = []
    all_data = {}
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append(f"   Diagnostic Mode: {'ON' if DIAGNOSTIC_MODE else 'OFF'}")
        updates.append("=" * 60)
        
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
        if monthly_files_current:
            updates.append(f"\n📁 Current year: {len(monthly_files_current)} files")
            for file_obj in monthly_files_current:
                month = extract_month_from_filename(file_obj)
                file_obj.seek(0)
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(file_obj.read())
                    tmp_path = tmp.name
                
                try:
                    if DIAGNOSTIC_MODE:
                        print(f"\n{'='*80}")
                        print(f"PROCESSING: {month} - {file_obj.name}")
                        print(f"{'='*80}")
                    
                    data = extract_from_pdf(tmp_path)
                    if data:
                        all_data[month] = data
                        updates.append(f"   ✅ {month}: {len(data)} accounts")
                        if 'TOTAL REVENUS' in data:
                            updates.append(f"      Revenue: ${data['TOTAL REVENUS']:,.2f}")
                        if 'BÉNÉFICE NET' in data:
                            updates.append(f"      Net Income: ${data['BÉNÉFICE NET']:,.2f}")
                    else:
                        updates.append(f"   ❌ {month}: No data extracted")
                finally:
                    os.unlink(tmp_path)
        
        if monthly_files_previous:
            updates.append(f"\n📁 Previous year: {len(monthly_files_previous)} files")
            for file_obj in monthly_files_previous:
                month = extract_month_from_filename(file_obj)
                file_obj.seek(0)
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(file_obj.read())
                    tmp_path = tmp.name
                
                try:
                    data = extract_from_pdf(tmp_path)
                    if data:
                        all_data[month] = data
                        updates.append(f"   ✅ {month}: {len(data)} accounts")
                finally:
                    os.unlink(tmp_path)
        
        if all_data:
            updates.append(f"\n📝 Filling template with {len(all_data)} months...")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            
            updates.append(f"\n🔍 Validation:")
            validations = validate_template(wb, all_data)
            updates.extend(validations)
        else:
            updates.append("\n⚠️ No data extracted from any file!")
        
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
