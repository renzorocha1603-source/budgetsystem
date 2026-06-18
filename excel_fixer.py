# excel_fixer.py - COMPLETE CORRECTED VERSION
"""
Extracts P&L data from PDFs and fills CMO111.xlsx template
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
# PDF P&L PAGE - SPATIAL MAPPING
# ============================================
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

# ============================================
# TEMPLATE "2-Données historiques" - SPATIAL MAPPING
# ============================================
MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

# EXACT mapping: PDF account -> Template Row
# Based on the spatial mappings provided
PDF_TO_TEMPLATE = {
    # REVENUES
    'Revenus mensuels': 13,           # Template Row 13: Revenus mensuels
    'Revenus Journaliers': 12,        # Template Row 12: Revenus horaires
    'Revenus Lave-Auto': 14,          # Template Row 14: Revenus Lave-auto
    'Divers': 17,                     # Template Row 17: Autres revenus
    'Gratuités - mensuels': 20,       # Template Row 20: (Gratuités)
    'TOTAL REVENUS': 26,              # Template Row 26: TOTAL REVENUS
    
    # LABOUR
    'Salaires Stationnement': 29,     # Template Row 29: Salaire Stationnement
    'Uniformes': 32,                  # Template Row 32: Uniformes
    
    # MAINTENANCE
    'Fourn. de stationnement': 41,    # Template Row 41: Fournitures stationnement
    'Entretien réparation - Nettoyage': 35,  # Template Row 35: Nettoyage stationnement
    'Entretien réparation - Equipement': 37, # Template Row 37: Entretien équipement
    'Entretien réparation - Général': 36,    # Template Row 36: Entretien stationnement
    
    # OVERHEAD
    'Frais de carte de crédit': 53,   # Template Row 53: Frais de cartes de crédit
    'Frais de bureau': 49,            # Template Row 49: Fournitures de bureau
    'Télécommunication': 50,          # Template Row 50: Télécommunications
    'Taxes et permis': 58,            # Template Row 58: Taxes et permis
    'Assurances Cautionnement': 57,   # Template Row 57: Assurances et cautionnement
    'Réclamations': 56,               # Template Row 56: Réclamations
    'Honoraires de gestion': 63,      # Template Row 63: Honoraires de gestion de base
    
    # TOTALS
    "Total des frais d'exploitation": 84,  # Template Row 84: TOTAL DÉPENSES
    'BÉNÉFICE NET': 86,                   # Template Row 86: REVENUS NETS
}

# ============================================
# UTILITY FUNCTIONS
# ============================================

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
# PDF EXTRACTION
# ============================================

def find_pl_page(pdf_path):
    """Find P&L page by content markers"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        matches = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if matches >= 3:
            return doc, i
    doc.close()
    return None, None

def extract_table_from_page(page):
    """Extract 27 rows x 9 columns from P&L page"""
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

def extract_from_pdf(pdf_path):
    """Extract financial data from PDF, return {french_account: amount}"""
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        return {}
    
    page = doc[page_num]
    table = extract_table_from_page(page)
    doc.close()
    
    data = {}
    for row_num, account_name in PDF_ROW_MAP.items():
        if row_num <= len(table):
            raw = table[row_num - 1][AMOUNT_COL]
            amount = safe_float(raw)
            if amount is not None:
                data[account_name] = amount
    return data

# ============================================
# TEMPLATE FILLING
# ============================================

def fill_template(wb, all_data):
    """
    Fill '2-Données historiques' sheet with extracted data.
    all_data: {month_name: {pdf_account: amount}}
    """
    # Find the sheet
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        for sn in wb.sheetnames:
            ws = wb[sn]
            if ws.max_row >= 86:  # This sheet has 86+ rows
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
    MAIN FUNCTION - Extract P&L data from PDFs and fill template.
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
        
        # Process current year files (Jan-Apr 2026)
        if monthly_files_current:
            updates.append(f"\n📁 Current year: {len(monthly_files_current)} files")
            for file_obj in monthly_files_current:
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
                        if 'BÉNÉFICE NET' in data:
                            updates.append(f"      Net Income: ${data['BÉNÉFICE NET']:,.2f}")
                    else:
                        updates.append(f"   ❌ {month}: No data extracted")
                finally:
                    os.unlink(tmp_path)
        
        # Process previous year files (May-Dec 2025)
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
        
        # Fill template
        if all_data:
            updates.append(f"\n📝 Filling template...")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            
            # Validate
            updates.append(f"\n🔍 Validation:")
            validations = validate_template(wb, all_data)
            updates.extend(validations)
        else:
            updates.append("\n⚠️ No data extracted!")
        
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
