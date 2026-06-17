# excel_fixer.py - COMPLETE VERSION FOR PAGE 1
"""
Excel Fixer Module - PDF to Excel conversion + Template filling
Only Solutions Inc. - Budget System
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
    8: 'DÉPENSES',  # Section header - skip
    9: "DÉPENSES D'EXPLOITATION",  # Sub-header - skip
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
    24: 'AUTRES FRAIS',  # Section header - skip
    25: 'Honoraires de gestion',
    26: 'Total des autres frais',
    27: 'BÉNÉFICE NET'
}

# ============================================
# TEMPLATE MAPPING
# ============================================
# Month -> Column in template (1-based)
MONTH_COLUMN = {
    'Janvier': 2,    # B
    'Février': 3,    # C
    'Mars': 4,       # D
    'Avril': 5,      # E
    'Mai': 6,        # F
    'Juin': 7,       # G
    'Juillet': 8,    # H
    'Août': 9,       # I
    'Septembre': 10, # J
    'Octobre': 11,   # K
    'Novembre': 12,  # L
    'Décembre': 13   # M
}

# Maps P&L account (French) -> Template row number
TEMPLATE_ROW_MAP = {
    'Revenus mensuels': 11,           # Row 11: Revenus mensuels
    'Revenus Journaliers': 9,         # Row 9: Revenus horaires (closest match)
    'Revenus Lave-Auto': 12,          # Row 12: Revenus Lave-auto
    'Divers': 15,                     # Row 15: Autres revenus
    'Gratuités - mensuels': 18,       # Row 18: (Gratuités)
    'TOTAL REVENUS': 24,              # Row 24: TOTAL REVENUS
    'Salaires Stationnement': 27,     # Row 27: Salaire Stationnement
    'Uniformes': 30,                  # Row 30: Uniformes
    'Entretien réparation - Nettoyage': 34,    # Row 34: Nettoyage stationnement
    'Entretien réparation - Equipement': 36,   # Row 36: Entretien équipement
    'Entretien réparation - Général': 35,      # Row 35: Entretien stationnement
    'Fourn. de stationnement': 40,    # Row 40: Fournitures stationnement
    'Taxes et permis': 54,            # Row 54: Taxes et permis
    'Assurances Cautionnement': 53,   # Row 53: Assurances et cautionnement
    'Réclamations': 52,               # Row 52: Réclamations
    'Télécommunication': 47,          # Row 47: Telecommunications
    'Frais de cartes de crédit': 49,  # Row 49: Frais de cartes de crédit
    'Frais de bureau': 46,            # Row 46: Fournitures de bureau
    "Total des frais d'exploitation": 66,  # Row 66: TOTAL DÉPENSES
    'Honoraires de gestion': 56,      # Row 56: Honoraires de gestion (combined)
    'BÉNÉFICE NET': 68,              # Row 68: REVENUS NETS
}

# Accounts that MUST match for validation
VALIDATION_PAIRS = [
    ('BÉNÉFICE NET', 'REVENUS NETS', 68),
    ('TOTAL REVENUS', 'TOTAL REVENUS', 24),
    ("Total des frais d'exploitation", 'TOTAL DÉPENSES', 66),
]

# ============================================
# UTILITY FUNCTIONS
# ============================================

def safe_float(value):
    """
    Convert European number format to float.
    Handles: "7 106 417,00", "(42 000,00)", "1 234.56"
    """
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    
    value = str(value).strip()
    
    # Handle parentheses (negative)
    is_negative = value.startswith('(') and value.endswith(')')
    if is_negative:
        value = value[1:-1]
    
    # Remove spaces
    value = value.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    
    # Handle comma as decimal
    if ',' in value:
        value = value.replace(',', '.')
    
    # Remove any remaining non-numeric (except minus and dot)
    value = re.sub(r'[^\d\.\-]', '', value)
    
    try:
        result = float(value)
        return -result if is_negative else result
    except:
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
    
    # Try regex
    match = re.search(
        r'(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)',
        name, re.IGNORECASE
    )
    if match:
        return match.group(1).capitalize()
    
    return None

# ============================================
# PDF EXTRACTION FUNCTIONS
# ============================================

def find_pl_page(pdf_path):
    """Find P&L page by content markers (NOT page number)"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        matches = sum(1 for m in PAGE_MARKERS if m.lower() in text.lower())
        if matches >= 3:  # At least 3 of 4 markers
            print(f"   📄 P&L found on page {i+1} ({matches}/4 markers)")
            return doc, i
    doc.close()
    return None, None

def extract_table_from_page(page):
    """
    Extract 27 rows × 9 columns using hardcoded spatial layout.
    This is the SAME layout regardless of which page number.
    """
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
            # Extract text with 1pt margin to avoid borders
            rect = fitz.Rect(x + 1, y + 1, x + cell_w - 1, y + row_h - 1)
            text = page.get_text("text", clip=rect)
            row_data.append(' '.join(text.strip().split()))
            x += cell_w
        
        data.append(row_data)
    
    return data

def pdf_to_excel(pdf_path, output_dir=None):
    """
    Convert PDF P&L page to Excel for verification.
    Returns both the DataFrame and extracted financial data.
    """
    print(f"\n📄 Processing: {os.path.basename(pdf_path)}")
    
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        print(f"   ❌ P&L page not found")
        return None, {}
    
    page = doc[page_num]
    table = extract_table_from_page(page)
    doc.close()
    
    # Create DataFrame
    columns = [
        'Account', 'Mois Courant', 'Budget', 'Écart', 'An Préc',
        'Cumulatif', 'Cumul budget', 'Écart cumul', 'An Préc cumul'
    ]
    df = pd.DataFrame(table, columns=columns)
    
    # Save to Excel for verification
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        excel_path = os.path.join(output_dir, f"{Path(pdf_path).stem}_extracted.xlsx")
        df.to_excel(excel_path, index=False)
        print(f"   💾 Saved: {excel_path}")
    
    # Extract financial data
    financial_data = {}
    for row_num, account_name in PAGE10_ROW_MAPPING.items():
        # Skip section headers
        if account_name in ['DÉPENSES', "DÉPENSES D'EXPLOITATION", 'AUTRES FRAIS']:
            continue
        
        if row_num <= len(table):
            raw_amount = table[row_num - 1][AMOUNT_COL]
            amount = safe_float(raw_amount)
            
            if amount is not None:
                financial_data[account_name] = amount
                print(f"   ✅ Row {row_num:2d}: {account_name[:40]:40s} = ${amount:>12,.2f}")
            else:
                print(f"   ⚠️ Row {row_num:2d}: {account_name[:40]:40s} = '{raw_amount}' (not converted)")
    
    return df, financial_data

# ============================================
# TEMPLATE FILLING FUNCTIONS
# ============================================

def fill_template_sheet(wb, all_monthly_data):
    """
    Fill the DONNÉES HISTORIQUES sheet with extracted data.
    
    Args:
        wb: openpyxl Workbook
        all_monthly_data: dict {month_name: {french_account: amount}}
    
    Returns:
        updates: list of update messages
    """
    # Find the correct sheet
    sheet_name = None
    for sn in wb.sheetnames:
        if 'DONNÉE' in sn.upper() or 'HISTORIQUE' in sn.upper():
            sheet_name = sn
            break
    
    if sheet_name is None:
        # Try to find any sheet with data
        for sn in wb.sheetnames:
            ws = wb[sn]
            if ws.max_row > 50:  # Assume sheet with lots of rows is the data sheet
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
            
            # Write to cell
            cell = ws.cell(row=row_num, column=col)
            cell.value = amount
            
            # Format as currency
            cell.number_format = '#,##0.00'
            
            month_updates += 1
            total_cells_updated += 1
            
            # Get template row label for logging
            updates.append(f"   ✅ {month_name} (col {get_column_letter(col)}): "
                          f"Row {row_num} = ${amount:,.2f}")
        
        if month_updates > 0:
            updates.append(f"📊 {month_name}: {month_updates} cells filled")
    
    updates.append(f"\n📊 TOTAL: {total_cells_updated} cells filled in template")
    return updates

def validate_template(wb, all_monthly_data):
    """
    Validate that BÉNÉFICE NET (PDF) = REVENUS NETS (template).
    Also check TOTAL REVENUS and TOTAL DÉPENSES.
    """
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
        
        # Check BÉNÉFICE NET vs REVENUS NETS
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
        
        # Check TOTAL REVENUS
        total_revenus = account_data.get('TOTAL REVENUS')
        template_total = ws.cell(row=24, column=col).value
        if total_revenus is not None and template_total is not None:
            if abs(total_revenus - template_total) > 0.01:
                validation_results.append(
                    f"⚠️ {month_name}: TOTAL REVENUS mismatch: "
                    f"PDF=${total_revenus:,.2f} vs Template=${template_total:,.2f}"
                )
    
    return validation_results

# ============================================
# MAIN FUNCTIONS (called by app.py)
# ============================================

def get_parking_codes_from_pnl(file_obj):
    """
    Extract parking codes from uploaded file.
    Called by app.py to populate the parking code selector.
    """
    codes = []
    
    if hasattr(file_obj, 'name'):
        # From filename
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
        
        # From content
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

def process_pdf_file(file_obj):
    """
    Process a single PDF file object.
    Returns (month_name, financial_data_dict)
    """
    month = extract_month_from_filename(file_obj)
    
    # Save to temp file for PyMuPDF
    file_obj.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(file_obj.read())
        tmp_path = tmp.name
    
    try:
        _, financial_data = pdf_to_excel(tmp_path)
        return month, financial_data
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """
    MAIN FUNCTION - Called by app.py when user clicks "Run Workflow".
    
    Process:
    1. Convert all PDFs to data
    2. Fill template DONNÉES HISTORIQUES sheet
    3. Validate BÉNÉFICE NET = REVENUS NETS
    4. Return updated Excel file
    
    Returns:
        (excel_bytes, updates_list)
    """
    updates = []
    all_monthly_data = {}
    
    try:
        # ============================================
        # STEP 1: Load template
        # ============================================
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW STARTED")
        updates.append("=" * 60)
        
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        updates.append(f"   Sheets available: {wb.sheetnames}")
        
        # ============================================
        # STEP 2: Process current year PDFs (Jan-Apr 2026)
        # ============================================
        if monthly_files_current:
            updates.append(f"\n📁 Processing {len(monthly_files_current)} current year files...")
            
            for file_obj in monthly_files_current:
                month, data = process_pdf_file(file_obj)
                if month and data:
                    all_monthly_data[month] = data
                    updates.append(f"   ✅ {month}: {len(data)} accounts extracted")
                    
                    # Show key figures
                    if 'TOTAL REVENUS' in data:
                        updates.append(f"      Revenue: ${data['TOTAL REVENUS']:,.2f}")
                    if 'BÉNÉFICE NET' in data:
                        updates.append(f"      Net Income: ${data['BÉNÉFICE NET']:,.2f}")
                else:
                    updates.append(f"   ❌ Failed to extract data")
        
        # ============================================
        # STEP 3: Process previous year PDFs (May-Dec 2025)
        # ============================================
        if monthly_files_previous:
            updates.append(f"\n📁 Processing {len(monthly_files_previous)} previous year files...")
            
            for file_obj in monthly_files_previous:
                month, data = process_pdf_file(file_obj)
                if month and data:
                    all_monthly_data[month] = data
                    updates.append(f"   ✅ {month}: {len(data)} accounts extracted")
                else:
                    updates.append(f"   ❌ Failed to extract data")
        
        # ============================================
        # STEP 4: Fill template
        # ============================================
        if all_monthly_data:
            updates.append(f"\n📝 Filling template with {len(all_monthly_data)} months of data...")
            fill_updates = fill_template_sheet(wb, all_monthly_data)
            updates.extend(fill_updates)
            
            # ============================================
            # STEP 5: Validate
            # ============================================
            updates.append(f"\n🔍 Validating data integrity...")
            validation_results = validate_template(wb, all_monthly_data)
            updates.extend(validation_results)
        else:
            updates.append("\n⚠️ No data extracted from any files!")
        
        # ============================================
        # STEP 6: Save and return
        # ============================================
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("\n" + "=" * 60)
        updates.append("✅ WORKFLOW COMPLETED SUCCESSFULLY")
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
    # Test with a single PDF
    import sys
    
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        df, data = pdf_to_excel(pdf_path, "extracted_data")
        
        if data:
            print(f"\n{'='*60}")
            print("EXTRACTION SUMMARY")
            print(f"{'='*60}")
            for account, amount in data.items():
                print(f"  {account}: ${amount:,.2f}")
