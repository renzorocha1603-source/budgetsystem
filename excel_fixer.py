import fitz  # PyMuPDF
import pandas as pd
import os
from pathlib import Path
import re
from datetime import datetime

# ============================================
# HARDCODED SPATIAL LAYOUT
# ============================================
COLUMN_WIDTHS = [0.22, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10]
AMOUNT_COL = 1  # Mois Courant - ALWAYS column 1
NUM_ROWS = 27
TABLE_TOP = 0.08
TABLE_BOTTOM = 0.95
TABLE_LEFT = 0.03
TABLE_RIGHT = 0.97

# Content markers to identify the P&L page
PAGE_MARKERS = [
    "1981 McGill College",
    "Revenus mensuels", 
    "BÉNÉFICE NET",
    "Mois Courant"
]

# Complete row-to-account mapping
ROW_MAPPING = {
    1: 'Revenus mensuels',
    2: 'Revenus Journaliers',
    3: 'Revenus Lave-Auto',
    4: 'Divers',
    5: 'Revenus de stationnement',
    6: 'Gratuités - mensuels',
    7: 'TOTAL REVENUS',
    8: 'DÉPENSES',  # Section header
    9: "DÉPENSES D'EXPLOITATION",  # Sub-header
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
    24: 'AUTRES FRAIS',  # Section header
    25: 'Honoraires de gestion',
    26: 'Total des autres frais',
    27: 'BÉNÉFICE NET'
}

# Mapping to your DH_ROW_MAPPING English labels
LABEL_MAPPING = {
    'Revenus mensuels': 'Monthly Revenues',
    'Revenus Journaliers': 'Daily Revenues',
    'Revenus Lave-Auto': 'Car Wash Revenues',
    'Divers': 'Miscellaneous Income',
    'Revenus de stationnement': 'Total Parking Revenue',
    'Gratuités - mensuels': 'Monthly Gratuities',
    'TOTAL REVENUS': 'Total Revenue',
    'Salaires Stationnement': 'Parking Salaries',
    'Uniformes': 'Uniforms',
    'Fourn. de stationnement': 'Parking Supplies',
    'Entretien réparation - Nettoyage': 'Maintenance - Cleaning',
    'Entretien réparation - Equipement': 'Maintenance - Equipment',
    'Entretien réparation - Général': 'Maintenance - General',
    'Taxes et permis': 'Taxes & Permits',
    'Assurances Cautionnement': 'Insurance & Bonding',
    'Réclamations': 'Claims',
    'Télécommunication': 'Telecommunication',
    'Frais de cartes de crédit': 'Credit Card Fees',
    'Frais de bureau': 'Office Expenses',
    "Total des frais d'exploitation": 'Total Operating Expenses',
    "RÉSULTAT D'EXPLOITATION": 'Operating Surplus',
    'Honoraires de gestion': 'Management Fees',
    'Total des autres frais': 'Total Other Expenses',
    'BÉNÉFICE NET': 'Net Income'
}

# ============================================
# CORE FUNCTIONS
# ============================================

def safe_float(value, debug=False):
    """
    Handle European number formats with debugging.
    Examples: "7 106 417,00" -> 7106417.00
              "1 234,56" -> 1234.56
              "(42 000,00)" -> -42000.00
    """
    if value is None or value == "":
        if debug: print(f"      safe_float: None/empty input")
        return None
    
    if isinstance(value, (int, float)):
        if debug: print(f"      safe_float: Already numeric -> {float(value)}")
        return float(value)
    
    original = str(value).strip()
    cleaned = original
    
    if debug: print(f"      safe_float input: '{original}'")
    
    # Handle parentheses (negative numbers)
    is_negative = False
    if cleaned.startswith('(') and cleaned.endswith(')'):
        is_negative = True
        cleaned = cleaned[1:-1]
    
    # Remove spaces (thousand separators)
    cleaned = cleaned.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    
    # Handle European decimal comma
    if "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    
    # Remove any remaining non-numeric chars except minus and decimal
    cleaned = re.sub(r'[^\d\.\-]', '', cleaned)
    
    try:
        result = float(cleaned)
        if is_negative:
            result = -result
        if debug: print(f"      safe_float output: {result}")
        return result
    except (ValueError, TypeError) as e:
        if debug: print(f"      safe_float FAILED: {e}")
        return None

def find_pl_page(pdf_path):
    """
    Find the P&L page by content markers.
    Works regardless of page number (5, 8, 10, etc.)
    """
    print(f"\n{'='*60}")
    print(f"FINDING P&L PAGE IN: {os.path.basename(pdf_path)}")
    print(f"{'='*60}")
    
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    print(f"Total pages in PDF: {total_pages}")
    
    for page_num in range(total_pages):
        page = doc[page_num]
        text = page.get_text()
        
        # Check all markers
        found = []
        missing = []
        for marker in PAGE_MARKERS:
            if marker.lower() in text.lower():
                found.append(marker)
            else:
                missing.append(marker)
        
        match_pct = len(found) / len(PAGE_MARKERS) * 100
        print(f"\nPage {page_num + 1}: {match_pct:.0f}% match")
        print(f"  Found: {found}")
        if missing:
            print(f"  Missing: {missing}")
        
        # All markers found = definitive match
        if len(found) == len(PAGE_MARKERS):
            print(f"\n✅ FOUND P&L on page {page_num + 1}")
            print(f"   Page dimensions: {page.rect.width:.0f} x {page.rect.height:.0f}")
            return doc, page_num, page
    
    print(f"\n❌ P&L page NOT FOUND in {pdf_path}")
    doc.close()
    return None, None, None

def extract_table_from_page(page, page_num):
    """
    Extract 27 rows × 9 columns using hardcoded spatial layout.
    """
    print(f"\n{'='*60}")
    print(f"EXTRACTING TABLE FROM PAGE {page_num + 1}")
    print(f"{'='*60}")
    
    w = page.rect.width
    h = page.rect.height
    
    # Calculate table boundaries
    top = h * TABLE_TOP
    bottom = h * TABLE_BOTTOM
    left = w * TABLE_LEFT
    right = w * TABLE_RIGHT
    table_w = right - left
    row_h = (bottom - top) / NUM_ROWS
    
    print(f"Table area: {left:.0f},{top:.0f} to {right:.0f},{bottom:.0f}")
    print(f"Row height: {row_h:.1f} points")
    print(f"Extracting {NUM_ROWS} rows × {len(COLUMN_WIDTHS)} columns...")
    
    table_data = []
    
    for row_idx in range(NUM_ROWS):
        row_data = []
        y_top = top + (row_idx * row_h)
        y_bottom = y_top + row_h
        
        x_pos = left
        for col_idx, col_width in enumerate(COLUMN_WIDTHS):
            x_right = x_pos + (col_width * table_w)
            
            # Extract text with 1pt margin to avoid borders
            cell_rect = fitz.Rect(
                x_pos + 1, y_top + 1,
                x_right - 1, y_bottom - 1
            )
            
            cell_text = page.get_text("text", clip=cell_rect)
            cell_text = ' '.join(cell_text.strip().split())  # Clean whitespace
            row_data.append(cell_text)
            
            x_pos = x_right
        
        table_data.append(row_data)
        
        # Debug: Show key rows
        if row_idx + 1 in [1, 7, 22, 23, 27]:
            print(f"\nRow {row_idx + 1}:")
            print(f"  Account: '{row_data[0]}'")
            print(f"  Mois Courant (col 1): '{row_data[1]}'")
            print(f"  Budget (col 2): '{row_data[2]}'")
    
    return table_data

def table_to_dataframe(table_data, pdf_path):
    """
    Convert extracted table to DataFrame and save to Excel.
    """
    columns = [
        'Account', 'Mois Courant', 'Budget période', 'Écart Budget',
        'An. Préc.', 'Cumulatif courant', 'Cumulatif budget',
        'Écart Budget Cumul.', 'An. Préc. Cumul.'
    ]
    
    df = pd.DataFrame(table_data, columns=columns)
    
    # Save to Excel
    pdf_name = Path(pdf_path).stem
    excel_path = pdf_path.replace('.pdf', '_extracted.xlsx')
    df.to_excel(excel_path, index=False, sheet_name='P&L Data')
    print(f"\n✅ Saved extracted table to: {excel_path}")
    
    return df, excel_path

def extract_financial_data(df, debug=True):
    """
    Extract all financial data using hardcoded row mappings.
    Returns dict with English labels.
    """
    print(f"\n{'='*60}")
    print(f"EXTRACTING FINANCIAL DATA (Column {AMOUNT_COL}: Mois Courant)")
    print(f"{'='*60}")
    
    financial_data = {}
    skipped = []
    errors = []
    
    for row_num, french_label in ROW_MAPPING.items():
        # Skip section headers
        if french_label in ['DÉPENSES', "DÉPENSES D'EXPLOITATION", 'AUTRES FRAIS']:
            continue
        
        if row_num > len(df):
            print(f"Row {row_num}: OUT OF RANGE (max {len(df)})")
            continue
        
        row_idx = row_num - 1  # 0-based
        raw_account = str(df.iloc[row_idx, 0])
        raw_amount = str(df.iloc[row_idx, AMOUNT_COL])
        
        # Convert amount
        amount = safe_float(raw_amount, debug=debug)
        
        # Map to English
        english_label = LABEL_MAPPING.get(french_label, french_label)
        
        if debug:
            print(f"\nRow {row_num}: {french_label} -> {english_label}")
            print(f"  Raw account: '{raw_account}'")
            print(f"  Raw amount: '{raw_amount}'")
            print(f"  Converted: {amount}")
        
        if amount is not None:
            financial_data[english_label] = amount
        else:
            errors.append((row_num, french_label, raw_amount))
    
    # Summary
    print(f"\n{'='*60}")
    print(f"EXTRACTION COMPLETE")
    print(f"  Successfully extracted: {len(financial_data)} accounts")
    print(f"  Errors: {len(errors)}")
    
    if errors:
        print(f"\n  Failed extractions:")
        for row, label, raw in errors:
            print(f"    Row {row} ({label}): '{raw}'")
    
    return financial_data

def process_monthly_pdf(pdf_path):
    """
    Main function: Find P&L page, extract data, return results.
    """
    print(f"\n{'='*80}")
    print(f"PROCESSING: {os.path.basename(pdf_path)}")
    print(f"{'='*80}")
    
    # Step 1: Find the P&L page
    doc, page_num, page = find_pl_page(pdf_path)
    
    if doc is None:
        raise ValueError(f"No P&L page found in {pdf_path}")
    
    # Step 2: Extract table from the page
    table_data = extract_table_from_page(page, page_num)
    
    # Step 3: Convert to DataFrame and save Excel
    df, excel_path = table_to_dataframe(table_data, pdf_path)
    
    # Step 4: Extract financial data
    financial_data = extract_financial_data(df, debug=True)
    
    # Step 5: Cleanup
    doc.close()
    
    # Step 6: Print key results
    print(f"\n{'='*60}")
    print(f"KEY FINANCIAL RESULTS")
    print(f"{'='*60}")
    for label in ['Total Revenue', 'Total Operating Expenses', 'Operating Surplus', 'Net Income']:
        amount = financial_data.get(label)
        if amount is not None:
            print(f"  {label}: ${amount:,.2f}")
        else:
            print(f"  {label}: MISSING")
    
    return financial_data, excel_path

def update_template(template_path, financial_data, month_year):
    """
    Update CMO111.xlsx template with extracted data.
    """
    print(f"\n{'='*60}")
    print(f"UPDATING TEMPLATE: {os.path.basename(template_path)}")
    print(f"  Month: {month_year}")
    print(f"  Accounts to update: {len(financial_data)}")
    print(f"{'='*60}")
    
    # Load template
    # This is where you'd add your template updating logic
    # For now, just verify the data is valid
    required_accounts = [
        'Total Revenue',
        'Total Operating Expenses', 
        'Operating Surplus',
        'Management Fees',
        'Net Income'
    ]
    
    missing = [acc for acc in required_accounts if acc not in financial_data]
    
    if missing:
        print(f"⚠️  Missing required accounts: {missing}")
        return False
    
    # Verify logical consistency
    total_rev = financial_data.get('Total Revenue', 0)
    total_exp = financial_data.get('Total Operating Expenses', 0)
    operating_surplus = financial_data.get('Operating Surplus', 0)
    net_income = financial_data.get('Net Income', 0)
    
    expected_surplus = total_rev - total_exp
    if abs(operating_surplus - expected_surplus) > 0.01:
        print(f"⚠️  Operating Surplus mismatch: {operating_surplus} vs expected {expected_surplus}")
    
    print(f"✅ Data validation complete")
    print(f"   Total Revenue: ${total_rev:,.2f}")
    print(f"   Total Expenses: ${total_exp:,.2f}")
    print(f"   Operating Surplus: ${operating_surplus:,.2f}")
    print(f"   Net Income: ${net_income:,.2f}")
    
    return True

# ============================================
# MAIN EXECUTION
# ============================================

if __name__ == "__main__":
    import sys
    
    # Process PDFs
    pdf_files = [
        "01-CMO111 Rapport mensuel de gestion janvier 2026.pdf",
        "02-CMO111 Rapport mensuel de gestion février 2026.pdf",
        "03-CMO111 Rapport mensuel de gestion mars 2026.pdf",
        "04-CMO111 Rapport mensuel de gestion avril 2026.pdf"
    ]
    
    template = "CMO111.xlsx"
    all_results = {}
    
    for pdf_file in pdf_files:
        if not os.path.exists(pdf_file):
            print(f"\n⚠️  File not found: {pdf_file}")
            continue
        
        try:
            financial_data, excel_path = process_monthly_pdf(pdf_file)
            
            # Extract month/year from filename
            match = re.search(r'(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s+(\d{4})', pdf_file, re.IGNORECASE)
            if match:
                month_year = f"{match.group(1)} {match.group(2)}"
            else:
                month_year = "Unknown"
            
            # Update template if it exists
            if os.path.exists(template):
                update_template(template, financial_data, month_year)
            
            all_results[month_year] = financial_data
            
        except Exception as e:
            print(f"\n❌ ERROR processing {pdf_file}: {str(e)}")
            import traceback
            traceback.print_exc()
    
    # Final summary
    print(f"\n{'='*80}")
    print(f"PROCESSING COMPLETE - {len(all_results)} months processed")
    print(f"{'='*80}")
    for month, data in all_results.items():
        net = data.get('Net Income', 'N/A')
        print(f"  {month}: Net Income = {net}")
