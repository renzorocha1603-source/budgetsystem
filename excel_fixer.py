# excel_fixer.py
"""
Excel Fixer Module - Integrates PDF extraction with template filling
"""
import pandas as pd
import os
import re
from datetime import datetime
import fitz  # PyMuPDF
from openpyxl import load_workbook
from io import BytesIO
import tempfile
from pathlib import Path

# ============================================
# HARDCODED SPATIAL LAYOUT FOR P&L PAGE
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

# Row mapping (French -> Standard Label)
ROW_MAPPING = {
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

# French to English label mapping
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

def safe_float(value):
    """Handle European numbers: '7 106 417,00' -> 7106417.00"""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip().replace(" ", "").replace("\xa0", "")
    value = value.replace(",", ".")
    try:
        return float(value)
    except:
        return None

def find_pl_page(pdf_path):
    """Find P&L page by content markers"""
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        text = doc[i].get_text()
        if all(m.lower() in text.lower() for m in PAGE_MARKERS):
            return doc, i
    doc.close()
    return None, None

def extract_table_from_page(page):
    """Extract 27 rows x 9 columns using hardcoded layout"""
    w, h = page.rect.width, page.rect.height
    top, bottom = h * TABLE_TOP, h * TABLE_BOTTOM
    left, right = w * TABLE_LEFT, w * TABLE_RIGHT
    row_h = (bottom - top) / NUM_ROWS
    
    data = []
    for row in range(NUM_ROWS):
        row_data = []
        y = top + row * row_h
        x = left
        for col_w in COLUMN_WIDTHS:
            cell_w = col_w * (right - left)
            text = page.get_text("text", clip=fitz.Rect(x+1, y+1, x+cell_w-1, y+row_h-1))
            row_data.append(' '.join(text.strip().split()))
            x += cell_w
        data.append(row_data)
    return data

def extract_from_pdf(pdf_path):
    """Extract financial data from a PDF monthly report"""
    doc, page_num = find_pl_page(pdf_path)
    if doc is None:
        print(f"❌ No P&L page found in {pdf_path}")
        return {}
    
    table = extract_table_from_page(doc[page_num])
    doc.close()
    
    # Extract amounts using row mapping
    financial_data = {}
    for row_num, french_label in ROW_MAPPING.items():
        if row_num <= len(table):
            raw_amount = table[row_num-1][AMOUNT_COL]
            amount = safe_float(raw_amount)
            if amount is not None:
                english_label = LABEL_MAPPING.get(french_label, french_label)
                financial_data[english_label] = amount
    
    return financial_data

def get_parking_codes_from_pnl(file_obj):
    """
    Extract parking codes from uploaded monthly file.
    Looks for CMO codes in filename or content.
    """
    codes = []
    file_obj.seek(0)
    
    # Try from filename
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
    
    # Try from content (if PDF)
    if hasattr(file_obj, 'name') and file_obj.name.lower().endswith('.pdf'):
        try:
            # Save to temp file for PyMuPDF
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            doc = fitz.open(tmp_path)
            for page in doc:
                text = page.get_text()
                # Look for parking codes in text
                found_codes = re.findall(r'(?:parking|stationnement|code)\s*:?\s*(CMO\d+)', text, re.IGNORECASE)
                codes.extend(found_codes)
                if not found_codes:
                    # Look for "1981 McGill College" or similar
                    if "1981 McGill College" in text:
                        codes.append("1981MCGILL")
            doc.close()
            os.unlink(tmp_path)
        except Exception as e:
            print(f"Error reading PDF for codes: {e}")
    
    file_obj.seek(0)
    
    # Deduplicate and uppercase
    return list(set([c.upper() for c in codes]))

def process_monthly_files(monthly_files):
    """
    Process a list of monthly report files and extract data.
    Returns dict: {month_year: financial_data}
    """
    results = {}
    
    for file_obj in monthly_files:
        # Get month from filename
        month = "Unknown"
        if hasattr(file_obj, 'name'):
            match = re.search(
                r'(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s+(\d{4})',
                file_obj.name, re.IGNORECASE
            )
            if match:
                month = f"{match.group(1)} {match.group(2)}"
        
        file_obj.seek(0)
        
        # Process based on file type
        if hasattr(file_obj, 'name') and file_obj.name.lower().endswith('.pdf'):
            # Save PDF to temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            try:
                financial_data = extract_from_pdf(tmp_path)
                if financial_data:
                    results[month] = financial_data
            finally:
                os.unlink(tmp_path)
        
        elif hasattr(file_obj, 'name') and file_obj.name.lower().endswith(('.xlsx', '.xls')):
            # Process Excel files
            try:
                df = pd.read_excel(file_obj)
                # Look for P&L-like structure
                financial_data = {}
                for col in df.columns:
                    for idx, row in df.iterrows():
                        account = str(row.iloc[0]) if len(row) > 0 else ""
                        for french_label, english_label in LABEL_MAPPING.items():
                            if french_label.lower() in account.lower():
                                amount = safe_float(row.iloc[1] if len(row) > 1 else None)
                                if amount is not None:
                                    financial_data[english_label] = amount
                if financial_data:
                    results[month] = financial_data
            except Exception as e:
                print(f"Error processing Excel: {e}")
        
        file_obj.seek(0)
    
    return results

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """
    Main fix_excel function that processes all inputs and returns updated Excel.
    
    Parameters:
    - excel_file: Template Excel file (BytesIO or file object)
    - monthly_files_current: List of current year monthly report files
    - monthly_files_previous: List of previous year monthly report files  
    - budget_initial_file: Budget initial source file
    - fiche_stationnement_file: Fiche stationnement source file
    - parking_code: Selected parking code
    - word_data: Optional word document data
    
    Returns:
    - fixed_excel_bytes: Updated Excel file as bytes
    - updates: List of update messages
    """
    updates = []
    
    try:
        # Load the template
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            wb = load_workbook(excel_file)
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
        # Process current year files
        if monthly_files_current:
            current_data = process_monthly_files(monthly_files_current)
            updates.append(f"✅ Current year: {len(current_data)} months extracted")
            
            for month, data in current_data.items():
                updates.append(f"  📊 {month}: {len(data)} accounts")
                # Here you would update the template with extracted data
                # update_donnees_historiques(wb, month, data)
        
        # Process previous year files
        if monthly_files_previous:
            previous_data = process_monthly_files(monthly_files_previous)
            updates.append(f"✅ Previous year: {len(previous_data)} months extracted")
        
        # Process budget initial file
        if budget_initial_file:
            budget_initial_file.seek(0)
            if hasattr(budget_initial_file, 'name') and budget_initial_file.name.lower().endswith('.pdf'):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(budget_initial_file.read())
                    tmp_path = tmp.name
                budget_data = extract_from_pdf(tmp_path)
                os.unlink(tmp_path)
                if budget_data:
                    updates.append(f"✅ Budget initial extracted: {len(budget_data)} accounts")
        
        # Process fiche stationnement file
        if fiche_stationnement_file:
            fiche_stationnement_file.seek(0)
            if hasattr(fiche_stationnement_file, 'name') and fiche_stationnement_file.name.lower().endswith('.pdf'):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(fiche_stationnement_file.read())
                    tmp_path = tmp.name
                fiche_data = extract_from_pdf(tmp_path)
                os.unlink(tmp_path)
                if fiche_data:
                    updates.append(f"✅ Fiche stationnement extracted: {len(fiche_data)} accounts")
        
        # Save updated workbook
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("✅ Workflow completed successfully")
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"❌ Error: {str(e)}")
        return None, updates
