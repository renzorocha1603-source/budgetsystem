import io
import re
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

MONTHS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
             "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

MONTHS_EN = ["January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]

# Template row mapping for Donnees Historiques
# Maps P&L row labels to template row numbers
TEMPLATE_ROW_MAP = {
    # Revenue section
    "Revenus horaires": 42,
    "Revenus mensuels": 43,
    "Revenus Lave-auto": 45,
    "Revenus hôtel": 46,
    "Revenus d'intérêts": 48,
    "Autres revenus": 49,
    "Total revenus Bruts": 50,
    "(Gratuités)": 52,
    "(Rabais)": 54,
    "TOTAL REVENUS": 58,
    # Expense section
    "Salaire Stationnement": 61,
    "Salaire Superviseur": 62,
    "Formation & Recrutement": 63,
    "Uniformes": 64,
    "Total Frais de personnel": 65,
    "Nettoyage stationnement": 67,
    "Entretien stationnement": 68,
    "Entretien équipement": 69,
    "Signalisation": 70,
    "Lignage": 71,
    "Déneigement": 72,
    "Fournitures stationnement": 73,
    "Total Entretien - réparations": 74,
}

# P&L row labels to match against template categories
PNL_CATEGORY_MAP = {
    # Revenue mappings (P&L label -> template category)
    "Parking Revenue": "Total revenus Bruts",
    "Monthly Revenues": "Revenus mensuels",
    "Transient Revenue": "Revenus horaires",
    "Hotel Revenue": "Revenus hôtel",
    "Car-Wash Revenue": "Revenus Lave-auto",
    "Interests": "Revenus d'intérêts",
    "Miscellaneous": "Autres revenus",
    "Discount-Gratuities - Transient": "(Gratuités)",
    "Discount-Gratuities - Monthly": "(Rabais)",
    "TOTAL REVENUE": "TOTAL REVENUS",
    # Expense mappings
    "Parking wages": "Salaire Stationnement",
    "Total Operation expenses": "Total Entretien - réparations",
    "OPERATION SURPLUS": None,  # Skip - not needed in template
    "NET INCOME": None,  # Skip
}

SHEET_PATTERNS = {
    "Budget Initial": ["budget initial", "budget"],
    "Fiche Stationnement": ["fiche stationnement", "fiche de stationnement", "stationnement"],
    "Donnees Historiques": ["donnees historiques", "données historiques", "historiques", "historique"],
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    if not filename:
        return None
    name = filename.rsplit('.', 1)[0]
    match = re.search(r'(CMO\d+)', name, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    if 'LUNA' in name.upper():
        return 'LUNA'
    parts = name.replace('(', ' ').replace(')', ' ').split('_')[0].split()[0]
    return parts.upper()


def find_sheet_by_pattern(wb, patterns):
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            pattern_clean = pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
            if pattern_clean in sheet_lower:
                return sheet_name
    return None


def find_pnl_sheet(pnl_file, parking_code):
    """Find the P&L sheet for a specific parking code."""
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    file_bytes = pnl_file.read() if hasattr(pnl_file, 'read') else None
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    if file_bytes is None:
        return None, None
    
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        
        # Exact match first
        for sheet_name in xl.sheet_names:
            if sheet_name.upper().strip() == parking_code.upper().strip():
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                return df, sheet_name
        
        # Partial match
        for sheet_name in xl.sheet_names:
            if parking_code.upper() in sheet_name.upper():
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                return df, sheet_name
        
        return None, None
    except Exception as e:
        print(f"Error reading P&L: {e}")
        return None, None


def safe_float(value, default=0.0):
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '').replace('€', '').replace('(', '').replace(')', '')
            if value.startswith('-'):
                value = value[1:]
                return -safe_float(value, default)
        return float(value)
    except (ValueError, TypeError):
        return default


def find_row_by_label(df, label, col_idx=0, max_rows=100):
    """Find a row by its label in a specific column."""
    label_lower = label.lower().strip()
    for row_idx in range(min(max_rows, len(df))):
        cell = str(df.iloc[row_idx, col_idx]).strip().lower() if pd.notna(df.iloc[row_idx, col_idx]) else ""
        if label_lower in cell:
            return row_idx
    return None


# ============================================================================
# DATA EXTRACTION
# ============================================================================

def extract_pnl_data(pnl_file, parking_code):
    """
    Extract ALL data from the P&L sheet into a structured format.
    Returns dict with:
    - 'monthly': {row_label: {month_name: value}}
    - 'yearly': {row_label: total_value}
    """
    df, sheet_name = find_pnl_sheet(pnl_file, parking_code)
    if df is None:
        return None
    
    result = {'monthly': {}, 'yearly': {}}
    
    # The P&L has: Col A = labels, Col B-M = Jan-Dec, Col N = Year Total
    # Data starts around row 8-9
    
    for row_idx in range(len(df)):
        label = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        if not label:
            continue
        
        # Extract monthly data (columns B-M = indices 1-12)
        monthly = {}
        for month_idx in range(12):
            if month_idx + 1 < len(df.columns):
                val = safe_float(df.iloc[row_idx, month_idx + 1])
                monthly[MONTHS_EN[month_idx]] = val
        
        # Extract yearly total (column N = index 13)
        yearly_total = safe_float(df.iloc[row_idx, 13]) if len(df.columns) > 13 else 0
        
        # Only store if there's actual data
        has_data = any(v != 0 for v in monthly.values()) or yearly_total != 0
        if has_data:
            result['monthly'][label] = monthly
            result['yearly'][label] = yearly_total
    
    return result


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, pnl_data, parking_code):
    """Update Budget Initial - cell S8 with TOTAL REVENUE yearly total."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["Budget Initial: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # Get TOTAL REVENUE from P&L
        total = pnl_data['yearly'].get('TOTAL REVENUE', 0)
        
        if total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f}")
        else:
            updates.append(f"⚠️ Budget Initial: No TOTAL REVENUE found")
    except Exception as e:
        updates.append(f"❌ Budget Initial: {e}")
    return updates


def update_fiche_stationnement(wb, pnl_data, parking_code, word_data=None):
    """
    Update Fiche Stationnement - K17-K26 with yearly totals.
    Maps specific P&L categories to K17-K25.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # Define which P&L categories go to which cells (in order of importance)
        cell_mapping = [
            ("Parking Revenue", "K17"),
            ("TOTAL REVENUE", "K18"),
            ("Monthly Revenues", "K19"),
            ("Transient Revenue", "K20"),
            ("Total Operation expenses", "K21"),
            ("Parking wages", "K22"),
            ("OPERATION SURPLUS", "K23"),
            ("Percent Management fee", "K24"),
            ("NET INCOME", "K25"),
        ]
        
        total_revenue = 0
        for pnl_label, cell in cell_mapping:
            value = pnl_data['yearly'].get(pnl_label, 0)
            if value != 0:
                ws[cell] = value
                ws[cell].number_format = '#,##0.00'
                total_revenue += value if value > 0 else 0
                updates.append(f"✅ {cell} = {pnl_label}: ${value:,.2f}")
        
        # K26 = Sum of all positive values or TOTAL REVENUE
        ws["K26"] = pnl_data['yearly'].get('TOTAL REVENUE', total_revenue)
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"✅ K26 = Total Revenue: ${pnl_data['yearly'].get('TOTAL REVENUE', 0):,.2f}")
        
        if word_data:
            for row in range(43, 56):
                if "Nb abonnés" in word_data:
                    ws[f"H{row}"] = word_data["Nb abonnés"]
                if "Informations" in word_data:
                    ws[f"I{row}"] = word_data["Informations"]
                if "Avant taxes" in word_data:
                    ws[f"J{row}"] = word_data["Avant taxes"]
                    ws[f"L{row}"] = word_data["Avant taxes"]
            updates.append("✅ Updated Word data")
    except Exception as e:
        updates.append(f"❌ Fiche Stationnement: {e}")
    return updates


def update_donnees_historiques(wb, pnl_data, parking_code):
    """
    Update Donnees Historiques with monthly data.
    Maps P&L monthly values to the correct template rows.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # Map P&L data to template rows using PNL_CATEGORY_MAP
        cells_updated = 0
        
        for pnl_label, template_label in PNL_CATEGORY_MAP.items():
            if template_label is None:
                continue
            
            # Get the template row for this label
            template_row = TEMPLATE_ROW_MAP.get(template_label)
            if template_row is None:
                continue
            
            # Get monthly data from P&L
            monthly_values = pnl_data['monthly'].get(pnl_label, {})
            if not monthly_values:
                continue
            
            # Fill each month column (B-M = columns 1-12 in template)
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values and monthly_values[month_name] != 0:
                    col_letter = chr(ord('B') + month_idx)  # B, C, D, ... M
                    cell = f"{col_letter}{template_row}"
                    ws[cell] = monthly_values[month_name]
                    ws[cell].number_format = '#,##0.00'
                    cells_updated += 1
        
        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: Updated {cells_updated} monthly cells")
        else:
            updates.append("⚠️ Donnees Historiques: No matching data found")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {e}")
    return updates


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(excel_file, pnl_file, parking_code=None, word_data=None):
    """
    Main function to process the Excel template with P&L data.
    
    Args:
        excel_file: Uploaded Excel template (file-like object)
        pnl_file: Uploaded P&L file (file-like object) 
        parking_code: Parking code to process (extracted from filename if None)
        word_data: Optional dict with data for Fiche Stationnement H-M rows
    
    Returns:
        (BytesIO with updated Excel, list of update messages)
    """
    updates = []
    
    # Extract parking code from filename if not provided
    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)
    
    if not parking_code:
        return None, ["❌ Could not determine parking code from filename"]
    
    updates.append(f"🔍 Processing: {parking_code}")
    
    # Extract ALL data from P&L file
    pnl_data = extract_pnl_data(pnl_file, parking_code)
    if pnl_data is None:
        return None, [f"❌ Could not find P&L data for {parking_code}"]
    
    yearly_count = sum(1 for v in pnl_data['yearly'].values() if v != 0)
    monthly_count = sum(1 for m in pnl_data['monthly'].values() if any(v != 0 for v in m.values()))
    updates.append(f"📊 Found {yearly_count} yearly totals, {monthly_count} monthly breakdowns")
    
    # Read the Excel template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"❌ Error reading template: {e}"]
    
    # Process each sheet
    updates.extend(update_budget_initial(wb, pnl_data, parking_code))
    updates.extend(update_fiche_stationnement(wb, pnl_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, pnl_data, parking_code))
    
    # Count successes
    success_count = sum(1 for u in updates if u.startswith("✅"))
    if success_count == 0:
        updates.append("💡 No updates were made.")
    
    # Reset P&L file pointer
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
