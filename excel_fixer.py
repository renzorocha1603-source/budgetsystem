import io
import re
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

MONTHS_EN = ["January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]

SHEET_PATTERNS = {
    "Budget Initial": ["budget initial", "budget"],
    "Fiche Stationnement": ["fiche stationnement", "fiche de stationnement", "stationnement"],
    "Donnees Historiques": ["donnees historiques", "données historiques", "historiques", "historique"],
}

# ONLY map rows that have EXACT P&L matches - do NOT fill rows without data
DONNEES_HISTORIQUES_MAP = {
    "Transient Revenue": 42,
    "Monthly Revenues": 43,
    "Car-Wash Revenue": 45,
    "Hotel Revenue": 46,
    "Interests": 48,
    "Miscellaneous": 49,
    "Parking Revenue": 50,
    "Discount-Gratuities - Transient": 52,
    "Discount-Gratuities - Monthly": 54,
    "TOTAL REVENUE": 58,
    "Parking wages": 61,
}

FICHE_STATIONNEMENT_MAP = [
    ("Parking Revenue", "K17"),
    ("Monthly Revenues", "K18"),
    ("Transient Revenue", "K19"),
    ("TOTAL REVENUE", "K20"),
    ("Total Operation expenses", "K21"),
    ("Parking wages", "K22"),
    ("OPERATION SURPLUS", "K23"),
    ("Percent Management fee", "K24"),
]

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
        
        for sheet_name in xl.sheet_names:
            if sheet_name.upper().strip() == parking_code.upper().strip():
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                return df, sheet_name
        
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


# ============================================================================
# DATA EXTRACTION
# ============================================================================

def extract_pnl_data(pnl_file, parking_code):
    """
    Extract ALL data from the P&L sheet.
    P&L structure: Col A = labels, Col B-M = Jan-Dec (indices 1-12), Col N = Year Total (index 13)
    Returns dict with 'monthly' and 'yearly' data.
    """
    df, sheet_name = find_pnl_sheet(pnl_file, parking_code)
    if df is None:
        return None
    
    print(f"P&L sheet: '{sheet_name}', shape: {df.shape}")
    
    result = {'monthly': {}, 'yearly': {}}
    
    # Track which labels we've seen to avoid duplicates (take the LAST one with data)
    seen_labels = {}
    
    for row_idx in range(len(df)):
        label = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        if not label or label.lower() in ['code', 'profit & loss', '']:
            continue
        
        # Skip filter/date rows
        if re.search(r'\d{2}-\d{2}-\d{2}', label):
            continue
        
        # Skip section headers (rows with no data in any column)
        has_any_data = False
        for col_idx in range(1, min(14, len(df.columns))):
            if safe_float(df.iloc[row_idx, col_idx]) != 0:
                has_any_data = True
                break
        
        if not has_any_data:
            continue
        
        # Extract monthly data (columns B-M = indices 1-12)
        monthly = {}
        for month_idx in range(12):
            col_idx = month_idx + 1  # Column B=1, C=2, ... M=12
            if col_idx < len(df.columns):
                val = safe_float(df.iloc[row_idx, col_idx])
                monthly[MONTHS_EN[month_idx]] = val
        
        # Extract yearly total (column N = index 13)
        yearly_total = 0
        if len(df.columns) > 13:
            yearly_total = safe_float(df.iloc[row_idx, 13])
        
        # Store (overwrites if same label appears again with data)
        result['monthly'][label] = monthly
        result['yearly'][label] = yearly_total
        
        if yearly_total > 0:
            print(f"  {label}: Yearly=${yearly_total:,.2f}, Jan=${monthly.get('January', 0):,.2f}")
    
    return result


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, pnl_data, parking_code):
    """Update Budget Initial - S8 with TOTAL REVENUE yearly total."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["Budget Initial: Sheet not found"]
        
        ws = wb[sheet_name]
        total = pnl_data['yearly'].get('TOTAL REVENUE', 0)
        
        if total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f}")
        else:
            updates.append("⚠️ Budget Initial: No TOTAL REVENUE found")
    except Exception as e:
        updates.append(f"❌ Budget Initial: {e}")
    return updates


def update_fiche_stationnement(wb, pnl_data, parking_code, word_data=None):
    """Update Fiche Stationnement K17-K26 with YEARLY totals (Column N)."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        
        for pnl_label, cell in FICHE_STATIONNEMENT_MAP:
            yearly_value = pnl_data['yearly'].get(pnl_label, 0)
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00'
                updates.append(f"✅ {cell} = {pnl_label}: ${yearly_value:,.2f}")
            else:
                updates.append(f"⚠️ {cell} = {pnl_label}: No data (skipped)")
        
        # K26 = TOTAL REVENUE
        total_revenue = pnl_data['yearly'].get('TOTAL REVENUE', 0)
        ws["K26"] = total_revenue
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"✅ K26 = TOTAL REVENUE: ${total_revenue:,.2f}")
        
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
    ONLY fills rows that have EXACT P&L label matches.
    Does NOT touch rows without matching data.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        cells_updated = 0
        rows_filled = []
        
        for pnl_label, template_row in DONNEES_HISTORIQUES_MAP.items():
            # Check if this label exists in P&L monthly data
            if pnl_label not in pnl_data['monthly']:
                continue
            
            monthly_values = pnl_data['monthly'][pnl_label]
            
            # Skip if all values are zero
            if all(v == 0 for v in monthly_values.values()):
                continue
            
            # Fill each month column (B-M)
            row_cells = 0
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values and monthly_values[month_name] != 0:
                    col_letter = chr(ord('B') + month_idx)
                    cell = f"{col_letter}{template_row}"
                    ws[cell] = monthly_values[month_name]
                    ws[cell].number_format = '#,##0.00'
                    cells_updated += 1
                    row_cells += 1
            
            if row_cells > 0:
                rows_filled.append(f"  Row {template_row}: {pnl_label} ({row_cells} months)")
        
        if cells_updated > 0:
            updates.append(f"✅ Donnees Historiques: {cells_updated} cells in {len(rows_filled)} rows")
            for row_info in rows_filled:
                updates.append(row_info)
        else:
            available = list(pnl_data['monthly'].keys())
            updates.append(f"⚠️ Donnees Historiques: No matches. Available P&L labels ({len(available)}): {', '.join(available[:8])}...")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {e}")
    return updates


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(excel_file, pnl_file, parking_code=None, word_data=None):
    updates = []
    
    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)
    
    if not parking_code:
        return None, ["❌ Could not determine parking code from filename"]
    
    updates.append(f"🔍 Processing: {parking_code}")
    
    # Extract data from P&L
    pnl_data = extract_pnl_data(pnl_file, parking_code)
    if pnl_data is None:
        return None, [f"❌ Could not find P&L data for {parking_code}"]
    
    yearly_count = sum(1 for v in pnl_data['yearly'].values() if v != 0)
    monthly_count = sum(1 for m in pnl_data['monthly'].values() if any(v != 0 for v in m.values()))
    updates.append(f"📊 P&L: {yearly_count} yearly totals, {monthly_count} monthly rows")
    
    # Read template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"❌ Error reading template: {e}"]
    
    # Process sheets
    updates.extend(update_budget_initial(wb, pnl_data, parking_code))
    updates.extend(update_fiche_stationnement(wb, pnl_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, pnl_data, parking_code))
    
    success_count = sum(1 for u in updates if u.startswith("✅"))
    if success_count == 0:
        updates.append("💡 No updates made. Check parking code.")
    
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
