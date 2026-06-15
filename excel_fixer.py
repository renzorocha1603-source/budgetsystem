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

DONNEES_HISTORIQUES_LABELS = {
    "Transient Revenue": ["revenus horaires"],
    "Monthly Revenues": ["revenus mensuels"],
    "Car-Wash Revenue": ["revenus lave-auto", "lave-auto"],
    "Hotel Revenue": ["revenus hôtel", "revenus hotel"],
    "Interests": ["revenus d'intérêts", "revenus d'interets", "intérêts"],
    "Miscellaneous": ["autres revenus"],
    "Parking Revenue": ["total revenus bruts"],
    "Discount-Gratuities - Transient": ["gratuités", "gratuites"],
    "Discount-Gratuities - Monthly": ["rabais"],
    "TOTAL REVENUE": ["total revenus"],
    "Parking wages": ["salaire stationnement"],
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


def detect_year_from_pnl(pnl_file, parking_code):
    """
    Try to detect what year a P&L file is for.
    Checks the sheet for year references in headers.
    """
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    file_bytes = pnl_file.read() if hasattr(pnl_file, 'read') else None
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    if file_bytes is None:
        return None
    
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        
        # Find the parking code sheet
        for sheet_name in xl.sheet_names:
            if parking_code.upper() in sheet_name.upper():
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                
                # Look for year references in first 10 rows
                for row_idx in range(min(10, len(df))):
                    for col_idx in range(min(20, len(df.columns))):
                        cell = str(df.iloc[row_idx, col_idx])
                        years = re.findall(r'(20\d{2})', cell)
                        if years:
                            # Return the most common year found
                            from collections import Counter
                            year_counts = Counter(years)
                            most_common_year = year_counts.most_common(1)[0][0]
                            return int(most_common_year)
        
        return None
    except:
        return None


def detect_year_from_filename(filename):
    """Try to extract year from filename like 'P&L_2025.xlsx' or '2025_P&L.xlsx'"""
    if not filename:
        return None
    years = re.findall(r'(20\d{2})', filename)
    if years:
        return int(years[0])
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


def find_template_rows_by_label(ws, french_labels, max_rows=100):
    """Search Column A for French labels. Returns list of matching row numbers."""
    matching_rows = []
    for row_idx in range(1, max_rows + 1):
        cell_value = str(ws.cell(row=row_idx, column=1).value or "").strip().lower()
        if not cell_value:
            continue
        cell_clean = cell_value.replace('.', ' ').replace('-', ' ').replace('_', ' ').replace('  ', ' ')
        for french_label in french_labels:
            label_clean = french_label.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ').replace('  ', ' ')
            if label_clean in cell_clean:
                matching_rows.append(row_idx)
                break
    return matching_rows


def read_year_mapping_from_template(ws):
    """
    Read the year mapping from Donnees Historiques template.
    Looks for a row that has years like 2026, 2025, etc. in columns B-M.
    Returns dict: {month_index: year} or None if not found.
    """
    for row_idx in range(35, 50):  # Year mapping is usually around row 40
        year_map = {}
        for col_idx in range(2, 14):  # Columns B-M
            cell_value = str(ws.cell(row=row_idx, column=col_idx).value or "").strip()
            year_match = re.search(r'(20\d{2})', cell_value)
            if year_match:
                year_map[col_idx - 2] = int(year_match.group(1))  # 0-based month index
        
        if len(year_map) >= 6:  # If we found at least 6 months with years
            return year_map
    
    # Fallback: current year for Jan-Apr, previous year for May-Dec
    current_year = datetime.now().year
    year_map = {}
    for i in range(4):
        year_map[i] = current_year
    for i in range(4, 12):
        year_map[i] = current_year - 1
    return year_map


# ============================================================================
# DATA EXTRACTION
# ============================================================================

def extract_pnl_data(pnl_file, parking_code):
    """
    Extract ALL data from the P&L sheet.
    Returns dict with 'monthly' and 'yearly' data.
    """
    df, sheet_name = find_pnl_sheet(pnl_file, parking_code)
    if df is None:
        return None
    
    result = {'monthly': {}, 'yearly': {}}
    
    for row_idx in range(len(df)):
        label = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        if not label or label.lower() in ['code', 'profit & loss', '']:
            continue
        
        if re.search(r'\d{2}-\d{2}-\d{2}', label):
            continue
        
        # Skip rows with no data
        has_any_data = False
        for col_idx in range(1, min(14, len(df.columns))):
            if safe_float(df.iloc[row_idx, col_idx]) != 0:
                has_any_data = True
                break
        
        if not has_any_data:
            continue
        
        # Monthly data (columns B-M = indices 1-12)
        monthly = {}
        for month_idx in range(12):
            col_idx = month_idx + 1
            if col_idx < len(df.columns):
                val = safe_float(df.iloc[row_idx, col_idx])
                monthly[MONTHS_EN[month_idx]] = val
        
        # Yearly total (column N = index 13)
        yearly_total = 0
        if len(df.columns) > 13:
            yearly_total = safe_float(df.iloc[row_idx, 13])
        
        result['monthly'][label] = monthly
        result['yearly'][label] = yearly_total
    
    return result


def merge_monthly_data(current_year_data, previous_year_data, year_map):
    """
    Merge monthly data from current and previous years based on year_map.
    year_map: {month_index: year} - tells which year each month comes from.
    Returns merged monthly data.
    """
    merged = {}
    
    for month_idx, year in year_map.items():
        month_name = MONTHS_EN[month_idx]
        
        # Determine which data source to use
        current_year = datetime.now().year
        
        if year == current_year and current_year_data:
            # Use current year data
            for label, monthly in current_year_data['monthly'].items():
                if label not in merged:
                    merged[label] = {}
                merged[label][month_name] = monthly.get(month_name, 0)
        elif year == current_year - 1 and previous_year_data:
            # Use previous year data
            for label, monthly in previous_year_data['monthly'].items():
                if label not in merged:
                    merged[label] = {}
                merged[label][month_name] = monthly.get(month_name, 0)
    
    return merged


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, previous_year_data, parking_code):
    """Update Budget Initial - S8 with PREVIOUS year TOTAL REVENUE."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["Budget Initial: Sheet not found"]
        
        ws = wb[sheet_name]
        total = previous_year_data['yearly'].get('TOTAL REVENUE', 0) if previous_year_data else 0
        
        if total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f} (previous year)")
        else:
            updates.append("⚠️ Budget Initial: No TOTAL REVENUE found in previous year P&L")
    except Exception as e:
        updates.append(f"❌ Budget Initial: {e}")
    return updates


def update_fiche_stationnement(wb, two_years_ago_data, parking_code, word_data=None):
    """
    Update Fiche Stationnement K17-K26 with data from 2 YEARS AGO.
    For 2027 budget, this is 2025 data.
    """
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        
        if two_years_ago_data is None:
            updates.append("⚠️ Fiche Stationnement: No data from 2 years ago available")
            return updates
        
        for pnl_label, cell in FICHE_STATIONNEMENT_MAP:
            yearly_value = two_years_ago_data['yearly'].get(pnl_label, 0)
            if yearly_value != 0:
                ws[cell] = yearly_value
                ws[cell].number_format = '#,##0.00'
                updates.append(f"✅ {cell} = {pnl_label}: ${yearly_value:,.2f}")
        
        total_revenue = two_years_ago_data['yearly'].get('TOTAL REVENUE', 0)
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


def update_donnees_historiques(wb, merged_monthly_data, parking_code):
    """Update Donnees Historiques with merged monthly data."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        
        # Read year mapping from template
        year_map = read_year_mapping_from_template(ws)
        updates.append(f"📅 Year mapping: Jan={year_map.get(0)}, Apr={year_map.get(3)}, May={year_map.get(4)}, Dec={year_map.get(11)}")
        
        cells_updated = 0
        rows_filled = []
        
        for pnl_label, french_labels in DONNEES_HISTORIQUES_LABELS.items():
            if pnl_label not in merged_monthly_data:
                continue
            
            monthly_values = merged_monthly_data[pnl_label]
            
            if all(v == 0 for v in monthly_values.values()):
                continue
            
            matching_rows = find_template_rows_by_label(ws, french_labels, max_rows=100)
            
            if not matching_rows:
                continue
            
            template_row = matching_rows[0]
            row_cells = 0
            
            for month_idx, month_name in enumerate(MONTHS_EN):
                if month_name in monthly_values:
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
            found_labels = []
            for row_idx in range(1, 100):
                val = str(ws.cell(row=row_idx, column=1).value or "").strip()
                if val and len(val) > 3:
                    found_labels.append(f"R{row_idx}:{val[:30]}")
            updates.append(f"⚠️ No matches. Template labels: {'; '.join(found_labels[:12])}")
    except Exception as e:
        updates.append(f"❌ Donnees Historiques: {e}")
    return updates


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def fix_excel(excel_file, pnl_current_year, pnl_previous_year=None, parking_code=None, word_data=None):
    """
    Main function to process the Excel template with P&L data.
    
    Args:
        excel_file: Uploaded Excel template
        pnl_current_year: P&L file for CURRENT year (e.g., 2026 for 2027 budget)
        pnl_previous_year: P&L file for PREVIOUS year (e.g., 2025 for 2027 budget)
        parking_code: Parking code (extracted from filename if None)
        word_data: Optional dict for Fiche Stationnement H-M rows
    """
    updates = []
    
    if not parking_code and hasattr(excel_file, 'name'):
        parking_code = extract_parking_code_from_filename(excel_file.name)
    
    if not parking_code:
        return None, ["❌ Could not determine parking code from filename"]
    
    updates.append(f"🔍 Processing: {parking_code}")
    
    # Detect years from uploaded files
    current_year_name = "?"
    previous_year_name = "?"
    
    if hasattr(pnl_current_year, 'name'):
        detected = detect_year_from_filename(pnl_current_year.name)
        if detected:
            current_year_name = str(detected)
    
    if pnl_previous_year and hasattr(pnl_previous_year, 'name'):
        detected = detect_year_from_filename(pnl_previous_year.name)
        if detected:
            previous_year_name = str(detected)
    
    updates.append(f"📂 Current year file: {current_year_name} | Previous year file: {previous_year_name}")
    
    # Extract data from both P&L files
    current_year_data = extract_pnl_data(pnl_current_year, parking_code)
    previous_year_data = None
    if pnl_previous_year:
        previous_year_data = extract_pnl_data(pnl_previous_year, parking_code)
    
    if current_year_data is None and previous_year_data is None:
        return None, [f"❌ Could not find P&L data for {parking_code} in any uploaded file"]
    
    # Read template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"❌ Error reading template: {e}"]
    
    # Read year mapping from Donnees Historiques template
    dh_sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
    year_map = None
    if dh_sheet_name:
        year_map = read_year_mapping_from_template(wb[dh_sheet_name])
    
    # Merge monthly data based on year mapping
    merged_monthly = {}
    if year_map and current_year_data and previous_year_data:
        merged_monthly = merge_monthly_data(current_year_data, previous_year_data, year_map)
    elif current_year_data:
        # If only one file, use it for all months
        merged_monthly = current_year_data['monthly']
    
    # For Fiche Stationnement: use data from 2 years ago (previous year if only one file)
    # For now, use previous_year_data if available, otherwise current_year_data
    two_years_ago_data = previous_year_data if previous_year_data else current_year_data
    
    # Update sheets
    updates.extend(update_budget_initial(wb, previous_year_data if previous_year_data else current_year_data, parking_code))
    updates.extend(update_fiche_stationnement(wb, two_years_ago_data, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, merged_monthly, parking_code))
    
    success_count = sum(1 for u in updates if u.startswith("✅"))
    if success_count == 0:
        updates.append("💡 No updates made. Check files and parking code.")
    
    # Reset file pointers
    pnl_current_year.seek(0) if hasattr(pnl_current_year, 'seek') else None
    if pnl_previous_year:
        pnl_previous_year.seek(0) if hasattr(pnl_previous_year, 'seek') else None
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
