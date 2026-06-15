import io
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

SEASONAL_MULTIPLIERS = {
    "SC": [1.2, 1.2, 1.2, 1.1, 0.8, 0.8, 0.8, 0.8, 1.1, 1.1, 1.2, 1.2],
    "RG": [0.8, 0.8, 0.9, 0.9, 1.3, 1.3, 1.3, 1.2, 1.0, 1.0, 0.8, 0.8]
}

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

MONTHS_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SHEET_PATTERNS = {
    "Budget Initial": ["budget initial", "budget"],
    "Fiche Stationnement": ["fiche stationnement", "fiche de stationnement", "stationnement"],
    "Donnees Historiques": ["donnees historiques", "données historiques", "historiques", "historique"],
}

PARKING_CODE_PATTERNS = [
    r'(CMO\d+)',
    r'(LUNA)',
    r'([A-Z]{2,5}\d{2,4})',
]

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    if not filename:
        return None
    name = filename.rsplit('.', 1)[0]
    for pattern in PARKING_CODE_PATTERNS:
        match = re.search(pattern, name, re.IGNORECASE)
        if match:
            return match.group(1).upper()
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


def find_parking_row_in_df(df, parking_code):
    """Find the row containing a parking code in ANY column."""
    code_clean = parking_code.upper().strip()
    
    for col_idx in range(len(df.columns)):
        for row_idx in range(len(df)):
            value = df.iloc[row_idx, col_idx]
            if pd.notna(value):
                cell_str = str(value).upper().strip()
                if code_clean in cell_str:
                    return row_idx
                # Partial match for CMO codes
                if code_clean.startswith('CMO') and cell_str.startswith('CMO'):
                    if code_clean[-3:] in cell_str or cell_str[-3:] in code_clean:
                        return row_idx
    return None


def find_all_parking_codes_in_df(df):
    codes = []
    for col_idx in range(len(df.columns)):
        for row_idx in range(len(df)):
            value = df.iloc[row_idx, col_idx]
            if pd.notna(value):
                cell_str = str(value).strip()
                for pattern in PARKING_CODE_PATTERNS:
                    match = re.search(pattern, cell_str, re.IGNORECASE)
                    if match:
                        codes.append(match.group(1).upper())
    return list(set(codes))


def safe_float(value, default=0.0):
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '').replace('€', '').replace('(', '').replace(')', '')
        return float(value)
    except (ValueError, TypeError):
        return default


def find_year_columns(df, max_cols=50):
    """Find columns that contain year data (2024, 2025, etc.)"""
    year_cols = {}
    for col_idx in range(min(max_cols, len(df.columns))):
        for row_idx in range(min(15, len(df))):
            cell = str(df.iloc[row_idx, col_idx]).strip()
            matches = re.findall(r'(20\d{2})', cell)
            for year_str in matches:
                year = int(year_str)
                if 2020 <= year <= 2030:
                    year_cols[year] = col_idx
    return year_cols


def find_month_columns_in_df(df, max_cols=20):
    """Find column indices for each month."""
    month_cols = {}
    for col_idx in range(1, min(max_cols, len(df.columns))):
        for row_idx in range(min(10, len(df))):
            cell = str(df.iloc[row_idx, col_idx]).strip().lower()
            for i, month in enumerate(MONTHS):
                if month.lower() in cell:
                    month_cols[i] = col_idx
            for i, month in enumerate(MONTHS_ABBR):
                if month.lower() == cell[:3] and len(cell) <= 4:
                    month_cols[i] = col_idx
    if not month_cols:
        for i in range(12):
            if i + 1 < len(df.columns):
                month_cols[i] = i + 1
    return month_cols


def read_all_pnl_sheets(pnl_file):
    """Read ALL sheets from P&L file and return list of (sheet_name, DataFrame)."""
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    file_bytes = pnl_file.read() if hasattr(pnl_file, 'read') else None
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    if file_bytes is None:
        return []
    
    sheets = []
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                sheets.append((sheet_name, df))
            except:
                pass
    except Exception as e:
        print(f"Error reading P&L: {e}")
    
    return sheets


# ============================================================================
# DATA EXTRACTION - Searches ALL sheets independently
# ============================================================================

def extract_budget_initial_total(pnl_file, parking_code):
    """Extract total for Budget Initial - searches ALL sheets."""
    sheets = read_all_pnl_sheets(pnl_file)
    
    for sheet_name, df in sheets:
        row = find_parking_row_in_df(df, parking_code)
        if row is None:
            continue
        
        print(f"Budget Initial: Found {parking_code} in sheet '{sheet_name}' at row {row}")
        
        # Get year columns
        year_cols = find_year_columns(df)
        if not year_cols:
            # Use last numeric column
            for col_idx in range(len(df.columns) - 1, 0, -1):
                if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                    val = safe_float(df.iloc[row, col_idx])
                    if val > 0:
                        return val
            continue
        
        # Try to find a "Total" row near this parking
        latest_year = max(year_cols.keys())
        total_col = year_cols[latest_year]
        
        # First try the parking row itself
        total = safe_float(df.iloc[row, total_col])
        if total > 0:
            return total
        
        # Search for "Total" label in rows below parking
        for search_row in range(row, min(row + 100, len(df))):
            if pd.notna(df.iloc[search_row, 0]):
                label = str(df.iloc[search_row, 0]).strip().lower()
                if any(word in label for word in ['total', 'grand total', 'sum']):
                    val = safe_float(df.iloc[search_row, total_col])
                    if val > 0:
                        return val
        
        # Sum all numeric values in parking row
        total = 0
        for col_idx in range(1, len(df.columns)):
            val = safe_float(df.iloc[row, col_idx])
            total += abs(val)
        if total > 0:
            return total
    
    return None


def extract_revenue_categories(pnl_file, parking_code):
    """Extract revenue categories - searches ALL sheets."""
    sheets = read_all_pnl_sheets(pnl_file)
    
    for sheet_name, df in sheets:
        row = find_parking_row_in_df(df, parking_code)
        if row is None:
            continue
        
        print(f"Revenue: Found {parking_code} in sheet '{sheet_name}' at row {row}")
        
        # Find year columns
        year_cols = find_year_columns(df)
        if not year_cols:
            # Try last numeric column
            for col_idx in range(len(df.columns) - 1, 0, -1):
                if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                    revenue_col = col_idx
                    break
            else:
                continue
        else:
            latest_year = max(year_cols.keys())
            revenue_col = year_cols[latest_year]
        
        print(f"Revenue: Using column {revenue_col}")
        
        # Print ALL rows around this parking for debugging
        print(f"Revenue: Data around row {row}:")
        for debug_row in range(max(0, row - 2), min(row + 20, len(df))):
            cat = str(df.iloc[debug_row, 0]).strip() if pd.notna(df.iloc[debug_row, 0]) else ""
            val = safe_float(df.iloc[debug_row, revenue_col])
            if cat or val > 0:
                print(f"  Row {debug_row}: [{cat}] = ${val:,.2f}")
        
        # Extract categories from this parking's section
        revenue_data = {}
        for data_row in range(row + 1, min(row + 200, len(df))):
            category = str(df.iloc[data_row, 0]).strip() if pd.notna(df.iloc[data_row, 0]) else ""
            value = safe_float(df.iloc[data_row, revenue_col])
            
            # Stop conditions
            if re.match(r'^[A-Z]{2,5}\d{2,6}$', category.upper()):
                break
            if not category and value == 0:
                empty_count = 0
                for check in range(data_row, min(data_row + 5, len(df))):
                    c = str(df.iloc[check, 0]).strip() if pd.notna(df.iloc[check, 0]) else ""
                    v = safe_float(df.iloc[check, revenue_col])
                    if not c and v == 0:
                        empty_count += 1
                if empty_count >= 3:
                    break
            
            if category and category.lower() not in ['total', 'grand total', 'sum', '']:
                if value > 0:
                    revenue_data[category] = value
        
        if revenue_data:
            return revenue_data
    
    return {}


def extract_monthly_data(pnl_file, parking_code):
    """Extract monthly data - searches ALL sheets."""
    sheets = read_all_pnl_sheets(pnl_file)
    
    for sheet_name, df in sheets:
        row = find_parking_row_in_df(df, parking_code)
        if row is None:
            continue
        
        print(f"Monthly: Found {parking_code} in sheet '{sheet_name}' at row {row}")
        
        month_cols = find_month_columns_in_df(df)
        
        if len(month_cols) < 3:
            # If month columns not found, try columns 1-12
            for i in range(12):
                if i + 1 < len(df.columns):
                    month_cols[i] = i + 1
        
        # Debug: show what's in the month columns
        print(f"Monthly: Using month columns: {month_cols}")
        for month_idx, col_idx in month_cols.items():
            if month_idx < len(MONTHS):
                val = safe_float(df.iloc[row, col_idx])
                if val > 0:
                    print(f"  {MONTHS[month_idx]}: ${val:,.2f}")
        
        monthly_data = {}
        for month_idx, col_idx in month_cols.items():
            if month_idx < len(MONTHS) and col_idx < len(df.columns):
                value = safe_float(df.iloc[row, col_idx])
                if value > 0:
                    monthly_data[MONTHS[month_idx]] = value
        
        if monthly_data:
            return monthly_data
    
    return {}


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, pnl_file, parking_code):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["Budget Initial: Sheet not found"]
        
        ws = wb[sheet_name]
        total = extract_budget_initial_total(pnl_file, parking_code)
        
        if total and total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"✅ Budget Initial: S8 = ${total:,.2f}")
        else:
            updates.append(f"⚠️ Budget Initial: No total found for {parking_code}")
    except Exception as e:
        updates.append(f"❌ Budget Initial: {e}")
    return updates


def update_fiche_stationnement(wb, pnl_file, parking_code, word_data=None):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        revenue_data = extract_revenue_categories(pnl_file, parking_code)
        
        if not revenue_data:
            return [f"⚠️ Fiche Stationnement: No revenue data for {parking_code}"]
        
        sorted_categories = sorted(revenue_data.items(), key=lambda x: x[1], reverse=True)
        
        total_revenue = 0
        for i, (category, value) in enumerate(sorted_categories):
            if i >= 9:
                break
            cell = f"K{17 + i}"
            ws[cell] = value
            ws[cell].number_format = '#,##0.00'
            total_revenue += value
            updates.append(f"✅ {cell} = {category}: ${value:,.2f}")
        
        ws["K26"] = total_revenue
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"✅ K26 = Total: ${total_revenue:,.2f}")
        
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


def update_donnees_historiques(wb, pnl_file, parking_code):
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        monthly_data = extract_monthly_data(pnl_file, parking_code)
        
        if not monthly_data:
            return [f"⚠️ Donnees Historiques: No monthly data for {parking_code}"]
        
        yellow_rows = [r for r in range(36, 77) if r not in [44, 47, 65]]
        columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        
        cells_updated = 0
        for row in yellow_rows:
            for col_idx, col_letter in enumerate(columns):
                month_name = MONTHS[col_idx] if col_idx < len(MONTHS) else None
                if month_name and month_name in monthly_data and monthly_data[month_name] > 0:
                    cell = f"{col_letter}{row}"
                    ws[cell] = monthly_data[month_name]
                    ws[cell].number_format = '#,##0.00'
                    cells_updated += 1
        
        if cells_updated > 0:
            updates.append(f"✅ Updated {cells_updated} monthly cells")
        else:
            updates.append("⚠️ No monthly data to update")
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
    
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"❌ Error reading template: {e}"]
    
    updates.append(f"📋 Template has {len(wb.sheetnames)} sheets")
    
    # Check what's in the P&L file
    sheets = read_all_pnl_sheets(pnl_file)
    updates.append(f"📊 P&L file has {len(sheets)} sheets")
    
    all_codes = []
    for s_name, s_df in sheets:
        codes = find_all_parking_codes_in_df(s_df)
        all_codes.extend(codes)
    all_codes = list(set(all_codes))
    updates.append(f"🏷️ P&L codes: {', '.join(all_codes[:15])}")
    
    # Process sheets
    updates.extend(update_budget_initial(wb, pnl_file, parking_code))
    updates.extend(update_fiche_stationnement(wb, pnl_file, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, pnl_file, parking_code))
    
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
