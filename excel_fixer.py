import io
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

# Seasonal multipliers for monthly projections
SEASONAL_MULTIPLIERS = {
    "SC": [1.2, 1.2, 1.2, 1.1, 0.8, 0.8, 0.8, 0.8, 1.1, 1.1, 1.2, 1.2],
    "RG": [0.8, 0.8, 0.9, 0.9, 1.3, 1.3, 1.3, 1.2, 1.0, 1.0, 0.8, 0.8]
}

# Month names for column mapping
MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

MONTHS_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Sheet name patterns to match in template
SHEET_PATTERNS = {
    "Budget Initial": ["budget initial", "budget"],
    "Fiche Stationnement": ["fiche stationnement", "fiche de stationnement", "stationnement"],
    "Donnees Historiques": ["donnees historiques", "données historiques", "historiques", "historique"],
    "Revenu Depense": ["revenu depense", "revenu-depense", "revenu dépense", "revenu"],
    "Calcul Salaire": ["calcul salaire", "calcul-salaire", "salaire"],
}

# Parking code patterns to look for
PARKING_CODE_PATTERNS = [
    r'(CMO\d+)',           # CMO111, CMO142, etc.
    r'(LUNA)',             # LUNA
    r'([A-Z]{2,5}\d{2,4})', # Generic: 2-5 letters + 2-4 numbers
]

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    """
    Extract parking code from template filename.
    Examples: "CMO142_template.xlsx" -> "CMO142"
              "CMO142 (LUNA)_budget.xlsx" -> "CMO142"
    """
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
    """Find a sheet in the workbook that matches any of the given patterns."""
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            pattern_clean = pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
            if pattern_clean in sheet_lower:
                return sheet_name
    return None


def find_parking_data(df, parking_code):
    """
    Find the row containing data for a specific parking code.
    Searches ALL columns and rows for partial matches.
    """
    code_clean = parking_code.upper().strip()
    
    # Search all columns
    for col_idx in range(len(df.columns)):
        for row_idx in range(len(df)):
            value = df.iloc[row_idx, col_idx]
            if pd.notna(value):
                cell_str = str(value).upper().strip()
                if code_clean in cell_str:
                    return row_idx
                if cell_str in code_clean and len(cell_str) >= len(code_clean) - 3:
                    return row_idx
    
    # Try fuzzy matching - look for rows that START with similar pattern
    for col_idx in range(min(3, len(df.columns))):
        for row_idx in range(len(df)):
            value = df.iloc[row_idx, col_idx]
            if pd.notna(value):
                cell_str = str(value).strip().upper()
                # Check if first few chars match
                if len(code_clean) >= 4 and len(cell_str) >= 4:
                    if cell_str[:4] == code_clean[:4]:
                        return row_idx
    
    return None


def find_all_parking_codes(df):
    """Find all parking codes in the DataFrame."""
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
    """Safely convert a value to float."""
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


def find_year_column(df, target_year=None, max_search_cols=30):
    """
    Find the column index containing data for a specific year.
    Looks for headers like "2024", "2025", "Total 2024", etc.
    If target_year is None, returns the rightmost year found.
    """
    year_cols = {}
    
    for col_idx in range(min(max_search_cols, len(df.columns))):
        for row_idx in range(min(15, len(df))):
            cell = str(df.iloc[row_idx, col_idx]).strip()
            # Look for year patterns
            matches = re.findall(r'(20\d{2})', cell)
            for year_str in matches:
                year = int(year_str)
                if year not in year_cols or row_idx < 5:  # Prefer header rows
                    year_cols[year] = col_idx
    
    if target_year and target_year in year_cols:
        return year_cols[target_year]
    
    if year_cols:
        # Return the most recent year
        latest_year = max(year_cols.keys())
        return year_cols[latest_year]
    
    return None


def find_month_columns(df, max_search_cols=15):
    """
    Find column indices for each month.
    Returns dict mapping month index (0-11) to column index.
    """
    month_cols = {}
    
    for col_idx in range(1, min(max_search_cols, len(df.columns))):
        for row_idx in range(min(10, len(df))):
            cell = str(df.iloc[row_idx, col_idx]).strip().lower()
            
            # Check full month names
            for i, month in enumerate(MONTHS):
                if month.lower() in cell:
                    month_cols[i] = col_idx
            
            # Check abbreviations
            for i, month in enumerate(MONTHS_ABBR):
                if month.lower() == cell[:3] and len(cell) <= 4:
                    month_cols[i] = col_idx
            
            # Check numeric months (1, 2, 3...)
            for i in range(1, 13):
                if cell == str(i) or cell == f"{i:02d}":
                    month_cols[i-1] = col_idx
    
    # If still empty, assume columns 1-12 are Jan-Dec
    if not month_cols:
        for i in range(12):
            if i + 1 < len(df.columns):
                month_cols[i] = i + 1
    
    return month_cols


def read_pnl_file(pnl_file):
    """
    Read P&L file and find the sheet with parking data.
    Returns (DataFrame, sheet_name) or (None, None).
    """
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    file_bytes = pnl_file.read() if hasattr(pnl_file, 'read') else None
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    if file_bytes is None:
        return None, None
    
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        
        best_df = None
        best_sheet = None
        most_codes = 0
        
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                codes = find_all_parking_codes(df)
                if len(codes) > most_codes:
                    most_codes = len(codes)
                    best_df = df
                    best_sheet = sheet_name
            except:
                pass
        
        if best_df is not None:
            print(f"Using sheet '{best_sheet}' with {most_codes} parking codes: {find_all_parking_codes(best_df)}")
            return best_df, best_sheet
        
        # Fallback: try first sheet
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        return df, xl.sheet_names[0] if xl.sheet_names else "Sheet1"
    
    except Exception as e:
        print(f"Error reading P&L file: {e}")
        return None, None


# ============================================================================
# DATA EXTRACTION FUNCTIONS
# ============================================================================

def extract_budget_initial_data(pnl_file, parking_code):
    """
    Extract previous year total for Budget Initial (cell S8).
    Gets the TOTAL value from the most recent complete year.
    """
    df, sheet_name = read_pnl_file(pnl_file)
    if df is None:
        return None
    
    parking_row = find_parking_data(df, parking_code)
    if parking_row is None:
        print(f"Parking code '{parking_code}' not found in P&L file")
        return None
    
    print(f"Found {parking_code} at row {parking_row} in sheet '{sheet_name}'")
    
    # Get the total from the rightmost year column
    year_col = find_year_column(df)
    
    if year_col is None:
        # Use the last numeric column
        for col_idx in range(len(df.columns) - 1, 0, -1):
            if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                year_col = col_idx
                break
    
    if year_col is None:
        return None
    
    # Get value from parking row at year column
    total = safe_float(df.iloc[parking_row, year_col])
    
    # If that's 0, try nearby rows for "Total" label
    if total == 0:
        for row_idx in range(parking_row, min(parking_row + 100, len(df))):
            if pd.notna(df.iloc[row_idx, 0]):
                label = str(df.iloc[row_idx, 0]).strip().lower()
                if any(word in label for word in ['total', 'grand total', 'sum', 'total revenue']):
                    val = safe_float(df.iloc[row_idx, year_col])
                    if val > 0:
                        total = val
                        break
    
    return total


def extract_fiche_stationnement_data(pnl_file, parking_code):
    """
    Extract revenue data for Fiche Stationnement (K17-K26).
    Gets all revenue categories and their values from the most recent year.
    """
    df, sheet_name = read_pnl_file(pnl_file)
    if df is None:
        return {}
    
    parking_row = find_parking_data(df, parking_code)
    if parking_row is None:
        return {}
    
    print(f"Extracting revenue data for {parking_code} from row {parking_row}")
    
    # Find the revenue/year column
    revenue_col = find_year_column(df)
    
    if revenue_col is None:
        for col_idx in range(len(df.columns) - 1, 0, -1):
            if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                revenue_col = col_idx
                break
    
    if revenue_col is None:
        return {}
    
    print(f"Using revenue column index: {revenue_col}")
    
    revenue_data = {}
    
    # Extract all rows belonging to this parking
    for row_idx in range(parking_row, min(parking_row + 200, len(df))):
        category = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        value = safe_float(df.iloc[row_idx, revenue_col])
        
        # Stop if we hit another parking code or completely empty row
        if row_idx > parking_row:
            if re.match(r'^[A-Z]{2,5}\d{2,6}$', category.upper()):
                break
            if not category and value == 0:
                # Check if next few rows are also empty
                all_empty = True
                for check_row in range(row_idx, min(row_idx + 3, len(df))):
                    check_cat = str(df.iloc[check_row, 0]).strip() if pd.notna(df.iloc[check_row, 0]) else ""
                    check_val = safe_float(df.iloc[check_row, revenue_col])
                    if check_cat or check_val > 0:
                        all_empty = False
                        break
                if all_empty:
                    break
        
        # Skip headers, totals, and empty
        if not category or category.lower() in ['total', 'grand total', 'sum', 'parking code', 'code', '']:
            continue
        
        if value > 0:
            revenue_data[category] = value
    
    print(f"Found {len(revenue_data)} revenue categories")
    return revenue_data


def extract_donnees_historiques_data(pnl_file, parking_code):
    """
    Extract monthly data for Donnees Historiques.
    Gets Jan-Dec values for the parking code.
    """
    df, sheet_name = read_pnl_file(pnl_file)
    if df is None:
        return {}
    
    parking_row = find_parking_data(df, parking_code)
    if parking_row is None:
        return {}
    
    monthly_data = {}
    month_cols = find_month_columns(df)
    
    if month_cols:
        for month_idx, col_idx in month_cols.items():
            if parking_row < len(df) and col_idx < len(df.columns):
                value = safe_float(df.iloc[parking_row, col_idx])
                monthly_data[MONTHS[month_idx]] = value
    else:
        # Fallback: columns 1-12
        for i in range(12):
            if parking_row < len(df) and i + 1 < len(df.columns):
                value = safe_float(df.iloc[parking_row, i + 1])
                monthly_data[MONTHS[i]] = value
    
    # Filter out months with zero values
    monthly_data = {k: v for k, v in monthly_data.items() if v > 0}
    
    print(f"Found {len(monthly_data)} months of data")
    return monthly_data


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, pnl_file, parking_code):
    """Update Budget Initial sheet - cell S8."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            return ["Budget Initial: Sheet not found"]
        
        ws = wb[sheet_name]
        total = extract_budget_initial_data(pnl_file, parking_code)
        
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
    """Update Fiche Stationnement - K17-K26."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            return ["Fiche Stationnement: Sheet not found"]
        
        ws = wb[sheet_name]
        revenue_data = extract_fiche_stationnement_data(pnl_file, parking_code)
        
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
        
        # Word data if provided
        if word_data:
            for row in range(43, 56):
                if "Nb abonnés" in word_data:
                    ws[f"H{row}"] = word_data["Nb abonnés"]
                if "Informations" in word_data:
                    ws[f"I{row}"] = word_data["Informations"]
                if "Avant taxes" in word_data:
                    ws[f"J{row}"] = word_data["Avant taxes"]
                    ws[f"L{row}"] = word_data["Avant taxes"]
            updates.append("✅ Updated additional data from Word")
    
    except Exception as e:
        updates.append(f"❌ Fiche Stationnement: {e}")
    
    return updates


def update_donnees_historiques(wb, pnl_file, parking_code):
    """Update Donnees Historiques - yellow cells."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            return ["Donnees Historiques: Sheet not found"]
        
        ws = wb[sheet_name]
        monthly_data = extract_donnees_historiques_data(pnl_file, parking_code)
        
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
            updates.append(f"✅ Updated {cells_updated} cells with monthly data")
        else:
            updates.append("⚠️ No monthly data to update")
    
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
        return None, ["❌ Could not determine parking code from filename. Rename your template to include the code (e.g., CMO142_template.xlsx)"]
    
    updates.append(f"🔍 Processing: {parking_code}")
    
    # Read the Excel template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"❌ Error reading template: {e}"]
    
    updates.append(f"📋 Sheets found: {len(wb.sheetnames)}")
    
    # Process each sheet
    updates.extend(update_budget_initial(wb, pnl_file, parking_code))
    updates.extend(update_fiche_stationnement(wb, pnl_file, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, pnl_file, parking_code))
    
    # Count successful updates
    success_count = sum(1 for u in updates if u.startswith("✅"))
    
    if success_count == 0:
        # Show available codes to help debug
        df, _ = read_pnl_file(pnl_file)
        if df is not None:
            codes = find_all_parking_codes(df)
            updates.append(f"💡 Available codes in P&L: {', '.join(codes[:10])}")
    
    # Reset P&L file pointer
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
