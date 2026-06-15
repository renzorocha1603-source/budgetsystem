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

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    """Extract parking code from template filename."""
    if not filename:
        return None
    name = filename.rsplit('.', 1)[0]
    # Try CMO pattern first
    match = re.search(r'(CMO\d+)', name, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    # Try LUNA
    if 'LUNA' in name.upper():
        return 'LUNA'
    # Fallback
    parts = name.replace('(', ' ').replace(')', ' ').split('_')[0].split()[0]
    return parts.upper()


def find_sheet_by_pattern(wb, patterns):
    """Find a sheet in the workbook matching patterns."""
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            pattern_clean = pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
            if pattern_clean in sheet_lower:
                return sheet_name
    return None


def find_pnl_sheet_for_code(pnl_file, parking_code):
    """
    Find the P&L sheet for a specific parking code.
    Each parking code has its own sheet named after the code (e.g., "CMO142").
    """
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    file_bytes = pnl_file.read() if hasattr(pnl_file, 'read') else None
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    if file_bytes is None:
        return None, None
    
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        
        # First try exact match
        for sheet_name in xl.sheet_names:
            if sheet_name.upper().strip() == parking_code.upper().strip():
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                print(f"Found exact sheet match: '{sheet_name}'")
                return df, sheet_name
        
        # Try partial match (parking code in sheet name)
        for sheet_name in xl.sheet_names:
            if parking_code.upper() in sheet_name.upper():
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                print(f"Found partial sheet match: '{sheet_name}'")
                return df, sheet_name
        
        # Try searching ALL sheets for the parking code in cells
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
                for col_idx in range(min(5, len(df.columns))):
                    for row_idx in range(min(20, len(df))):
                        value = df.iloc[row_idx, col_idx]
                        if pd.notna(value) and parking_code.upper() in str(value).upper():
                            print(f"Found code in sheet '{sheet_name}' at row {row_idx}")
                            return df, sheet_name
            except:
                pass
        
        print(f"Could not find sheet for {parking_code}")
        return None, None
    
    except Exception as e:
        print(f"Error reading P&L: {e}")
        return None, None


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


def is_parking_code(value):
    """Check if a value looks like a parking code (e.g., CMO142)."""
    if pd.isna(value) or value is None:
        return False
    return bool(re.match(r'^[A-Z]{2,5}\d{2,6}$', str(value).strip().upper()))


# ============================================================================
# DATA EXTRACTION FUNCTIONS
# ============================================================================

def extract_budget_initial_total(pnl_file, parking_code):
    """
    Extract total for Budget Initial.
    Looks in the parking code's sheet for a total value in the last numeric column.
    """
    df, sheet_name = find_pnl_sheet_for_code(pnl_file, parking_code)
    if df is None:
        return None
    
    print(f"Budget Initial: Using sheet '{sheet_name}'")
    
    # Find the row with the parking code label
    code_row = None
    for col_idx in range(min(5, len(df.columns))):
        for row_idx in range(min(20, len(df))):
            value = df.iloc[row_idx, col_idx]
            if pd.notna(value) and parking_code.upper() in str(value).upper():
                code_row = row_idx
                break
        if code_row is not None:
            break
    
    if code_row is None:
        # Try to find any row with data
        for row_idx in range(len(df)):
            for col_idx in range(min(5, len(df.columns))):
                if pd.notna(df.iloc[row_idx, col_idx]):
                    code_row = row_idx
                    break
            if code_row is not None:
                break
    
    if code_row is None:
        return None
    
    # Look for "Total" or "Grand Total" in Column A
    total_value = None
    for row_idx in range(len(df)):
        if pd.notna(df.iloc[row_idx, 0]):
            label = str(df.iloc[row_idx, 0]).strip().lower()
            if 'total' in label or 'grand total' in label:
                # Get value from the last numeric column
                for col_idx in range(len(df.columns) - 1, 0, -1):
                    val = safe_float(df.iloc[row_idx, col_idx])
                    if val > 0:
                        total_value = val
                        print(f"Found total at row {row_idx}: ${total_value:,.2f}")
                        return total_value
    
    # Fallback: sum all values in Column B (since revenue data is in A-B columns)
    if total_value is None:
        total = 0
        for row_idx in range(len(df)):
            val = safe_float(df.iloc[row_idx, 1])  # Column B
            if val > 0:
                total += val
        if total > 0:
            print(f"Summed all values: ${total:,.2f}")
            return total
    
    return None


def extract_revenue_categories(pnl_file, parking_code):
    """
    Extract revenue categories from the P&L sheet.
    Data is in Column A (category name) and Column B (value).
    """
    df, sheet_name = find_pnl_sheet_for_code(pnl_file, parking_code)
    if df is None:
        return {}
    
    print(f"Revenue: Using sheet '{sheet_name}'")
    print(f"Revenue: Sheet has {len(df)} rows, {len(df.columns)} columns")
    
    revenue_data = {}
    
    # Skip header rows and find where actual data starts
    data_start = 0
    for row_idx in range(len(df)):
        col_a = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        col_b = safe_float(df.iloc[row_idx, 1])
        
        # Skip rows that are clearly headers
        if col_a.lower() in ['code', 'profit & loss', 'profit and loss', '']:
            continue
        
        # Skip rows with dates or filters
        if re.search(r'\d{2}-\d{2}-\d{2}', col_a):
            continue
        
        if col_b > 0 and not is_parking_code(col_a):
            data_start = row_idx
            break
    
    print(f"Revenue: Data starts at row {data_start}")
    
    # Extract all revenue categories
    for row_idx in range(data_start, len(df)):
        category = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        value = safe_float(df.iloc[row_idx, 1])
        
        # Stop conditions
        if category.lower() in ['total', 'grand total']:
            break
        
        if not category and value == 0:
            # Check if next rows are also empty
            all_empty = True
            for check in range(row_idx, min(row_idx + 3, len(df))):
                c = str(df.iloc[check, 0]).strip() if pd.notna(df.iloc[check, 0]) else ""
                v = safe_float(df.iloc[check, 1])
                if c or v > 0:
                    all_empty = False
                    break
            if all_empty:
                break
            continue
        
        if category and value > 0 and not is_parking_code(category):
            # Skip filter/header rows
            if not re.search(r'\d{2}-\d{2}-\d{2}', category):
                revenue_data[category] = value
                print(f"  {category}: ${value:,.2f}")
    
    print(f"Revenue: Found {len(revenue_data)} categories")
    return revenue_data


def extract_monthly_data(pnl_file, parking_code):
    """
    Extract monthly data from the P&L sheet.
    Looks for rows with month names or month abbreviations in Column A,
    with values in Column B.
    """
    df, sheet_name = find_pnl_sheet_for_code(pnl_file, parking_code)
    if df is None:
        return {}
    
    print(f"Monthly: Using sheet '{sheet_name}'")
    
    monthly_data = {}
    
    # Look for rows that contain month names
    for row_idx in range(len(df)):
        category = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        value = safe_float(df.iloc[row_idx, 1])
        
        if not category or value <= 0:
            continue
        
        # Check if this row contains a month reference
        cat_lower = category.lower()
        for i, month in enumerate(MONTHS):
            if month.lower() in cat_lower:
                monthly_data[MONTHS[i]] = value
                print(f"  {MONTHS[i]}: ${value:,.2f}")
                break
        
        for i, month in enumerate(MONTHS_ABBR):
            if month.lower() in cat_lower and len(cat_lower) <= 10:
                monthly_data[MONTHS[i]] = value
                print(f"  {MONTHS[i]}: ${value:,.2f}")
                break
    
    # If no month data found, try to split "Monthly" rows into 12 equal parts
    if not monthly_data:
        monthly_rows = []
        for row_idx in range(len(df)):
            category = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
            value = safe_float(df.iloc[row_idx, 1])
            if 'monthly' in category.lower() and value > 0:
                monthly_rows.append((category, value))
        
        if monthly_rows:
            # Take the first monthly total and divide by 12
            for cat, val in monthly_rows[:1]:
                monthly_value = val / 12
                for month in MONTHS:
                    monthly_data[month] = monthly_value
                print(f"  Divided '{cat}' (${val:,.2f}) into 12 months of ${monthly_value:,.2f}")
                break
    
    print(f"Monthly: Found {len(monthly_data)} months")
    return monthly_data


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
    
    # Read the Excel template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"❌ Error reading template: {e}"]
    
    updates.append(f"📋 Template has {len(wb.sheetnames)} sheets")
    
    # Process each sheet
    updates.extend(update_budget_initial(wb, pnl_file, parking_code))
    updates.extend(update_fiche_stationnement(wb, pnl_file, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, pnl_file, parking_code))
    
    # Count successes
    success_count = sum(1 for u in updates if u.startswith("✅"))
    if success_count == 0:
        updates.append("💡 No updates were made. Check if the parking code exists in the P&L file.")
    
    # Reset P&L file pointer
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
