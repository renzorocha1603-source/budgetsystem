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

# Sheet name patterns to match (order matters - first match wins)
SHEET_PATTERNS = {
    "Budget Initial": ["budget initial", "budget"],
    "Fiche Stationnement": ["fiche stationnement", "fiche de stationnement", "stationnement"],
    "Donnees Historiques": ["donnees historiques", "données historiques", "historiques", "historique"],
    "Revenu Depense": ["revenu depense", "revenu-depense", "revenu dépense", "revenu"],
    "Calcul Salaire": ["calcul salaire", "calcul-salaire", "salaire"],
}

def find_sheet_by_pattern(wb, patterns):
    """Find a sheet in the workbook that matches any of the given patterns."""
    for sheet_name in wb.sheetnames:
        sheet_lower = sheet_name.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
        for pattern in patterns:
            pattern_clean = pattern.lower().replace('.', ' ').replace('-', ' ').replace('_', ' ')
            if pattern_clean in sheet_lower:
                return sheet_name
    return None

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    """
    Extract parking code from template filename.
    Examples: "CMO142_template.xlsx" -> "CMO142"
              "CMO142 (LUNA)_budget.xlsx" -> "CMO142"
              "LUNA_2027_budget.xlsx" -> "LUNA"
    """
    if not filename:
        return None
    
    # Remove extension
    name = filename.rsplit('.', 1)[0]
    
    # Try to find CMO pattern first (most common)
    cmo_match = re.search(r'(CMO\d+)', name, re.IGNORECASE)
    if cmo_match:
        return cmo_match.group(1).upper()
    
    # Try LUNA
    if 'LUNA' in name.upper():
        return 'LUNA'
    
    # Try generic pattern: 2-4 letters + 3-6 numbers
    generic_match = re.search(r'([A-Z]{2,4}\d{3,6})', name, re.IGNORECASE)
    if generic_match:
        return generic_match.group(1).upper()
    
    # Fallback: use first word before underscore or space
    parts = name.replace('(', ' ').replace(')', ' ').split('_')[0].split()[0]
    return parts.upper()


def find_parking_data(df, parking_code, column_search_range=5):
    """
    Find the row containing data for a specific parking code.
    Searches first few columns for the code.
    Now does partial matching (e.g., "CMO142" matches "CMO142 (LUNA)")
    """
    code_clean = parking_code.upper().strip()
    
    for col_idx in range(min(column_search_range, len(df.columns))):
        for row_idx, value in enumerate(df.iloc[:, col_idx]):
            if pd.notna(value):
                cell_str = str(value).upper().strip()
                if code_clean in cell_str:
                    return row_idx
                if cell_str in code_clean:
                    return row_idx
    return None


def find_all_parking_codes(df, column_search_range=3):
    """Find all parking codes in the P&L file."""
    codes = []
    for col_idx in range(min(column_search_range, len(df.columns))):
        for row_idx, value in enumerate(df.iloc[:, col_idx]):
            if pd.notna(value):
                cell_str = str(value).strip()
                match = re.search(r'(CMO\d+)', cell_str, re.IGNORECASE)
                if match:
                    codes.append(match.group(1).upper())
    return list(set(codes))


def safe_float(value, default=0.0):
    """Safely convert a value to float."""
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '').replace('€', '')
        return float(value)
    except (ValueError, TypeError):
        return default


# ============================================================================
# DATA EXTRACTION FUNCTIONS
# ============================================================================

def extract_budget_initial_data(pnl_file, parking_code):
    """Extract previous year total for Budget Initial (cell S8)."""
    try:
        df = pd.read_excel(io.BytesIO(pnl_file.read()) if hasattr(pnl_file, 'read') else pnl_file, engine='openpyxl')
        pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
        
        available_codes = find_all_parking_codes(df)
        print(f"Available parking codes in P&L: {available_codes}")
        print(f"Looking for: {parking_code}")
        
        parking_row = find_parking_data(df, parking_code)
        if parking_row is None:
            print(f"Parking code '{parking_code}' not found in P&L file")
            return None
        
        print(f"Found parking code at row {parking_row}")
        
        total_value = None
        
        # Search for total in last few columns
        for col_idx in range(len(df.columns) - 1, max(len(df.columns) - 10, 0), -1):
            for row_idx in range(parking_row, min(parking_row + 50, len(df))):
                if pd.notna(df.iloc[row_idx, 0]):
                    cell_value = str(df.iloc[row_idx, 0]).strip().lower()
                    if 'total' in cell_value or 'grand total' in cell_value or 'sum' in cell_value:
                        val = safe_float(df.iloc[row_idx, col_idx])
                        if val > 0:
                            total_value = val
                            break
            if total_value:
                break
        
        if total_value is None:
            for col_idx in range(len(df.columns) - 1, 0, -1):
                val = safe_float(df.iloc[parking_row, col_idx])
                if val > 0:
                    total_value = val
                    break
        
        return total_value
    except Exception as e:
        print(f"Error extracting Budget Initial data: {e}")
        return None


def extract_fiche_stationnement_data(pnl_file, parking_code):
    """Extract revenue data for Fiche Stationnement (K17-K26)."""
    try:
        df = pd.read_excel(io.BytesIO(pnl_file.read()) if hasattr(pnl_file, 'read') else pnl_file, engine='openpyxl')
        pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
        
        parking_row = find_parking_data(df, parking_code)
        if parking_row is None:
            return {}
        
        revenue_data = {}
        
        # Find revenue column (look for year headers)
        revenue_col = None
        for col_idx in range(1, len(df.columns)):
            for row_idx in range(min(5, len(df))):
                cell = str(df.iloc[row_idx, col_idx]).strip()
                if re.search(r'20\d{2}', cell):
                    revenue_col = col_idx
                    break
            if revenue_col:
                break
        
        if revenue_col is None:
            for col_idx in range(len(df.columns) - 1, 0, -1):
                if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                    revenue_col = col_idx
                    break
        
        if revenue_col is None:
            return {}
        
        # Extract categories
        for row_idx in range(parking_row, min(parking_row + 100, len(df))):
            category = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
            value = safe_float(df.iloc[row_idx, revenue_col])
            
            if not category or category.lower() in ['total', 'grand total', 'sum', 'parking code', '']:
                continue
            
            if re.match(r'^[A-Z]{2,4}\d{2,6}$', category.upper()):
                continue
            
            if value > 0:
                revenue_data[category] = value
        
        return revenue_data
    except Exception as e:
        print(f"Error extracting Fiche Stationnement data: {e}")
        return {}


def extract_donnees_historiques_data(pnl_file, parking_code):
    """Extract monthly data for Donnees Historiques."""
    try:
        df = pd.read_excel(io.BytesIO(pnl_file.read()) if hasattr(pnl_file, 'read') else pnl_file, engine='openpyxl')
        pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
        
        parking_row = find_parking_data(df, parking_code)
        if parking_row is None:
            return {}
        
        monthly_data = {}
        
        month_cols = {}
        for col_idx in range(1, min(14, len(df.columns))):
            for row_idx in range(min(5, len(df))):
                cell = str(df.iloc[row_idx, col_idx]).strip().lower()
                for i, month in enumerate(MONTHS):
                    if month.lower() in cell:
                        month_cols[i] = col_idx
                        break
                for i, month in enumerate(MONTHS_ABBR):
                    if month.lower() == cell[:3]:
                        month_cols[i] = col_idx
                        break
        
        if not month_cols:
            for i in range(12):
                month_cols[i] = i + 1
        
        for month_idx, col_idx in month_cols.items():
            value = safe_float(df.iloc[parking_row, col_idx]) if parking_row < len(df) else 0
            monthly_data[MONTHS[month_idx]] = value
        
        return monthly_data
    except Exception as e:
        print(f"Error extracting Donnees Historiques data: {e}")
        return {}


# ============================================================================
# SHEET UPDATE FUNCTIONS
# ============================================================================

def update_budget_initial(wb, pnl_file, parking_code):
    """Update Budget Initial sheet - cell S8."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Budget Initial"])
        if not sheet_name:
            updates.append("Budget Initial: Sheet not found in template")
            return updates
        
        ws = wb[sheet_name]
        total = extract_budget_initial_data(pnl_file, parking_code)
        
        if total and total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"Budget Initial ({sheet_name}): Updated S8 with ${total:,.2f}")
        else:
            updates.append(f"Budget Initial ({sheet_name}): No total found for {parking_code}")
    except Exception as e:
        updates.append(f"Budget Initial: Error - {e}")
    
    return updates


def update_fiche_stationnement(wb, pnl_file, parking_code, word_data=None):
    """Update Fiche Stationnement - K17-K26."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Fiche Stationnement"])
        if not sheet_name:
            updates.append("Fiche Stationnement: Sheet not found in template")
            return updates
        
        ws = wb[sheet_name]
        revenue_data = extract_fiche_stationnement_data(pnl_file, parking_code)
        
        if not revenue_data:
            updates.append(f"Fiche Stationnement ({sheet_name}): No revenue data found for {parking_code}")
            return updates
        
        sorted_categories = sorted(revenue_data.items(), key=lambda x: x[1], reverse=True)
        
        total_revenue = 0
        for i, (category, value) in enumerate(sorted_categories):
            if i >= 9:
                break
            cell = f"K{17 + i}"
            ws[cell] = value
            ws[cell].number_format = '#,##0.00'
            total_revenue += value
            updates.append(f"Fiche Stationnement ({sheet_name}): {cell} = {category}: ${value:,.2f}")
        
        ws["K26"] = total_revenue
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"Fiche Stationnement ({sheet_name}): K26 = Total Revenue: ${total_revenue:,.2f}")
    
    except Exception as e:
        updates.append(f"Fiche Stationnement: Error - {e}")
    
    return updates


def update_donnees_historiques(wb, pnl_file, parking_code):
    """Update Donnees Historiques - yellow cells."""
    updates = []
    try:
        sheet_name = find_sheet_by_pattern(wb, SHEET_PATTERNS["Donnees Historiques"])
        if not sheet_name:
            updates.append("Donnees Historiques: Sheet not found in template")
            return updates
        
        ws = wb[sheet_name]
        monthly_data = extract_donnees_historiques_data(pnl_file, parking_code)
        
        if not monthly_data:
            updates.append(f"Donnees Historiques ({sheet_name}): No monthly data found for {parking_code}")
            return updates
        
        yellow_rows = [r for r in range(36, 77) if r not in [44, 47, 65]]
        columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        
        cells_updated = 0
        for row in yellow_rows:
            for col_idx, col_letter in enumerate(columns):
                if col_idx < len(MONTHS):
                    month = MONTHS[col_idx]
                    if month in monthly_data and monthly_data[month] > 0:
                        cell = f"{col_letter}{row}"
                        ws[cell] = monthly_data[month]
                        ws[cell].number_format = '#,##0.00'
                        cells_updated += 1
        
        if cells_updated > 0:
            updates.append(f"Donnees Historiques ({sheet_name}): Updated {cells_updated} cells")
        else:
            updates.append(f"Donnees Historiques ({sheet_name}): No monthly data to update")
    
    except Exception as e:
        updates.append(f"Donnees Historiques: Error - {e}")
    
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
        return None, ["Error: Could not determine parking code from filename. Please rename your template to include the parking code (e.g., CMO142_template.xlsx)"]
    
    updates.append(f"🔍 Processing parking code: {parking_code}")
    
    # Read the Excel template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"Error reading template: {e}"]
    
    updates.append(f"📋 Template sheets: {', '.join(wb.sheetnames)}")
    
    # Process each sheet
    updates.extend(update_budget_initial(wb, pnl_file, parking_code))
    updates.extend(update_fiche_stationnement(wb, pnl_file, parking_code, word_data))
    updates.extend(update_donnees_historiques(wb, pnl_file, parking_code))
    
    # Reset P&L file pointer
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
