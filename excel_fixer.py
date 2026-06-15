import io
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

# Parking codes that can appear in P&L files - this is dynamic
# The code will match ANY parking code it finds in the files

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

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_parking_code_from_filename(filename):
    """
    Extract parking code from template filename.
    Examples: "CMO142_template.xlsx" -> "CMO142"
              "LUNA_2027_budget.xlsx" -> "LUNA"
    """
    if not filename:
        return None
    
    # Remove extension
    name = filename.rsplit('.', 1)[0]
    
    # Try to find known patterns (CMO followed by numbers, or common codes)
    patterns = [
        r'(CMO\d+)',           # CMO111, CMO142, etc.
        r'(LUNA)',             # LUNA
        r'([A-Z]{2,4}\d{2,4})', # Generic: 2-4 letters + 2-4 numbers
    ]
    
    for pattern in patterns:
        match = re.search(pattern, name, re.IGNORECASE)
        if match:
            return match.group(1).upper()
    
    # Fallback: use first word before underscore
    parts = name.split('_')
    if parts:
        return parts[0].upper()
    
    return name.upper()


def find_parking_data(df, parking_code, column_search_range=5):
    """
    Find the row containing data for a specific parking code.
    Searches first few columns for the code.
    Returns the row index or None.
    """
    for col_idx in range(min(column_search_range, len(df.columns))):
        for row_idx, value in enumerate(df.iloc[:, col_idx]):
            if pd.notna(value) and parking_code.upper() in str(value).upper():
                return row_idx
    return None


def find_column_with_header(df, possible_headers, max_col=30):
    """
    Find column index that matches any of the possible headers.
    Returns column index or None.
    """
    for col_idx in range(min(max_col, len(df.columns))):
        for row_idx in range(min(5, len(df))):
            cell_value = str(df.iloc[row_idx, col_idx]).strip().lower()
            for header in possible_headers:
                if header.lower() in cell_value:
                    return col_idx
    return None


def safe_float(value, default=0.0):
    """Safely convert a value to float."""
    try:
        if pd.isna(value) or value is None:
            return default
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '')
        return float(value)
    except (ValueError, TypeError):
        return default


# ============================================================================
# DATA EXTRACTION FUNCTIONS
# ============================================================================

def extract_budget_initial_data(pnl_file, parking_code):
    """
    Extract previous year total for Budget Initial (cell S8).
    Looks for the parking code and gets the TOTAL/Sum from the previous year column.
    """
    try:
        df = pd.read_excel(io.BytesIO(pnl_file.read()) if hasattr(pnl_file, 'read') else pnl_file, engine='openpyxl')
        pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
        
        # Find the row for this parking code
        parking_row = find_parking_data(df, parking_code)
        if parking_row is None:
            return None
        
        # Look for "Total" or "Grand Total" row near this parking's section
        # Usually totals are at the bottom of each parking section
        total_value = None
        
        # Search for total in last few columns (N, O, P etc for yearly totals)
        for col_idx in range(len(df.columns) - 1, max(len(df.columns) - 10, 0), -1):
            for row_idx in range(parking_row, min(parking_row + 50, len(df))):
                cell_value = str(df.iloc[row_idx, 0]).strip().lower() if pd.notna(df.iloc[row_idx, 0]) else ""
                if 'total' in cell_value or 'grand total' in cell_value or 'sum' in cell_value:
                    val = safe_float(df.iloc[row_idx, col_idx])
                    if val > 0:
                        total_value = val
                        break
            if total_value:
                break
        
        # If no explicit total found, use the last numeric value in the parking's data
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
    """
    Extract revenue data for 1. Fiche Stationnement (K17-K26).
    Looks for 2024 data (or previous year) for each revenue category.
    Returns dict with categories and their values.
    """
    try:
        df = pd.read_excel(io.BytesIO(pnl_file.read()) if hasattr(pnl_file, 'read') else pnl_file, engine='openpyxl')
        pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
        
        parking_row = find_parking_data(df, parking_code)
        if parking_row is None:
            return {}
        
        revenue_data = {}
        
        # Find the column with revenue data (look for year headers like "2024", "2025", etc.)
        # The revenue data is usually in columns with year headers
        revenue_col = None
        for col_idx in range(1, len(df.columns)):
            for row_idx in range(min(5, len(df))):
                cell = str(df.iloc[row_idx, col_idx]).strip()
                if re.search(r'20\d{2}', cell):  # Matches 2024, 2025, etc.
                    revenue_col = col_idx
                    break
            if revenue_col:
                break
        
        if revenue_col is None:
            # Fallback: use the last numeric column
            for col_idx in range(len(df.columns) - 1, 0, -1):
                if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                    revenue_col = col_idx
                    break
        
        # Extract categories and values
        for row_idx in range(parking_row, min(parking_row + 100, len(df))):
            category = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
            value = safe_float(df.iloc[row_idx, revenue_col]) if revenue_col else 0
            
            # Skip header rows, empty rows, and total rows
            if not category or category.lower() in ['total', 'grand total', 'sum', 'parking code']:
                continue
            
            # Skip rows that are just parking codes
            if re.match(r'^[A-Z]{2,4}\d{2,6}$', category.upper()):
                continue
            
            if value > 0:
                revenue_data[category] = value
        
        return revenue_data
    except Exception as e:
        print(f"Error extracting Fiche Stationnement data: {e}")
        return {}


def extract_donnees_historiques_data(pnl_file, parking_code):
    """
    Extract monthly data for 2. Donnees Historiques.
    Gets Jan-Apr of current year (2026) and May-Dec of previous year (2025).
    """
    try:
        df = pd.read_excel(io.BytesIO(pnl_file.read()) if hasattr(pnl_file, 'read') else pnl_file, engine='openpyxl')
        pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
        
        parking_row = find_parking_data(df, parking_code)
        if parking_row is None:
            return {}
        
        monthly_data = {}
        current_year = datetime.now().year  # 2026
        
        # Find monthly columns (B-M usually)
        # Look for month names in the first few rows
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
        
        # If no month headers found, assume columns B-M are Jan-Dec
        if not month_cols:
            for i in range(12):
                month_cols[i] = i + 1  # B=1, C=2, ..., M=12
        
        # Extract data for each month
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
    """Update Budget Initial sheet - cell S8 with previous year total."""
    updates = []
    try:
        if "Budget Initial" not in wb.sheetnames:
            return updates
        
        ws = wb["Budget Initial"]
        total = extract_budget_initial_data(pnl_file, parking_code)
        
        if total and total > 0:
            ws["S8"] = total
            ws["S8"].number_format = '#,##0.00'
            updates.append(f"Budget Initial: Updated S8 with ${total:,.2f}")
    except Exception as e:
        updates.append(f"Budget Initial: Error - {e}")
    
    return updates


def update_fiche_stationnement(wb, pnl_file, parking_code, word_data=None):
    """Update 1. Fiche Stationnement - K17-K26 and H-M rows 43-55."""
    updates = []
    try:
        if "1. Fiche Stationnement" not in wb.sheetnames:
            return updates
        
        ws = wb["1. Fiche Stationnement"]
        revenue_data = extract_fiche_stationnement_data(pnl_file, parking_code)
        
        if not revenue_data:
            return updates
        
        # Sort categories by value (largest first) to prioritize important ones
        sorted_categories = sorted(revenue_data.items(), key=lambda x: x[1], reverse=True)
        
        # Fill K17 to K25 with categories
        total_revenue = 0
        for i, (category, value) in enumerate(sorted_categories):
            if i >= 9:  # K17 to K25 = 9 slots
                break
            cell = f"K{17 + i}"
            if cell in ws or True:  # openpyxl creates cell if it doesn't exist
                ws[cell] = value
                ws[cell].number_format = '#,##0.00'
                total_revenue += value
                updates.append(f"Fiche Stationnement: {cell} = {category}: ${value:,.2f}")
        
        # K26 = Total Revenue
        ws["K26"] = total_revenue
        ws["K26"].number_format = '#,##0.00'
        updates.append(f"Fiche Stationnement: K26 = Total Revenue: ${total_revenue:,.2f}")
        
        # H-M, rows 43-55 from word_data if provided
        if word_data:
            for row in range(43, 56):
                if "Nb abonnés" in word_data:
                    ws[f"H{row}"] = word_data["Nb abonnés"]
                if "Informations" in word_data:
                    ws[f"I{row}"] = word_data["Informations"]
                if "Avant taxes" in word_data:
                    ws[f"J{row}"] = word_data["Avant taxes"]
                    ws[f"L{row}"] = word_data["Avant taxes"]
            updates.append("Fiche Stationnement: Updated H-M rows 43-55 with word data")
    
    except Exception as e:
        updates.append(f"Fiche Stationnement: Error - {e}")
    
    return updates


def update_donnees_historiques(wb, pnl_file, parking_code):
    """Update 2. Donnees Historiques - yellow cells only."""
    updates = []
    try:
        if "2. Donnees Historiques" not in wb.sheetnames:
            return updates
        
        ws = wb["2. Donnees Historiques"]
        monthly_data = extract_donnees_historiques_data(pnl_file, parking_code)
        
        if not monthly_data:
            return updates
        
        # Yellow rows: 36-76, skip 44, 47, 65 (white/formula cells)
        yellow_rows = [r for r in range(36, 77) if r not in [44, 47, 65]]
        # Columns B-M
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
            updates.append(f"Donnees Historiques: Updated {cells_updated} cells with monthly data")
    
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
        return None, ["Error: Could not determine parking code from filename"]
    
    # Read the Excel template
    try:
        wb = load_workbook(io.BytesIO(excel_file.read()))
        excel_file.seek(0) if hasattr(excel_file, 'seek') else None
    except Exception as e:
        return None, [f"Error reading template: {e}"]
    
    # Get available sheets
    available_sheets = wb.sheetnames
    
    # Process each sheet based on what's available in the template
    if "Budget Initial" in available_sheets:
        updates.extend(update_budget_initial(wb, pnl_file, parking_code))
    
    if "1. Fiche Stationnement" in available_sheets:
        updates.extend(update_fiche_stationnement(wb, pnl_file, parking_code, word_data))
    
    if "2. Donnees Historiques" in available_sheets:
        updates.extend(update_donnees_historiques(wb, pnl_file, parking_code))
    
    # Reset P&L file pointer for potential reuse
    pnl_file.seek(0) if hasattr(pnl_file, 'seek') else None
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates
