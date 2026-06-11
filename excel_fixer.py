import io
import re
import pandas as pd
import pdfplumber
from openpyxl import load_workbook

# --- CONFIGURATION ---

# Seasonal multipliers for different parking types
SEASONAL_MULTIPLIERS = {
    "SC": [1.2, 1.2, 1.2, 1.1, 0.8, 0.8, 0.8, 0.8, 1.1, 1.1, 1.2, 1.2],  # School
    "RG": [0.8, 0.8, 0.9, 0.9, 1.3, 1.3, 1.3, 1.2, 1.0, 1.0, 0.8, 0.8]   # Regular/Tourism
}

# Cell mappings for each sheet
CELL_MAPPINGS = {
    "Budget Initial": {
        "previous_year_total": ["S8"]  # Column N in source, S8 in template
    },
    "1. Fiche Stationnement": {
        "revenue_categories": {
            # These are the default categories. The code will dynamically match any category found in the P&L file.
            "Transient Revenue": ["K17"],
            "Monthly Revenue": ["K18"],
            "VIP": ["K19"],
            "Reserved - Reguliers": ["K20"],
            # Additional categories will be added dynamically if they exist in the P&L file.
        },
        "total_revenue": ["K26"],
        "additional_data": {
            "Nb abonnés": "H",
            "Informations": "I",
            "Avant taxes": ["J", "L"],  # Two columns for prices
        },
        "additional_data_rows": list(range(43, 56))  # Rows 43 to 55
    },
    "2. Donnees Historiques": {
        "monthly_data": {
            "rows": list(range(36, 77)),  # Rows 36 to 76
            "skip_rows": [44, 47, 65],    # White/formula cells
            "columns": list("BCDEFGHIJKLM")  # Columns B to M
        }
    }
}

HOURLY_RATE = 25.0  # Default hourly rate for supervisors


# --- CORE FUNCTIONS ---

def extract_previous_year_total(pnl_file, parking_code, year=2025):
    """
    Extract the previous year total for a specific parking code from the P&L file.
    Assumes the total is in column N (index 13) and parking codes are in column A (index 0).
    Works for ANY parking code (e.g., CMO111, CMO142, etc.).
    """
    try:
        df = pd.read_excel(pnl_file, engine='openpyxl')
        # Filter by parking code (case-insensitive)
        parking_data = df[df.iloc[:, 0].astype(str).str.contains(parking_code, case=False, na=False)]
        if not parking_data.empty:
            return float(parking_data.iloc[0, 13])  # Column N (14th column, 0-based index 13)
    except Exception as e:
        print(f"Error extracting previous year total for {parking_code}: {e}")
    return 0.0


def extract_2024_revenue_data(pnl_file, parking_code):
    """
    Extract 2024 revenue data for a specific parking code from the P&L file.
    Dynamically matches ANY revenue category found in the P&L file.
    Returns a dictionary with revenue categories and their values.
    """
    revenue_data = {}
    try:
        df = pd.read_excel(pnl_file, engine='openpyxl')
        # Filter by parking code (case-insensitive)
        parking_data = df[df.iloc[:, 0].astype(str).str.contains(parking_code, case=False, na=False)]
        
        if not parking_data.empty:
            # Assuming revenue categories are in the first column and 2024 data is in column M (index 12)
            for _, row in parking_data.iterrows():
                category = str(row.iloc[0]).strip()
                if category and pd.notna(row.iloc[12]):  # Column M (13th column, 0-based index 12)
                    revenue_data[category] = float(row.iloc[12])
    except Exception as e:
        print(f"Error extracting 2024 revenue data for {parking_code}: {e}")
    return revenue_data


def extract_2026_and_2025_data(pnl_file, parking_code):
    """
    Extract Jan-Apr 2026 and May-Dec 2025 data for a specific parking code.
    Works for ANY parking code (e.g., CMO111, CMO142, etc.).
    Returns a dictionary with monthly data.
    """
    monthly_data = {}
    try:
        df = pd.read_excel(pnl_file, engine='openpyxl')
        # Filter by parking code (case-insensitive)
        parking_data = df[df.iloc[:, 0].astype(str).str.contains(parking_code, case=False, na=False)]
        
        if not parking_data.empty:
            # Assuming months are in columns B to M (index 1 to 12)
            # Jan-Apr 2026: columns B to E (index 1 to 4)
            # May-Dec 2025: columns F to M (index 5 to 12)
            months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            for i, month in enumerate(months):
                if i < 4:  # Jan-Apr 2026
                    monthly_data[month] = float(parking_data.iloc[0, i + 1])  # Columns B to E
                else:  # May-Dec 2025
                    monthly_data[month] = float(parking_data.iloc[0, i + 1])  # Columns F to M
    except Exception as e:
        print(f"Error extracting monthly data for {parking_code}: {e}")
    return monthly_data


def update_budget_initial(wb, pnl_file, parking_code):
    """
    Update the Budget Initial sheet with the previous year total for ANY parking code.
    """
    updates_log = []
    try:
        if "Budget Initial" in wb.sheetnames:
            ws = wb["Budget Initial"]
            previous_year_total = extract_previous_year_total(pnl_file, parking_code, year=2025)
            
            for cell in CELL_MAPPINGS["Budget Initial"]["previous_year_total"]:
                if cell in ws:
                    ws[cell] = previous_year_total
                    ws[cell].number_format = '#,##0.00'
                    updates_log.append(f"✓ Updated {cell} with previous year total for {parking_code}: ${previous_year_total:,.2f}")
                    break
    except Exception as e:
        updates_log.append(f"⚠️ Error updating Budget Initial for {parking_code}: {e}")
    return updates_log


def update_fiche_stationnement(wb, pnl_file, parking_code, word_data=None):
    """
    Update the 1. Fiche Stationnement sheet with revenue data and additional data from Word.
    Dynamically matches ANY revenue category found in the P&L file.
    """
    updates_log = []
    try:
        if "1. Fiche Stationnement" in wb.sheetnames:
            ws = wb["1. Fiche Stationnement"]
            
            # Update revenue data (K17 to K25)
            revenue_data = extract_2024_revenue_data(pnl_file, parking_code)
            
            # Dynamically map revenue categories to cells K17-K25
            # If a category is not in CELL_MAPPINGS, it will be assigned to the next available cell (K17, K18, etc.)
            used_cells = []
            for category, value in revenue_data.items():
                # Check if the category is already mapped
                if category in CELL_MAPPINGS["1. Fiche Stationnement"]["revenue_categories"]:
                    cells = CELL_MAPPINGS["1. Fiche Stationnement"]["revenue_categories"][category]
                else:
                    # Assign to the next available cell (K17, K18, etc.)
                    cells = [f"K{17 + len(used_cells)}"]
                
                for cell in cells:
                    if cell in ws and cell not in used_cells:
                        ws[cell] = value
                        ws[cell].number_format = '#,##0.00'
                        updates_log.append(f"✓ Updated {cell} with {category}: ${value:,.2f}")
                        used_cells.append(cell)
                        break
            
            # Update total revenue (K26)
            total_revenue = sum(revenue_data.values())
            for cell in CELL_MAPPINGS["1. Fiche Stationnement"]["total_revenue"]:
                if cell in ws:
                    ws[cell] = total_revenue
                    ws[cell].number_format = '#,##0.00'
                    updates_log.append(f"✓ Updated {cell} with Total Revenue for {parking_code}: ${total_revenue:,.2f}")
                    break
            
            # Update additional data (H-M, rows 43-55) from Word sheet
            if word_data:
                for row in CELL_MAPPINGS["1. Fiche Stationnement"]["additional_data_rows"]:
                    for field, col in CELL_MAPPINGS["1. Fiche Stationnement"]["additional_data"].items():
                        if isinstance(col, list):  # Multiple columns (e.g., Avant taxes)
                            for c in col:
                                cell = f"{c}{row}"
                                if cell in ws and field in word_data:
                                    ws[cell] = word_data[field]
                                    updates_log.append(f"✓ Updated {cell} with {field}: {word_data[field]}")
                        else:  # Single column
                            cell = f"{col}{row}"
                            if cell in ws and field in word_data:
                                ws[cell] = word_data[field]
                                updates_log.append(f"✓ Updated {cell} with {field}: {word_data[field]}")
    except Exception as e:
        updates_log.append(f"⚠️ Error updating Fiche Stationnement for {parking_code}: {e}")
    return updates_log


def update_donnees_historiques(wb, pnl_file, parking_code):
    """
    Update the 2. Donnees Historiques sheet with Jan-Apr 2026 and May-Dec 2025 data for ANY parking code.
    """
    updates_log = []
    try:
        if "2. Donnees Historiques" in wb.sheetnames:
            ws = wb["2. Donnees Historiques"]
            monthly_data = extract_2026_and_2025_data(pnl_file, parking_code)
            
            rows = CELL_MAPPINGS["2. Donnees Historiques"]["monthly_data"]["rows"]
            skip_rows = CELL_MAPPINGS["2. Donnees Historiques"]["monthly_data"]["skip_rows"]
            columns = CELL_MAPPINGS["2. Donnees Historiques"]["monthly_data"]["columns"]
            
            months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            
            for row in rows:
                if row in skip_rows:
                    continue  # Skip white/formula cells
                for col_idx, month in enumerate(months):
                    cell = f"{columns[col_idx]}{row}"
                    if cell in ws and month in monthly_data:
                        ws[cell] = monthly_data[month]
                        ws[cell].number_format = '#,##0.00'
                        updates_log.append(f"✓ Updated {cell} with {month} data for {parking_code}: ${monthly_data[month]:,.2f}")
    except Exception as e:
        updates_log.append(f"⚠️ Error updating Donnees Historiques for {parking_code}: {e}")
    return updates_log


def fix_excel(excel_file, pnl_file, config):
    """
    Main function to fix the Excel template based on the config.
    
    Args:
        excel_file: The uploaded Excel template file (BytesIO or file-like object).
        pnl_file: The uploaded P&L file (BytesIO or file-like object).
        config: Dictionary with configuration options:
            - parking_code: str (e.g., "CMO111", "CMO142", etc.)
            - parking_type: str (e.g., "SC" or "RG")
            - supervisor_hours: float (e.g., 1.0)
            - word_data: dict (optional, for Fiche Stationnement additional data)
    
    Returns:
        tuple: (fixed_excel_bytes, updates_log)
            - fixed_excel_bytes: BytesIO object with the updated Excel file.
            - updates_log: List of strings describing the updates applied.
    """
    wb = load_workbook(io.BytesIO(excel_file.read()))
    updates_log = []
    
    # Update each sheet based on config
    updates_log.extend(update_budget_initial(wb, pnl_file, config['parking_code']))
    updates_log.extend(update_fiche_stationnement(wb, pnl_file, config['parking_code'], config.get('word_data')))
    updates_log.extend(update_donnees_historiques(wb, pnl_file, config['parking_code']))
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, updates_log
