"""
Excel Fixer Module - Pure budget calculation logic
This file handles ONLY the Excel processing. Easy to modify without touching the UI.
"""

import pandas as pd
import pdfplumber
from openpyxl import load_workbook
import io
import re

# Seasonal multipliers for different parking types
SEASONAL_MULTIPLIERS = {
    "SC": [1.2, 1.2, 1.2, 1.1, 0.8, 0.8, 0.8, 0.8, 1.1, 1.1, 1.2, 1.2],
    "RG": [0.8, 0.8, 0.9, 0.9, 1.3, 1.3, 1.3, 1.2, 1.0, 1.0, 0.8, 0.8]
}

# Cell mappings - EASY TO MODIFY based on your Excel structure
CELL_MAPPINGS = {
    'transient': ['K17', 'J17', 'L17', 'C17', 'D17', 'E17', 'B17'],
    'monthly': ['K18', 'J18', 'L18', 'C18', 'D18', 'E18', 'B18'],
    'total': ['K26', 'J26', 'L26', 'C26', 'D26', 'E26', 'K30', 'K32'],
    'supervisor': ['K30', 'K31', 'D10', 'C10', 'E10', 'K40', 'B10'],
    'monthly_start_row': [20, 21, 22]  # Try these row numbers for monthly projections
}

# Hourly rates - EASY TO MODIFY
HOURLY_RATE = 25.0

def extract_pdf_data(pdf_file):
    """
    Extract revenue data from PDF
    Returns: dict with transient, monthly, total revenue
    """
    result = {
        'transient': 0,
        'monthly': 0,
        'total': 0
    }
    
    with pdfplumber.open(pdf_file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() or ""
    
    # Find numbers in PDF
    numbers = re.findall(r'\$?([\d,]+\.?\d*)', text)
    result['total'] = float(numbers[0].replace(',', '')) if numbers else 40000
    
    # Search for specific revenue types
    for line in text.split('\n'):
        if 'transient' in line.lower():
            nums = re.findall(r'\$?([\d,]+\.?\d*)', line)
            if nums:
                result['transient'] = float(nums[0].replace(',', ''))
        if 'monthly' in line.lower() or 'subscription' in line.lower():
            nums = re.findall(r'\$?([\d,]+\.?\d*)', line)
            if nums:
                result['monthly'] = float(nums[0].replace(',', ''))
    
    # If not found, estimate
    if result['transient'] == 0 and result['monthly'] == 0:
        result['transient'] = result['total'] * 0.6
        result['monthly'] = result['total'] * 0.4
    
    return result

def fix_excel(excel_file, revenue_data, config):
    """
    Fix Excel template with extracted data
    config: {parking_name, parking_type, supervisor_hours}
    Returns: BytesIO object with fixed Excel file
    """
    
    # Load workbook
    wb = load_workbook(io.BytesIO(excel_file.read()))
    ws = wb.active
    
    updates_log = []
    
    # 1. Update revenue cells
    updates_log.extend(_update_cell(ws, CELL_MAPPINGS['transient'], revenue_data['transient'], "Transient Revenue"))
    updates_log.extend(_update_cell(ws, CELL_MAPPINGS['monthly'], revenue_data['monthly'], "Monthly Revenue"))
    updates_log.extend(_update_cell(ws, CELL_MAPPINGS['total'], revenue_data['total'], "Total Revenue"))
    
    # 2. Apply seasonal multipliers for monthly projections
    parking_type_key = "SC" if config['parking_type'] == "SC (School)" else "RG"
    multipliers = SEASONAL_MULTIPLIERS.get(parking_type_key, SEASONAL_MULTIPLIERS["SC"])
    
    base_monthly = revenue_data['total'] / 12
    monthly_updates = 0
    
    for i, mult in enumerate(multipliers):
        projected = base_monthly * mult
        for row_offset in CELL_MAPPINGS['monthly_start_row']:
            cell = f'K{row_offset + i}'
            try:
                if cell in ws:
                    ws[cell] = projected
                    ws[cell].number_format = '#,##0.00'
                    monthly_updates += 1
                    break
            except:
                pass
    
    if monthly_updates > 0:
        updates_log.append(f"✓ Updated {monthly_updates} monthly projection cells")
    
    # 3. Update supervisor cost
    supervisor_cost = HOURLY_RATE * config['supervisor_hours'] * 30
    updates_log.extend(_update_cell(ws, CELL_MAPPINGS['supervisor'], supervisor_cost, "Supervisor Cost"))
    
    # 4. Add formulas if cells exist
    try:
        if 'K31' in ws:
            ws['K31'] = '=SUM(K20:K30)'
            updates_log.append("✓ Added total expenses formula (K31)")
        if 'K32' in ws:
            ws['K32'] = '=K26-K31'
            updates_log.append("✓ Added net profit formula (K32)")
    except:
        pass
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, updates_log

def _update_cell(ws, cell_list, value, label):
    """Helper to update a cell, trying multiple possible locations"""
    for cell in cell_list:
        try:
            if cell in ws:
                ws[cell] = value
                ws[cell].number_format = '#,##0.00'
                return [f"✓ Updated {cell} with {label}: ${value:,.2f}"]
        except:
            pass
    return [f"⚠️ Could not find cell for {label} (tried: {', '.join(cell_list[:3])})"]

# Easy configuration - just edit this dictionary
DEFAULT_CONFIG = {
    "parking_name": "CMO142",
    "parking_type": "SC (School)",
    "supervisor_hours": 1
}
