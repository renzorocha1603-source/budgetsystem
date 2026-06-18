# excel_fixer.py - WITH READ DEBUG
import io
import re
import pandas as pd
import json
import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
import os

# ============================================================================
# MISTRAL CONFIGURATION
# ============================================================================
MISTRAL_API_KEY = "em5oqjSdA1Nus9iUpa1MNAJtQA4YfCtK"
MISTRAL_URL = "https://api.mistral.ai/v1/chat/completions"
MISTRAL_MODEL = "mistral-small-latest"

# ============================================================================
# CORRECTED P&L MAPPING
# ============================================================================
PANDL_TO_TEMPLATE = {
    23: 12,   # Transient Revenue
    17: 13,   # Monthly Revenues
    26: 14,   # Car-Wash Revenue
    24: 15,   # Hotel Revenue
    31: 16,   # Interests
    28: 17,   # Miscellaneous
    33: 20,   # Discount-Gratuities - Transient
    34: 22,   # Discount-Gratuities - Monthly
    38: 29,   # Parking wages
    40: 30,   # Other wages
    41: 31,   # Training & Recr.
    42: 32,   # Uniforms
    45: 35,   # R&M - Cleaning
    50: 36,   # R&M - General
    46: 37,   # R&M - Equipment
    47: 38,   # R&M - Signs
    48: 39,   # R&M - Lines
    54: 40,   # Snow Removal
    43: 41,   # Parking supplies
    44: 42,   # Misc. Re-Billing
    60: 46,   # Public services
    72: 49,   # Office expenses
    64: 50,   # Telecommunication
    55: 51,   # Rent
    59: 52,   # Vehicle expenses
    68: 53,   # Credit Card fees
    70: 54,   # Bank fees
    69: 55,   # Cash transportation fees
    63: 56,   # Claims
    62: 57,   # Insurance & Guarantee
    61: 58,   # Tax & license
    65: 59,   # Professional services
    56: 60,   # Equipment rent
    67: 61,   # Ad. & Promotion
    81: 62,   # Percent Management fee
    80: 63,   # Management Fees (Basic)
    84: 64,   # Incentives
    85: 67,   # Depreciation
    53: 69,   # Security
    57: 70,   # Co-ownership expenses
    58: 71,   # Shuttle expenses
    66: 72,   # Computer services
}

PANDL_VALIDATION = {
    29: "_PARKING_REVENUE_",
    36: "_TOTAL_REVENUS_",
    75: "_TOTAL_EXPENSES_",
    77: "_OPERATION_SURPLUS_",
    92: "_BENEFICE_NET_",
}

PANDL_MONTH_COLUMNS = {
    'January': 3,
    'February': 4,
    'March': 5,
    'April': 6,
}

# ============================================================================
# MONTH DETECTION
# ============================================================================

MONTH_MAP = {
    'janvier': 'January', 'février': 'February', 'fevrier': 'February',
    'mars': 'March', 'avril': 'April', 'mai': 'May', 'juin': 'June',
    'juillet': 'July', 'août': 'August', 'aout': 'August',
    'septembre': 'September', 'octobre': 'October',
    'novembre': 'November', 'décembre': 'December', 'decembre': 'December'
}

MONTH_COLUMN = {
    'January': 2, 'February': 3, 'March': 4, 'April': 5,
    'May': 6, 'June': 7, 'July': 8, 'August': 9,
    'September': 10, 'October': 11, 'November': 12, 'December': 13
}

# ============================================================================
# EXCEL P&L EXTRACTION
# ============================================================================

def extract_from_pnl_excel(file_bytes, parking_code, debug_updates=None):
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        
        tab_name = None
        for sn in xl.sheet_names:
            if parking_code.upper() in sn.upper():
                tab_name = sn
                break
        
        if tab_name is None:
            if debug_updates is not None:
                debug_updates.append(f"❌ Tab not found for {parking_code}")
            return {}
        
        if debug_updates is not None:
            debug_updates.append(f"📊 Found tab: {tab_name}")
        
        df = pd.read_excel(xl, sheet_name=tab_name, header=None)
        
        if debug_updates is not None:
            debug_updates.append(f"📐 P&L: {len(df)} rows x {len(df.columns)} cols")
            
            # DEBUG: Show what we're actually reading
            debug_updates.append("🔍 READ DEBUG - What values are we getting:")
            for pnl_row in [17, 23, 24, 26, 28, 31, 33, 34, 36, 38, 40, 42, 45, 46, 50, 51, 53, 62, 64, 68, 75, 77, 92]:
                if pnl_row <= len(df):
                    col_a = str(df.iloc[pnl_row-1, 0])[:50] if len(df.columns) > 0 else 'N/A'
                    col_c = df.iloc[pnl_row-1, 2] if len(df.columns) > 2 else 'N/A'
                    col_d = df.iloc[pnl_row-1, 3] if len(df.columns) > 3 else 'N/A'
                    debug_updates.append(f"  P&L Row {pnl_row}: A='{col_a}' | Jan={col_c} | Feb={col_d}")
        
        data = {}
        
        # Extract template accounts
        for pnl_row, template_row in PANDL_TO_TEMPLATE.items():
            if pnl_row > len(df):
                continue
            
            for month_name, col_idx in PANDL_MONTH_COLUMNS.items():
                if col_idx < len(df.columns):
                    val = df.iloc[pnl_row - 1, col_idx - 1]
                    try:
                        amount = float(val) if pd.notna(val) else 0.0
                    except (ValueError, TypeError):
                        amount = 0.0
                    
                    if month_name not in data:
                        data[month_name] = {}
                    data[month_name][template_row] = amount
        
        # Extract validation totals
        for pnl_row, validation_key in PANDL_VALIDATION.items():
            if pnl_row > len(df):
                continue
            
            for month_name, col_idx in PANDL_MONTH_COLUMNS.items():
                if col_idx < len(df.columns):
                    val = df.iloc[pnl_row - 1, col_idx - 1]
                    try:
                        amount = float(val) if pd.notna(val) else 0.0
                    except (ValueError, TypeError):
                        amount = 0.0
                    
                    if month_name not in data:
                        data[month_name] = {}
                    data[month_name][validation_key] = amount
        
        if debug_updates is not None:
            debug_updates.append(f"✅ Extracted {len(data)} months from P&L")
            
            # DEBUG: Show final data
            for month_name in ['January', 'February']:
                if month_name in data:
                    debug_updates.append(f"  Final {month_name} data (first 10):")
                    for k, v in list(data[month_name].items())[:10]:
                        debug_updates.append(f"    Row {k}: {v}")
        
        return data
        
    except Exception as e:
        if debug_updates is not None:
            debug_updates.append(f"❌ P&L error: {e}")
        return {}

# ============================================================================
# ALLISON VALIDATION
# ============================================================================

def validate_with_allison(all_data, debug_updates=None):
    if debug_updates is not None:
        debug_updates.append("🤖 Allison validating...")
    
    summary = ""
    for month_name, month_data in all_data.items():
        ben = month_data.get('_BENEFICE_NET_', 'N/A')
        rev = month_data.get('_TOTAL_REVENUS_', 'N/A')
        exp = month_data.get('_TOTAL_EXPENSES_', 'N/A')
        summary += f"\n{month_name}: Net={ben}, Rev={rev}, Exp={exp}\n"
    
    prompt = f"""Validate this P&L extraction quickly:
{summary}
Check: Does NET INCOME = TOTAL REVENUS - TOTAL EXPENSES?
Reply "VALID" or list issues briefly."""
    
    try:
        resp = requests.post(
            MISTRAL_URL,
            headers={"Authorization": f"Bearer {MISTRAL_API_KEY}", "Content-Type": "application/json"},
            json={"model": MISTRAL_MODEL, "messages": [{"role": "user", "content": prompt}], "temperature": 0.1},
            timeout=30
        )
        if resp.status_code == 200:
            return resp.json()["choices"][0]["message"]["content"]
    except:
        pass
    return None

# ============================================================================
# MAIN EXTRACTION
# ============================================================================

def extract_monthly_data(file_obj, parking_code, debug_updates=None):
    file_obj.seek(0)
    file_bytes = file_obj.read()
    
    if hasattr(file_obj, 'name') and file_obj.name.lower().endswith(('.xlsx', '.xls')):
        if debug_updates is not None:
            debug_updates.append("📊 Excel file detected")
        return extract_from_pnl_excel(file_bytes, parking_code, debug_updates)
    
    return {}

# ============================================================================
# TEMPLATE FILLING
# ============================================================================

def fill_template(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        return ["❌ Sheet not found"]
    
    ws = wb[sheet_name]
    updates = []
    total = 0
    
    for month_en, month_data in all_data.items():
        col = MONTH_COLUMN.get(month_en)
        if col is None:
            continue
        
        col_letter = get_column_letter(col)
        month_cells = 0
        
        for key, amount in month_data.items():
            if str(key).startswith('_'):
                continue
            ws.cell(row=key, column=col).value = amount
            ws.cell(row=key, column=col).number_format = '#,##0.00 $'
            month_cells += 1
            total += 1
            updates.append(f"   ✅ {month_en} ({col_letter}{key}): {amount:,.2f} $")
        
        if month_cells > 0:
            updates.append(f"📊 {month_en}: {month_cells} cells")
    
    updates.append(f"\n📊 TOTAL: {total} yellow cells (formulas auto-calculate)")
    return updates

# ============================================================================
# VALIDATION
# ============================================================================

def validate_results(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        return ["⚠️ Cannot validate"]
    
    results = []
    
    for month_en, month_data in all_data.items():
        col = MONTH_COLUMN.get(month_en)
        if col is None:
            continue
        
        ben = month_data.get('_BENEFICE_NET_')
        rev = month_data.get('_TOTAL_REVENUS_')
        exp = month_data.get('_TOTAL_EXPENSES_')
        
        results.append(f"\n📊 {month_en}:")
        
        if ben is not None and rev is not None and exp is not None:
            expected = rev - exp
            diff = abs(ben - expected)
            if diff > 0.01:
                results.append(f"   ⚠️ P&L Net: {ben:,.2f} $ | Expected: {expected:,.2f} $")
            else:
                results.append(f"   ✅ NET INCOME = {ben:,.2f} $")
        elif ben is not None:
            results.append(f"   ⚠️ Net: {ben:,.2f} $")
    
    return results

# ============================================================================
# APP.PY INTERFACE
# ============================================================================

def get_parking_codes_from_pnl(file_obj):
    codes = []
    if hasattr(file_obj, 'name'):
        file_obj.seek(0)
        if file_obj.name.lower().endswith(('.xlsx', '.xls')):
            try:
                xl = pd.ExcelFile(io.BytesIO(file_obj.read()))
                for sn in xl.sheet_names:
                    match = re.search(r'(CMO\d+|VMO\d+)', sn, re.IGNORECASE)
                    if match:
                        codes.append(match.group(1).upper())
            except:
                pass
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    updates = []
    all_data = {}
    debug = []
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append("=" * 60)
        
        if isinstance(excel_file, bytes):
            wb = load_workbook(io.BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(io.BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template: {parking_code or 'Unknown'}")
        
        all_files = []
        if monthly_files_current:
            all_files.extend(monthly_files_current)
        if monthly_files_previous:
            all_files.extend(monthly_files_previous)
        
        if not all_files:
            updates.append("⚠️ No files!")
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue(), updates
        
        updates.append(f"\n📁 {len(all_files)} files...")
        
        for file_obj in all_files:
            file_obj.seek(0)
            try:
                debug.append(f"--- {getattr(file_obj, 'name', 'unknown')} ---")
                data = extract_monthly_data(file_obj, parking_code, debug)
                
                if data:
                    for mn, md in data.items():
                        all_data[mn] = md
                        cnt = len([k for k in md if not str(k).startswith('_')])
                        ben = md.get('_BENEFICE_NET_', 'N/A')
                        updates.append(f"   ✅ {mn}: {cnt} accounts | Net: {ben} $")
                else:
                    updates.append(f"   ❌ No data")
            finally:
                pass
        
        if debug:
            updates.append("\n🔍 Debug:")
            updates.extend(debug)
        
        if all_data:
            updates.append(f"\n📝 Filling {len(all_data)} months...")
            updates.extend(fill_template(wb, all_data))
            updates.append(f"\n🤖 Allison:")
            ar = validate_with_allison(all_data, debug)
            if ar:
                updates.append(f"   {ar[:300]}")
            updates.append(f"\n🔍 Validation:")
            updates.extend(validate_results(wb, all_data))
        else:
            updates.append("\n⚠️ No data!")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        updates.append("\n✅ DONE")
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"\n❌ {e}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates
