# excel_fixer.py - SEARCH BY ACCOUNT NAME (NO HARDCODED ROWS)
import io
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
import os

# ============================================================================
# ACCOUNT NAME SEARCH PATTERNS (P&L account name -> Template Row)
# ============================================================================
ACCOUNT_SEARCH = {
    # REVENUES
    "transient revenue": 12,
    "revenus horaires": 12,
    "monthly revenues": 13,
    "revenus mensuels": 13,
    "car-wash revenue": 14,
    "revenus lave-auto": 14,
    "hotel revenue": 15,
    "revenus hotel": 15,
    "revenus hôtel": 15,
    "interests": 16,
    "intérêts": 16,
    "interets": 16,
    "miscellaneous": 17,
    "autres revenus": 17,
    "discount-gratuities - transient": 20,
    "gratuities - transient": 20,
    "discount-gratuities - monthly": 22,
    "gratuities - monthly": 22,
    
    # LABOUR
    "parking wages": 29,
    "salaire stationnement": 29,
    "other wages": 30,
    "salaire superviseur": 30,
    "training & recr.": 31,
    "formation & recrutement": 31,
    "uniformes": 32,
    "uniforms": 32,
    
    # MAINTENANCE
    "r&m - cleaning": 35,
    "nettoyage stationnement": 35,
    "r&m - general": 36,
    "entretien stationnement": 36,
    "r&m - equipment": 37,
    "entretien équipement": 37,
    "entretien equipement": 37,
    "r&m - signs": 38,
    "signalisation": 38,
    "r&m - lines": 39,
    "lignage": 39,
    "snow removal": 40,
    "déneigement": 40,
    "deneigement": 40,
    "parking supplies": 41,
    "fournitures stationnement": 41,
    "misc. re-billing": 42,
    "refacturations diverses": 42,
    
    # PUBLIC SERVICES
    "public services": 46,
    "services publics": 46,
    
    # OVERHEAD
    "office expenses": 49,
    "fournitures de bureau": 49,
    "telecommunication": 50,
    "télécommunication": 50,
    "télécommunications": 50,
    "rent": 51,
    "loyer": 51,
    "travel expenses": 52,
    "frais de déplacement": 52,
    "frais de deplacement": 52,
    "credit card fees": 53,
    "frais de cartes de crédit": 53,
    "frais de cartes de credit": 53,
    "bank fees": 54,
    "intérêts et frais de banque": 54,
    "interets et frais de banque": 54,
    "cash transportation fees": 55,
    "transport de fonds": 55,
    "claims": 56,
    "réclamations": 56,
    "reclamations": 56,
    "insurance & guarantee": 57,
    "assurances et cautionnement": 57,
    "tax & license": 58,
    "taxes et permis": 58,
    "professional services": 59,
    "comptabilité": 59,
    "comptabilite": 59,
    "equipment rent": 60,
    "location d'équipement": 60,
    "location d'equipement": 60,
    "ad. & promotion": 61,
    "publicité et promotion": 61,
    "publicite et promotion": 61,
    "percent management fee": 62,
    "honoraires de gestion en %": 62,
    "management fees (basic)": 63,
    "honoraires de gestion de base": 63,
    "incentives": 64,
    "incitatif annuel": 64,
    
    # OTHER EXPENSES
    "depreciation": 67,
    "amortissement": 67,
    "financial fees": 68,
    "intérêts sur emprunts": 68,
    "interets sur emprunts": 68,
    "security": 69,
    "sécurité": 69,
    "securite": 69,
    "co-ownership expenses": 70,
    "frais de copropriété": 70,
    "frais de copropriete": 70,
    "shuttle expenses": 71,
    "frais de navettes": 71,
    "computer services": 72,
    "services informatiques": 72,
    "bad debts": 73,
    "mauvaises créances": 73,
    "mauvaises creances": 73,
    "dues & subscription": 74,
    "cotisations": 74,
    "meal & entertainment": 76,
    "représentation repas": 76,
    "representation repas": 76,
}

# Validation accounts
VALIDATION_SEARCH = {
    "total revenue": "_TOTAL_REVENUS_",
    "total revenus": "_TOTAL_REVENUS_",
    "total des revenus": "_TOTAL_REVENUS_",
    "total operation expenses": "_TOTAL_EXPENSES_",
    "total operating expenses": "_TOTAL_EXPENSES_",
    "total des frais d'exploitation": "_TOTAL_EXPENSES_",
    "operation surplus": "_OPERATION_SURPLUS_",
    "operating surplus": "_OPERATION_SURPLUS_",
    "net income": "_BENEFICE_NET_",
    "bénéfice net": "_BENEFICE_NET_",
    "benefice net": "_BENEFICE_NET_",
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
# EXCEL P&L EXTRACTION - SEARCH BY NAME
# ============================================================================

def extract_from_pnl_excel(file_bytes, parking_code, debug_updates=None):
    """Extract data by searching account names in Column A."""
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
            debug_updates.append("🔍 Searching for accounts by name...")
        
        data = {}
        
        # Scan every row for matching account names
        for row_idx in range(len(df)):
            col_a = str(df.iloc[row_idx, 0]).strip().lower()
            if not col_a or col_a == 'nan':
                continue
            
            # Check template accounts
            for search_term, template_row in ACCOUNT_SEARCH.items():
                if search_term in col_a:
                    # Found a match! Extract values for all months
                    for month_name, col_idx in PANDL_MONTH_COLUMNS.items():
                        if col_idx < len(df.columns):
                            val = df.iloc[row_idx, col_idx - 1]
                            try:
                                amount = float(val) if pd.notna(val) else 0.0
                            except (ValueError, TypeError):
                                amount = 0.0
                            
                            if month_name not in data:
                                data[month_name] = {}
                            
                            # Only store if not already found (first match wins)
                            if template_row not in data[month_name]:
                                data[month_name][template_row] = amount
                                if debug_updates is not None and month_name == 'January':
                                    debug_updates.append(f"  ✅ Row {row_idx+1}: '{col_a[:50]}' → Template {template_row} = {amount}")
                    break  # Found match, stop searching for this row
            
            # Check validation accounts
            for search_term, validation_key in VALIDATION_SEARCH.items():
                if search_term in col_a:
                    for month_name, col_idx in PANDL_MONTH_COLUMNS.items():
                        if col_idx < len(df.columns):
                            val = df.iloc[row_idx, col_idx - 1]
                            try:
                                amount = float(val) if pd.notna(val) else 0.0
                            except (ValueError, TypeError):
                                amount = 0.0
                            
                            if month_name not in data:
                                data[month_name] = {}
                            
                            if validation_key not in data[month_name]:
                                data[month_name][validation_key] = amount
                                if debug_updates is not None and month_name == 'January':
                                    debug_updates.append(f"  📊 Row {row_idx+1}: '{col_a[:50]}' → {validation_key} = {amount}")
                    break
        
        if debug_updates is not None:
            debug_updates.append(f"✅ Extracted {len(data)} months from P&L")
        
        return data
        
    except Exception as e:
        if debug_updates is not None:
            debug_updates.append(f"❌ P&L error: {e}")
        return {}

# ============================================================================
# MAIN EXTRACTION
# ============================================================================

def extract_monthly_data(file_obj, parking_code, debug_updates=None):
    """Extract data from Excel P&L file."""
    file_obj.seek(0)
    file_bytes = file_obj.read()
    
    if hasattr(file_obj, 'name') and file_obj.name.lower().endswith(('.xlsx', '.xls')):
        if debug_updates is not None:
            debug_updates.append("📊 Excel file detected")
        return extract_from_pnl_excel(file_bytes, parking_code, debug_updates)
    
    if debug_updates is not None:
        debug_updates.append("❌ Not an Excel file")
    return {}

# ============================================================================
# TEMPLATE FILLING
# ============================================================================

def fill_template(wb, all_data):
    """Write extracted data to YELLOW cells only."""
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        return ["❌ Sheet '2-Données historiques' not found"]
    
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
    """Check if NET INCOME is consistent."""
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
                results.append(f"   ⚠️ P&L Net: {ben:,.2f} $ | Expected: {expected:,.2f} $ (diff: {diff:,.2f} $)")
            else:
                results.append(f"   ✅ NET INCOME = {ben:,.2f} $")
        elif ben is not None:
            results.append(f"   ⚠️ Net: {ben:,.2f} $")
    
    return results

# ============================================================================
# APP.PY INTERFACE
# ============================================================================

def get_parking_codes_from_pnl(file_obj):
    """Extract all parking codes from P&L Excel tabs."""
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
        else:
            matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
            codes.extend(matches)
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """Main function - extract from P&L Excel and fill template."""
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
                        updates.append(f"   ✅ {mn}: {cnt} accounts | P&L Net: {ben} $")
                else:
                    updates.append(f"   ❌ No data extracted")
            finally:
                pass
        
        if debug:
            updates.append("\n🔍 Debug:")
            updates.extend(debug)
        
        if all_data:
            updates.append(f"\n📝 Filling {len(all_data)} months...")
            updates.extend(fill_template(wb, all_data))
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
