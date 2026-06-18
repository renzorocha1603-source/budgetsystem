# excel_fixer.py - MINIMAL WORKING VERSION
import re
import io
import fitz
import tempfile
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Simple keyword mapping (Grok's idea)
ACCOUNT_MAPPING = {
    "revenus journaliers|daily|horaires": 12,
    "revenus mensuels|monthly": 13,
    "revenus lave-auto|car wash|lave": 14,
    "divers|autres revenus|miscellaneous": 17,
    "gratuités|gratuities": 20,
    "salaires stationnement|parking salaries": 29,
    "uniformes|uniforms": 32,
    "nettoyage|cleaning": 35,
    "entretien stationnement|général|general": 36,
    "entretien équipement|equipment": 37,
    "fourn. de stationnement|supplies|parking supplies": 41,
    "télécommunication|telecom": 50,
    "frais de cartes de crédit|credit card": 53,
    "frais de bureau|office expenses": 49,
    "réclamations|claims": 56,
    "assurances|insurance|cautionnement": 57,
    "taxes et permis|taxes": 58,
    "honoraires de gestion|management fees": 63,
}

MONTH_MAP = {
    'janvier': 'January', 'février': 'February', 'mars': 'March',
    'avril': 'April', 'mai': 'May', 'juin': 'June',
    'juillet': 'July', 'août': 'August', 'septembre': 'September',
    'octobre': 'October', 'novembre': 'November', 'décembre': 'December'
}

MONTH_COLUMN = {
    'January': 2, 'February': 3, 'March': 4, 'April': 5,
    'May': 6, 'June': 7, 'July': 8, 'August': 9,
    'September': 10, 'October': 11, 'November': 12, 'December': 13
}

def safe_float(value):
    """Clean Canadian French number: '43 585,46 $' -> 43585.46"""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip()
    is_negative = value.startswith('(') and value.endswith(')')
    if is_negative:
        value = value[1:-1]
    value = value.replace('$', '').replace(' ', '').replace('\xa0', '')
    value = value.replace(',', '.')
    try:
        result = float(value)
        return -result if is_negative else result
    except:
        return None

def find_row(pdf_label):
    """Match PDF account name to template row"""
    pdf_lower = pdf_label.lower()
    for key, row in ACCOUNT_MAPPING.items():
        for keyword in key.split('|'):
            if keyword.strip() in pdf_lower:
                return row
    return None

def extract_from_pdf(pdf_path):
    """Extract P&L data from PDF using text search"""
    doc = fitz.open(pdf_path)
    data = {}
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text("text")
        
        if "revenus mensuels" not in text.lower():
            continue
        
        # Search line by line
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if len(line) < 5:
                continue
            
            row = find_row(line)
            if row is None:
                continue
            
            # Find amount in this line
            # Canadian French format: "43 585,46 $" or "(1 206,86) $"
            match = re.search(r'\(?\s*(\d[\d\s]*,\d{2})\s*\)?\s*\$?', line)
            if match:
                amount = safe_float(match.group(1))
                if amount is not None and amount != 0:
                    data[row] = amount
        
        break  # Only process first matching page
    
    doc.close()
    return data

def get_month_from_filename(file_obj):
    """Extract month from filename"""
    name = file_obj.name.lower() if hasattr(file_obj, 'name') else ''
    for fr, en in MONTH_MAP.items():
        if fr in name:
            return en
    return None

def get_parking_codes_from_pnl(file_obj):
    """Extract parking codes"""
    codes = []
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """Main function - extract PDF data and fill template"""
    updates = []
    
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - WORKFLOW")
        updates.append("=" * 60)
        
        # Load template
        if isinstance(excel_file, bytes):
            wb = load_workbook(io.BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(io.BytesIO(excel_file.read()))
        
        # Find Donnees Historiques sheet
        sheet_name = None
        for sn in wb.sheetnames:
            if 'données' in sn.lower() or 'historique' in sn.lower():
                sheet_name = sn
                break
        if sheet_name is None:
            sheet_name = wb.sheetnames[0]
        
        ws = wb[sheet_name]
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
        # Process files
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
        
        updates.append(f"📁 Processing {len(all_files)} files...")
        total_cells = 0
        
        for file_obj in all_files:
            month_en = get_month_from_filename(file_obj)
            if month_en is None:
                updates.append(f"⚠️ Could not determine month for {file_obj.name}")
                continue
            
            col = MONTH_COLUMN.get(month_en)
            if col is None:
                continue
            
            # Save PDF to temp file
            file_obj.seek(0)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            try:
                data = extract_from_pdf(tmp_path)
                if data:
                    month_cells = 0
                    for row, amount in data.items():
                        cell = ws.cell(row=row, column=col)
                        cell.value = amount
                        cell.number_format = '#,##0.00'
                        month_cells += 1
                        total_cells += 1
                    
                    col_letter = get_column_letter(col)
                    updates.append(f"✅ {month_en} ({col_letter}): {month_cells} cells filled")
                else:
                    updates.append(f"❌ {month_en}: No data extracted")
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
        
        updates.append(f"\n📊 TOTAL: {total_cells} cells filled")
        
        # Save
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        updates.append("\n✅ WORKFLOW COMPLETE")
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"\n❌ ERROR: {str(e)}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates
