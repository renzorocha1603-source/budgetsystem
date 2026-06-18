# excel_fixer.py - PDF TO EXCEL CONVERSION VERSION
import fitz
import pandas as pd
import os
import re
from io import BytesIO
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

PAGE_MARKERS = ["1981 McGill College", "Revenus mensuels", "BÉNÉFICE NET", "Mois Courant"]

# Account name matching (what to look for in Column A of converted Excel)
ACCOUNT_PATTERNS = [
    ("total des frais d'exploitation", "Total des frais d'exploitation"),
    ("résultat d'exploitation", "RÉSULTAT D'EXPLOITATION"),
    ("resultat d'exploitation", "RÉSULTAT D'EXPLOITATION"),
    ("total revenus", "TOTAL REVENUS"),
    ("bénéfice net", "BÉNÉFICE NET"),
    ("benefice net", "BÉNÉFICE NET"),
    ("revenus mensuels", "Revenus mensuels"),
    ("revenus journaliers", "Revenus Journaliers"),
    ("revenus horaires", "Revenus Journaliers"),
    ("revenus lave-auto", "Revenus Lave-Auto"),
    ("gratuités", "Gratuités - mensuels"),
    ("salaires stationnement", "Salaires Stationnement"),
    ("salaire stationnement", "Salaires Stationnement"),
    ("uniformes", "Uniformes"),
    ("fourn. de stationnement", "Fourn. de stationnement"),
    ("fournitures stationnement", "Fourn. de stationnement"),
    ("nettoyage", "Entretien réparation - Nettoyage"),
    ("equipement", "Entretien réparation - Equipement"),
    ("équipement", "Entretien réparation - Equipement"),
    ("général", "Entretien réparation - Général"),
    ("general", "Entretien réparation - Général"),
    ("taxes et permis", "Taxes et permis"),
    ("assurances", "Assurances Cautionnement"),
    ("réclamations", "Réclamations"),
    ("reclamations", "Réclamations"),
    ("télécommunication", "Télécommunication"),
    ("telecommunication", "Télécommunication"),
    ("cartes de crédit", "Frais de cartes de crédit"),
    ("frais de bureau", "Frais de bureau"),
    ("honoraires de gestion", "Honoraires de gestion"),
    ("divers", "Divers"),
]

MONTH_COLUMN = {
    'Janvier': 2, 'Février': 3, 'Mars': 4, 'Avril': 5,
    'Mai': 6, 'Juin': 7, 'Juillet': 8, 'Août': 9,
    'Septembre': 10, 'Octobre': 11, 'Novembre': 12, 'Décembre': 13
}

PDF_TO_TEMPLATE = {
    'Revenus Journaliers': 12,
    'Revenus mensuels': 13,
    'Revenus Lave-Auto': 14,
    'Divers': 17,
    'Gratuités - mensuels': 20,
    'Salaires Stationnement': 29,
    'Uniformes': 32,
    'Entretien réparation - Nettoyage': 35,
    'Entretien réparation - Général': 36,
    'Entretien réparation - Equipement': 37,
    'Fourn. de stationnement': 41,
    'Frais de bureau': 49,
    'Télécommunication': 50,
    'Frais de cartes de crédit': 53,
    'Réclamations': 56,
    'Assurances Cautionnement': 57,
    'Taxes et permis': 58,
    'Honoraires de gestion': 63,
}

def safe_float(value):
    if value is None or value == "": return None
    if isinstance(value, (int, float)): return float(value)
    value = str(value).strip()
    is_negative = False
    if value.startswith('(') and value.endswith(')'):
        is_negative = True
        value = value[1:-1].strip()
    value = value.replace('$', '').strip()
    value = value.replace(" ", "").replace("\xa0", "").replace("\u202f", "")
    if value.startswith('-'):
        is_negative = True
        value = value[1:]
    elif value.endswith('-'):
        is_negative = True
        value = value[:-1]
    if ',' in value and '.' not in value:
        value = value.replace(',', '.')
    elif ',' in value and '.' in value:
        if value.rfind('.') > value.rfind(','):
            value = value.replace(',', '')
        else:
            value = value.replace('.', '').replace(',', '.')
    value = re.sub(r'[^\d\.\-]', '', value)
    try:
        result = float(value)
        return -result if is_negative else result
    except:
        return None

def extract_month_from_filename(file_obj):
    months = {
        'janvier': 'Janvier', 'février': 'Février', 'mars': 'Mars',
        'avril': 'Avril', 'mai': 'Mai', 'juin': 'Juin',
        'juillet': 'Juillet', 'août': 'Août', 'septembre': 'Septembre',
        'octobre': 'Octobre', 'novembre': 'Novembre', 'décembre': 'Décembre'
    }
    name = file_obj.name.lower() if hasattr(file_obj, 'name') else str(file_obj).lower()
    for key, value in months.items():
        if key in name: return value
    return None

def pdf_to_excel(pdf_path):
    """
    Convert P&L page from PDF to Excel using pandas.
    Extracts text from the P&L page and creates a clean Excel file.
    """
    doc = fitz.open(pdf_path)
    
    # Find P&L page
    pl_page = None
    for i in range(len(doc)):
        text = doc[i].get_text()
        if all(m.lower() in text.lower() for m in PAGE_MARKERS):
            pl_page = doc[i]
            break
    
    if pl_page is None:
        doc.close()
        return None, None
    
    # Get all text blocks sorted by position
    blocks = pl_page.get_text("blocks")
    blocks.sort(key=lambda b: (round(b[1], 0), b[0]))
    
    # Group blocks into rows (similar Y position)
    rows = []
    current_row = []
    current_y = None
    
    for block in blocks:
        y = round(block[1], 0)
        if current_y is None:
            current_y = y
        
        if abs(y - current_y) > 8:  # New row if Y differs by more than 8 points
            if current_row:
                rows.append(current_row)
            current_row = [block]
            current_y = y
        else:
            current_row.append(block)
    
    if current_row:
        rows.append(current_row)
    
    # Convert rows to table format (max 9 columns)
    table_data = []
    for row_blocks in rows:
        row_blocks.sort(key=lambda b: b[0])  # Sort by X
        row_texts = []
        for block in row_blocks[:9]:  # Max 9 columns
            text = block[4].strip()
            text = ' '.join(text.split())
            row_texts.append(text)
        # Pad to 9 columns
        while len(row_texts) < 9:
            row_texts.append('')
        table_data.append(row_texts)
    
    doc.close()
    
    # Create DataFrame
    columns = ['Account', 'Mois Courant', 'Budget', 'Écart', 'An Préc', 
               'Cumulatif', 'Cumul budget', 'Écart cumul', 'An Préc cumul']
    df = pd.DataFrame(table_data, columns=columns)
    
    # Save to Excel
    excel_path = pdf_path.replace('.pdf', '_converted.xlsx')
    df.to_excel(excel_path, index=False)
    
    return df, excel_path

def extract_from_converted_excel(df):
    """
    Read the converted Excel DataFrame.
    Find account names in Column A, get amount from Column B (Mois Courant).
    """
    data = {}
    
    for idx, row in df.iterrows():
        account_text = str(row['Account']).lower().strip()
        amount_text = str(row['Mois Courant']).strip()
        
        if not account_text or account_text == 'nan':
            continue
        
        # Try to match account name
        for pattern, standard_name in ACCOUNT_PATTERNS:
            if standard_name in data:
                continue
            if pattern in account_text:
                amount = safe_float(amount_text)
                if amount is not None:
                    data[standard_name] = amount
                break
    
    return data if len(data) >= 3 else None

def process_pdf_file(pdf_path):
    """Convert PDF to Excel and extract data"""
    df, excel_path = pdf_to_excel(pdf_path)
    if df is None:
        return None
    return extract_from_converted_excel(df)

def fill_template(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None:
        for sn in wb.sheetnames:
            if wb[sn].max_row >= 86:
                sheet_name = sn
                break
    if sheet_name is None: sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    updates = []
    total_cells = 0
    for month_name, pdf_data in all_data.items():
        col = MONTH_COLUMN.get(month_name)
        if col is None: continue
        month_cells = 0
        for pdf_account, amount in pdf_data.items():
            template_row = PDF_TO_TEMPLATE.get(pdf_account)
            if template_row is None: continue
            cell = ws.cell(row=template_row, column=col)
            cell.value = amount
            cell.number_format = '#,##0.00'
            month_cells += 1
            total_cells += 1
            col_letter = get_column_letter(col)
            updates.append(f"   ✅ {month_name} ({col_letter}{template_row}): ${amount:,.2f} - {pdf_account}")
        if month_cells > 0: updates.append(f"📊 {month_name}: {month_cells} cells filled")
    updates.append(f"\n📊 TOTAL: {total_cells} yellow cells filled (formulas auto-calculate)")
    return updates

def validate_template(wb, all_data):
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    if sheet_name is None: sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    results = []
    for month_name, pdf_data in all_data.items():
        col = MONTH_COLUMN.get(month_name)
        if col is None: continue
        benefice_net_pdf = pdf_data.get('BÉNÉFICE NET')
        revenus_nets_template = ws.cell(row=86, column=col).value
        if benefice_net_pdf is not None and revenus_nets_template is not None:
            diff = abs(benefice_net_pdf - revenus_nets_template)
            if diff > 0.01:
                results.append(f"⚠️ {month_name}: PDF BÉNÉFICE NET=${benefice_net_pdf:,.2f} ≠ Template REVENUS NETS=${revenus_nets_template:,.2f}")
            else:
                results.append(f"✅ {month_name}: BÉNÉFICE NET = REVENUS NETS = ${benefice_net_pdf:,.2f}")
        elif benefice_net_pdf is None:
            results.append(f"⚠️ {month_name}: BÉNÉFICE NET not extracted from PDF")
    return results

def get_parking_codes_from_pnl(file_obj):
    codes = []
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    updates = []
    all_data = {}
    try:
        updates.append("=" * 60)
        updates.append("🏗️ BUDGET SYSTEM - PDF→Excel→Template")
        updates.append("=" * 60)
        if isinstance(excel_file, bytes):
            wb = load_workbook(BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(BytesIO(excel_file.read()))
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        all_files = []
        if monthly_files_current: all_files.extend(monthly_files_current)
        if monthly_files_previous: all_files.extend(monthly_files_previous)
        if not all_files:
            updates.append("\n⚠️ No files!")
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return output.getvalue(), updates
        updates.append(f"\n📁 Processing {len(all_files)} files...")
        for file_obj in all_files:
            month = extract_month_from_filename(file_obj)
            file_obj.seek(0)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            try:
                data = process_pdf_file(tmp_path)
                if data and len(data) > 0:
                    all_data[month] = data
                    rev = data.get('TOTAL REVENUS', 'N/A')
                    net = data.get('BÉNÉFICE NET', 'N/A')
                    updates.append(f"   ✅ {month}: {len(data)} accounts | Rev: ${rev} | Net: ${net}")
                else:
                    updates.append(f"   ❌ {month}: No data extracted")
            finally:
                os.unlink(tmp_path)
        if all_data:
            updates.append(f"\n📝 Filling yellow cells ({len(all_data)} months)...")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
            updates.append(f"\n🔍 Validating...")
            validations = validate_template(wb, all_data)
            updates.extend(validations)
        else:
            updates.append("\n⚠️ No data extracted!")
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        updates.append("\n✅ WORKFLOW COMPLETE")
        return output.getvalue(), updates
    except Exception as e:
        updates.append(f"\n❌ ERROR: {str(e)}")
        import traceback
        updates.append(traceback.format_exc())
        return None, updates
