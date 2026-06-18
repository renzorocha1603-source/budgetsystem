# excel_fixer.py - CLEAN VERSION FROM SCRATCH
import re
import io
import fitz
import tempfile
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ============================================================================
# PDF ACCOUNT NAME → TEMPLATE ROW (YELLOW CELLS ONLY)
# ============================================================================
# Format: {french_name_pattern: template_row}
# Only maps to YELLOW (user input) rows - never formula rows
PDF_TO_TEMPLATE = {
    # REVENUES
    "revenus mensuels": 13,
    "revenus journaliers": 12,
    "revenus horaires": 12,
    "revenus lave-auto": 14,
    "divers": 17,
    "gratuités - mensuels": 20,
    "gratuités": 20,
    
    # LABOUR
    "salaires stationnement": 29,
    "salaire stationnement": 29,
    "uniformes": 32,
    
    # MAINTENANCE
    "nettoyage": 35,
    "entretien réparation - nettoyage": 35,
    "entretien réparation - equipement": 37,
    "entretien réparation - équipement": 37,
    "entretien réparation - général": 36,
    "entretien réparation - general": 36,
    "fourn. de stationnement": 41,
    "fournitures stationnement": 41,
    
    # OVERHEAD
    "frais de bureau": 49,
    "télécommunication": 50,
    "telecommunication": 50,
    "frais de cartes de crédit": 53,
    "frais de cartes de credit": 53,
    "réclamations": 56,
    "reclamations": 56,
    "assurances cautionnement": 57,
    "assurances": 57,
    "taxes et permis": 58,
    "honoraires de gestion": 63,
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
# CANADIAN FRENCH NUMBER PARSING
# ============================================================================
def parse_amount(text):
    """
    Parse Canadian French number formats:
    "43 585,46 $" -> 43585.46
    "(1 206,86) $" -> -1206.86
    "-1 206,86 $" -> -1206.86
    """
    if not text:
        return None
    
    text = str(text).strip()
    
    # Handle parentheses (negative)
    is_negative = False
    if text.startswith('(') and text.endswith(')'):
        is_negative = True
        text = text[1:-1].strip()
    
    # Remove $ signs
    text = text.replace('$', '').strip()
    
    # Remove ALL spaces (thousand separators)
    text = text.replace(' ', '').replace('\xa0', '').replace('\u202f', '')
    
    # Handle minus sign
    if text.startswith('-'):
        is_negative = True
        text = text[1:]
    
    # Replace comma with dot for decimal
    if ',' in text:
        text = text.replace(',', '.')
    
    # Remove any remaining non-numeric chars
    text = re.sub(r'[^\d\.\-]', '', text)
    
    try:
        value = float(text)
        return -value if is_negative else value
    except:
        return None

# ============================================================================
# PDF EXTRACTION - BY TEXT SEARCH (NO COORDINATES)
# ============================================================================
def extract_from_pdf(pdf_path):
    """
    Extract P&L data from PDF by searching text for account names.
    Grabs the first Canadian French number after each account name.
    Works on ANY page number - just searches all pages for the P&L content.
    """
    doc = fitz.open(pdf_path)
    data = {}
    
    # Combine text from all pages
    full_text = ""
    for page_num in range(len(doc)):
        full_text += doc[page_num].get_text("text") + "\n"
    doc.close()
    
    # Check if this is a P&L page (must have revenue accounts)
    if "revenus mensuels" not in full_text.lower() and "revenus journaliers" not in full_text.lower():
        return data
    
    # For each account we want to find
    for search_term, template_row in PDF_TO_TEMPLATE.items():
        # Find the search term in the text
        idx = full_text.lower().find(search_term)
        if idx == -1:
            continue
        
        # Get the text after the account name
        after_text = full_text[idx + len(search_term):]
        
        # Find the first Canadian French number pattern
        # Patterns: "43 585,46 $" or "(1 206,86) $" or "-1 206,46 $"
        patterns = [
            r'\(\s*(\d[\d\s]*,\d{2})\s*\)\s*\$?',   # (1 206,86) $
            r'-\s*(\d[\d\s]*,\d{2})\s*\$?',           # -1 206,46 $
            r'(\d[\d\s]*,\d{2})\s*\$?',                # 43 585,46 $
        ]
        
        for pattern in patterns:
            match = re.search(pattern, after_text[:200])  # Search within 200 chars
            if match:
                amount = parse_amount(match.group(1))
                if amount is not None and amount != 0:
                    data[template_row] = amount
                break
    
    return data

# ============================================================================
# OCR FALLBACK FOR IMAGE PDFS (JANUARY)
# ============================================================================
def extract_from_pdf_ocr(pdf_path):
    """OCR fallback for image-based PDFs"""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return {}
    
    doc = fitz.open(pdf_path)
    full_text = ""
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        try:
            text = pytesseract.image_to_string(img, lang='fra')
        except:
            try:
                text = pytesseract.image_to_string(img, lang='eng')
            except:
                text = pytesseract.image_to_string(img)
        
        full_text += text + "\n"
    
    doc.close()
    
    # Reuse the same extraction logic on OCR text
    data = {}
    for search_term, template_row in PDF_TO_TEMPLATE.items():
        idx = full_text.lower().find(search_term)
        if idx == -1:
            continue
        
        after_text = full_text[idx + len(search_term):]
        
        patterns = [
            r'\(\s*(\d[\d\s]*,\d{2})\s*\)\s*\$?',
            r'-\s*(\d[\d\s]*,\d{2})\s*\$?',
            r'(\d[\d\s]*,\d{2})\s*\$?',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, after_text[:200])
            if match:
                amount = parse_amount(match.group(1))
                if amount is not None and amount != 0:
                    data[template_row] = amount
                break
    
    return data

# ============================================================================
# TEMPLATE FILLING
# ============================================================================
def fill_template(wb, all_data):
    """Write extracted data to YELLOW cells only"""
    # Find the Donnees Historiques sheet
    sheet_name = None
    for sn in wb.sheetnames:
        if 'données' in sn.lower() or 'historique' in sn.lower():
            sheet_name = sn
            break
    
    if sheet_name is None:
        return ["❌ Sheet '2-Données historiques' not found"]
    
    ws = wb[sheet_name]
    updates = []
    total_cells = 0
    
    for month_en, month_data in all_data.items():
        col = MONTH_COLUMN.get(month_en)
        if col is None:
            updates.append(f"⚠️ Unknown month: {month_en}")
            continue
        
        col_letter = get_column_letter(col)
        month_cells = 0
        
        for template_row, amount in month_data.items():
            # Write to the yellow cell
            cell = ws.cell(row=template_row, column=col)
            cell.value = amount
            cell.number_format = '#,##0.00'
            month_cells += 1
            total_cells += 1
            
            updates.append(f"   ✅ {month_en} ({col_letter}{template_row}): ${amount:,.2f}")
        
        if month_cells > 0:
            updates.append(f"📊 {month_en}: {month_cells} cells filled")
    
    updates.append(f"\n📊 TOTAL: {total_cells} yellow cells filled (formulas auto-calculate)")
    return updates

# ============================================================================
# FUNCTIONS CALLED BY app.py
# ============================================================================
def get_parking_codes_from_pnl(file_obj):
    """Extract parking codes from filename or content"""
    codes = []
    if hasattr(file_obj, 'name'):
        matches = re.findall(r'(CMO\d+)', file_obj.name, re.IGNORECASE)
        codes.extend(matches)
    file_obj.seek(0)
    return list(set([c.upper() for c in codes]))

def fix_excel(excel_file, monthly_files_current=None, monthly_files_previous=None,
              budget_initial_file=None, fiche_stationnement_file=None,
              parking_code=None, word_data=None):
    """
    MAIN FUNCTION
    1. Extract data from each PDF by searching for account names
    2. Write values to YELLOW cells in template
    3. Formulas auto-calculate everything else
    """
    updates = []
    all_data = {}
    
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
        
        updates.append(f"✅ Template loaded: {parking_code or 'Unknown'}")
        
        # Collect all files
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
        
        updates.append(f"\n📁 Processing {len(all_files)} files...")
        
        for file_obj in all_files:
            # Get month from filename
            month_en = None
            if hasattr(file_obj, 'name'):
                name_lower = file_obj.name.lower()
                for fr, en in MONTH_MAP.items():
                    if fr in name_lower:
                        month_en = en
                        break
            
            if month_en is None:
                updates.append(f"⚠️ Could not determine month for {getattr(file_obj, 'name', 'unknown')}")
                continue
            
            # Save to temp file
            file_obj.seek(0)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(file_obj.read())
                tmp_path = tmp.name
            
            try:
                # Try digital extraction first
                data = extract_from_pdf(tmp_path)
                
                # If digital fails (less than 3 accounts), try OCR
                if len(data) < 3:
                    updates.append(f"   🔍 {month_en}: Digital extraction got {len(data)} accounts, trying OCR...")
                    ocr_data = extract_from_pdf_ocr(tmp_path)
                    if len(ocr_data) > len(data):
                        data = ocr_data
                
                if data:
                    all_data[month_en] = data
                    updates.append(f"   ✅ {month_en}: {len(data)} accounts extracted")
                else:
                    updates.append(f"   ❌ {month_en}: No data extracted")
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
        
        # Fill template
        if all_data:
            updates.append(f"\n📝 Filling YELLOW cells for {len(all_data)} months...")
            fill_updates = fill_template(wb, all_data)
            updates.extend(fill_updates)
        else:
            updates.append("\n⚠️ No data extracted from any file!")
        
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
