# excel_fixer.py - ABSOLUTE MINIMUM
import re
import io
import fitz
import tempfile
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
    updates.append("✅ fix_excel called successfully")
    updates.append(f"   Current files: {len(monthly_files_current) if monthly_files_current else 0}")
    updates.append(f"   Previous files: {len(monthly_files_previous) if monthly_files_previous else 0}")
    
    try:
        if isinstance(excel_file, bytes):
            wb = load_workbook(io.BytesIO(excel_file))
        else:
            excel_file.seek(0)
            wb = load_workbook(io.BytesIO(excel_file.read()))
        
        updates.append(f"✅ Template loaded. Sheets: {wb.sheetnames}")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output.getvalue(), updates
        
    except Exception as e:
        updates.append(f"❌ ERROR: {str(e)}")
        return None, updates
