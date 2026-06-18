# diagnose_converter.py - Run locally to see exact word positions
import fitz
import os
import sys

def diagnose(pdf_path):
    print("=" * 60)
    print(f"DIAGNOSING: {os.path.basename(pdf_path)}")
    print("=" * 60)
    
    doc = fitz.open(pdf_path)
    print(f"Total pages: {len(doc)}")
    
    # Find P&L page
    pl_page = None
    pl_page_num = None
    
    for i in range(len(doc)):
        text = doc[i].get_text()
        if "TOTAL REVENUS" in text and "BÉNÉFICE NET" in text:
            pl_page = doc[i]
            pl_page_num = i
            print(f"\n✅ P&L found on page {i+1}")
            break
    
    if pl_page is None:
        print("❌ P&L page not found!")
        for i in range(len(doc)):
            text = doc[i].get_text()
            if "revenus" in text.lower():
                print(f"\nPage {i+1} has 'revenus':")
                print(text[:500])
        doc.close()
        return
    
    # Get words with coordinates
    words = pl_page.get_text("words")
    print(f"\nTotal words on page: {len(words)}")
    
    # Find x-coordinate range
    all_x = [w[0] for w in words]
    min_x = min(all_x)
    max_x = max(all_x)
    print(f"X range: {min_x:.0f} to {max_x:.0f}")
    print(f"Page width: {pl_page.rect.width:.0f}")
    
    # Calculate approximate column boundaries (9 columns)
    col_width = (max_x - min_x) / 9
    print(f"\nEstimated column width: {col_width:.0f} points")
    
    # Group words by row (y position rounded to nearest 10)
    print(f"\n{'='*60}")
    print("ALL ROWS (y rounded to 10, showing first 2 columns)")
    print(f"{'='*60}")
    
    rows = {}
    for w in words:
        y_key = round(w[1] / 10) * 10
        if y_key not in rows:
            rows[y_key] = []
        rows[y_key].append(w)
    
    for y_key in sorted(rows.keys()):
        row_words = sorted(rows[y_key], key=lambda w: w[0])
        
        # Get text from column 0 (account name)
        col0_text = ' '.join([w[4] for w in row_words if w[0] < min_x + col_width])
        
        # Get text from column 1 (Mois Courant)
        col1_text = ' '.join([w[4] for w in row_words if min_x + col_width <= w[0] < min_x + 2*col_width])
        
        if col0_text.strip() or col1_text.strip():
            print(f"y={y_key:.0f}: [{col0_text[:50]:50s}] [{col1_text[:20]:20s}]")
    
    doc.close()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        diagnose(sys.argv[1])
    else:
        # Try to find February PDF
        test_file = "02-CMO111 Rapport mensuel de gestion février 2026.pdf"
        if os.path.exists(test_file):
            diagnose(test_file)
        else:
            print("Usage: python diagnose_converter.py <pdf_file>")
            print(f"File not found: {test_file}")
