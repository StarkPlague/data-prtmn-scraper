import os
import pdfplumber
import pandas as pd
from datetime import datetime

BATCH = 44
EXCEL_PATH = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\BATCH_{BATCH}-LOCAL-DIR.xlsx"
OUTPUT_BASE = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\output_course3"
KEYWORD = "RCA/2N"
COL_AITV = "AITV-EQ_ID"
COL_INSPECTION = "Inspection ID"
COL_FILE = "File"
os.makedirs(OUTPUT_BASE, exist_ok=True)

def find_keyword_page(pdf_path, keyword):
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if keyword.lower() in text.lower():
                return i
    return None

def merge_header_rows(table_rows):
    header = table_rows[0]
    if len(table_rows) > 1 and any(table_rows[1]):  
        second = table_rows[1]
        merged = []
        for a, b in zip(header, second):
            if a and b:
                merged.append(f"{a} {b}".strip())
            else:
                merged.append(a or b)
        return merged, table_rows[2:]
    return header, table_rows[1:]

def extract_course_table(pdf_path, start_page):
    all_rows = []
    header = None
    with pdfplumber.open(pdf_path) as pdf:
        for p in range(start_page, len(pdf.pages)+1):
            tbls = pdf.pages[p-1].extract_tables()
            for t in tbls:
                if len(t) > 1 and "course" in (t[0][0] or "").lower():
                    safe_rows = []
                    for r in t:
                        safe_row = []
                        for c in r:
                            safe_row.append(str(c).strip() if c is not None else "")
                        safe_rows.append(safe_row)
                    temp_header, body = merge_header_rows(safe_rows)
                    if header is None:
                        header = temp_header
                    for row in body:
                        while len(row) < len(header):
                            row.append("")
                        all_rows.append(row[:len(header)])
    return pd.DataFrame(all_rows, columns=header) if header else None

def clean_cols(columns):
    clean = []
    for c in columns:
        c2 = c.replace("\n"," ").replace("  "," ").strip()
        if "previous" in c2.lower() and "max" in c2.lower(): clean.append("Prev Insp Max")
        elif "previous" in c2.lower() and "min" in c2.lower(): clean.append("Prev Insp Min")
        elif "current" in c2.lower() and "max" in c2.lower(): clean.append("Curr Insp Max")
        elif "current" in c2.lower() and "min" in c2.lower(): clean.append("Curr Insp Min")
        else: clean.append(c2)
    return clean

def process():
    df = pd.read_excel(EXCEL_PATH)
    for _, row in df.iterrows():
        aitv = str(row[COL_AITV]).strip()
        ins = str(row[COL_INSPECTION]).strip()
        pdf_path = row[COL_FILE]
        print(f"\n>> {aitv} | {pdf_path}")

        page = find_keyword_page(pdf_path, KEYWORD)
        if not page:
            print("   SKIP (keyword not found)")
            continue

        table = extract_course_table(pdf_path, page)
        if table is None or table.empty:
            print("   NO TABLE FOUND")
            continue

        table.columns = clean_cols(table.columns)
        table.insert(0, "Manufacturing Date", "")
        table.insert(0, "Inspection Date", "")
        table.insert(0, "Inspection ID", ins)
        table.insert(0, "AITV-EQ_ID", aitv)

        outpath = os.path.join(OUTPUT_BASE, f"{aitv}.xlsx")
        table.to_excel(outpath, index=False)
        print(f"   SAVED: {outpath}")

if __name__ == "__main__":
    process()
