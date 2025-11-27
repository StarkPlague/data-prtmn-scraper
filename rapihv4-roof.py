import os
import re
import pdfplumber
import pandas as pd

BATCH = 42
EXCEL_PATH = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\BATCH_{BATCH}-LOCAL-DIR.xlsx"
OUTPUT_BASE = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\rapihv4-roof"
LOG_PATH = os.path.join(OUTPUT_BASE, f"log_batch_{BATCH}.txt")
KEYWORD_PATTERN = r"4.2.1.2"
COL_AITV = "AITV-EQ_ID"
COL_INSPECTION = "Inspection ID"
COL_FILE = "File"  
os.makedirs(OUTPUT_BASE, exist_ok=True)

# ------------------ Helpers ------------------
def write_log(message):
    print(message)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(message + "\n")

def is_toc_page(text: str) -> bool:
    """Heuristik sederhana untuk deteksi TOC."""
    if not text:
        return False
    return text.count(".....") > 3 or "table of contents" in text.lower()

def find_keyword_pages(pdf_path, keyword_pattern):
    """Cari semua halaman yang mengandung keyword, skip TOC."""
    found_pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if re.search(keyword_pattern, text, re.IGNORECASE):
                if not is_toc_page(text):
                    found_pages.append(i)
    return found_pages

def merge_header_rows(table_rows):
    """Gabungkan 2 baris header jadi 1."""
    if not table_rows:
        return None, []
    h1 = [str(c or "").strip() for c in table_rows[0]]
    if len(table_rows) > 1 and any(c and str(c).strip() for c in table_rows[1]):
        h2 = [str(c or "").strip() for c in table_rows[1]]
        merged = []
        mlen = max(len(h1), len(h2))
        h1 += [""] * (mlen - len(h1))
        h2 += [""] * (mlen - len(h2))
        for a, b in zip(h1, h2):
            merged.append(" ".join([x for x in [a, b] if x]).strip())
        return merged, table_rows[2:]
    return h1, table_rows[1:]

def clean_cols(columns):
    out = []
    for c in columns:
        c2 = (c or "").replace("\n", " ").replace("  ", " ").strip()
        cl = c2.lower()
        if "previous" in cl and "max" in cl: out.append("Prev Insp Max")
        elif "previous" in cl and "min" in cl: out.append("Prev Insp Min")
        elif "current" in cl and "max" in cl: out.append("Curr Insp Max")
        elif "current" in cl and "min" in cl: out.append("Curr Insp Min")
        else: out.append(c2)
    return out

def header_matches(header_text):
    """Cek apakah header cocok dengan tabel roof evaluation."""
    ht = header_text.lower()
    cues = [
        "thickness", "remaining service life", "corrosion rate",
        "nominal", "current inspection", "previous inspection",
        "period", "calculation corrosion rate", "remaining corrosion allowance"
    ]
    return sum(1 for k in cues if k in ht) >= 2

def extract_roof_tables(pdf_path, page_numbers):
    """Ekstrak semua tabel roof evaluation dari halaman-halaman tertentu."""
    all_rows = []
    header = None
    with pdfplumber.open(pdf_path) as pdf:
        for p in page_numbers:
            page = pdf.pages[p - 1]
            tbls = page.extract_tables() or []
            for t in tbls:
                if not t or len(t) < 2:
                    continue
                safe_rows = [[str(c).strip() if c else "" for c in r] for r in t]
                temp_header, body = merge_header_rows(safe_rows)
                header_text = " | ".join(temp_header)
                if not header_matches(header_text):
                    continue
                if header is None:
                    header = temp_header
                for row in body:
                    while len(row) < len(header):
                        row.append("")
                    all_rows.append(row[:len(header)])
    return pd.DataFrame(all_rows, columns=header) if header else None
# ------------------ Main ------------------

def process():
    df = pd.read_excel(EXCEL_PATH)
    for _, row in df.iterrows():
        aitv = str(row[COL_AITV]).strip()
        ins = str(row[COL_INSPECTION]).strip()
        pdf_path = row[COL_FILE]
        write_log(f"\n>> {aitv} | {pdf_path}")

        pages = find_keyword_pages(pdf_path, KEYWORD_PATTERN)
        if not pages:
            write_log("   SKIP (keyword not found)")
            continue

        table = extract_roof_tables(pdf_path, pages)
        if table is None or table.empty:
            write_log("   NO TABLE FOUND")
            continue

        table.columns = clean_cols(table.columns)
        table.insert(0, "Manufacturing Date", "")
        table.insert(0, "Inspection Date", "")
        table.insert(0, "Inspection ID", ins)
        table.insert(0, "AITV-EQ_ID", aitv)

        outpath = os.path.join(OUTPUT_BASE, f"{ins}.xlsx")
        table.to_excel(outpath, index=False)
        write_log(f"   SAVED: {outpath} | id {aitv}")

if __name__ == "__main__":
    process()