import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime

BATCH = "DATE-1"
EXCEL_PATH = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\BATCH_{BATCH}-LOCAL-DIR - Copy.xlsx"
OUTPUT_BASE = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\dates"
LOG_PATH = os.path.join(OUTPUT_BASE, f"log_batch_{BATCH}.txt")
COL_AITV = "AITV-EQ_ID"
COL_INSPECTION = "Inspection ID"
COL_FILE = "File"
os.makedirs(OUTPUT_BASE, exist_ok=True)

DATE_FORMATS = [
    "%b %d %Y", "%B %d %Y",
    "%b %d, %Y", "%B %d, %Y",   # <-- format dengan koma
    "%d %b %Y", "%d %B %Y",
    "%d-%b-%Y", "%d-%B-%Y",
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"
]



MONTHS = r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)"
DATE_REGEXES = [
    rf"\b{MONTHS}\s+\d{{1,2}}\s+\d{{4}}\b",
    rf"\b\d{{1,2}}\s+{MONTHS}\s+\d{{4}}\b",
    r"\b\d{4}-\d{2}-\d{2}\b",
    r"\b\d{1,2}/\d{1,2}/\d{4}\b",
    r"\b\d{1,2}-[A-Za-z]{3,9}-\d{4}\b",
    r"\b\d{1,2}\.\d{1,2}\.\d{4}\b",
]

def write_log(message):
    print(message)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(message + "\n")

def parse_date_str(s: str):
    s = re.sub(r"\s+", " ", s.strip())
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None

def extract_dates_from_text(text: str, keyword: str):
    dates = []
    if keyword.lower() in text.lower():
        # ambil substring setelah keyword
        parts = text.split(keyword)
        for part in parts[1:]:
            # ambil token setelah ":" kalau ada
            after_colon = part.split(":")[-1].strip()
            # pecah kata-kata
            tokens = after_colon.split()
            # coba gabungkan 3 kata pertama (misalnya "May 24, 2017")
            if len(tokens) >= 3:
                candidate = " ".join(tokens[:3])
                d = parse_date_str(candidate)
                if d:
                    dates.append(d)
            # fallback regex
            for rx in DATE_REGEXES:
                for m in re.finditer(rx, part):
                    d = parse_date_str(m.group(0))
                    if d:
                        dates.append(d)
    return dates

def extract_all_dates(pdf_path, keyword):
    all_dates = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            # cek text biasa
            all_dates.extend(extract_dates_from_text(text, keyword))
            # fallback: cek extract_words
            words = page.extract_words() or []
            joined = " ".join([w["text"] for w in words])
            all_dates.extend(extract_dates_from_text(joined, keyword))
    # unik + sort
    uniq = sorted({d.strftime("%Y-%m-%d"): d for d in all_dates}.values())
    return uniq

def extract_year_completed(pdf_path, keyword=r"Year\s*Completed"):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if keyword.lower() in text.lower():
                years = re.findall(r"\b(19\d{2}|20\d{2})\b", text)
                if years:
                    return years[0]
    return ""

def process():
    df = pd.read_excel(EXCEL_PATH)
    output_rows = []

    for _, row in df.iterrows():
        aitv = str(row[COL_AITV]).strip()
        ins = str(row[COL_INSPECTION]).strip()
        pdf_path = row[COL_FILE]
        write_log(f"\n>> {aitv} | {pdf_path}")

        insp_dates = extract_all_dates(pdf_path, r"Inspection\s*Date")
        comp_dates = extract_all_dates(pdf_path, r"Date\s*Completed")
        year_completed = extract_year_completed(pdf_path)

        write_log(f"   Inspection Dates found: {[d.strftime('%Y-%m-%d') for d in insp_dates]}")
        write_log(f"   Date Completed found: {[d.strftime('%Y-%m-%d') for d in comp_dates]}")

        if not insp_dates:
            write_log("   SKIP (no valid Inspection Date found)")
            continue

        insp_youngest = min(insp_dates)
        final_date = insp_youngest

        if comp_dates:
            same_year_comp = [d for d in comp_dates if d.year == insp_youngest.year]
            if same_year_comp:
                final_date = min(same_year_comp)
                write_log(f"   SAME YEAR → Date Completed used: {final_date.strftime('%Y-%m-%d')}")
            else:
                write_log(f"   DIFFERENT YEAR → use youngest Inspection Date: {insp_youngest.strftime('%Y-%m-%d')}")
        else:
            write_log(f"   NO Date Completed → use youngest Inspection Date: {insp_youngest.strftime('%Y-%m-%d')}")

        output_rows.append({
            "Inspection ID": ins,
            "AITV ID": aitv,
            "Inspection Date": final_date.strftime("%Y-%m-%d"),
            "Year": year_completed
        })

    outpath = os.path.join(OUTPUT_BASE, "summary.xlsx")
    pd.DataFrame(output_rows).to_excel(outpath, index=False)
    write_log(f"\n>> FINAL OUTPUT SAVED: {outpath}")

if __name__ == "__main__":
    process()