import pandas as pd
import pdfplumber

# daftar kategori inspeksi
kategori_list = [
    "IN SERVICE - STORAGE TANK INSPECTION",
    "INTERNAL INSPECTION",
    "OUT OF SERVICE - STORAGE TANK INSPECTION",
    "VISUAL TANK INSPECTION",
    "GENERAL VISUAL INSPECTION"
]

def cek_pdf(path):
    try:
        with pdfplumber.open(path) as pdf:
            # gabungkan semua teks untuk cek global
            all_text = " ".join(page.extract_text() or "" for page in pdf.pages)
            # teks halaman pertama
            first_page_text = pdf.pages[0].extract_text() or ""
            
            if "PT. SUCOFINDO" in all_text.upper():
                for kategori in kategori_list:
                    if kategori in first_page_text.upper():
                        return kategori
                return "PT. SUCOFINDO (tanpa kategori spesifik)"
            elif "BKI" in all_text.upper():
                return "BKI"
            elif "IRSINDO" in all_text.upper():
                return "IRSINDO"
            else:
                return ""
    except Exception as e:
        return f"Error: {e}"

# baca excel
df = pd.read_excel("D:\\Python\\STORAGE_TANK\\LIST-ST-SKIP.xlsx")

# cek setiap file
df["kategori"] = df["File"].apply(cek_pdf)

# simpan hasil
df.to_excel("D:\\Python\\STORAGE_TANK\\ST-SKIP-categorizer.xlsx", index=False)