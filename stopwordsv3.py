import pandas as pd

# Load file Excel tanpa header
df = pd.read_excel('D:\\Python\\STORAGE_TANK\\BATCH_44\\output_course3\\ST121.xlsx', sheet_name='Sheet1', header=None)

# Keyword yang ingin dihapus persis di baris itu saja
keywords = ['MAX', 'MIN']

# Fungsi untuk cek apakah baris mengandung keyword (tanpa menyentuh baris setelahnya)
def is_exact_keyword_row(row):
    str_row = row.astype(str)
    return any(kw in cell for cell in str_row for kw in keywords)

# Bangun ulang DataFrame: hapus hanya baris yang mengandung keyword, sisanya tetap
df_cleaned = df[[not is_exact_keyword_row(row) for _, row in df.iterrows()]].reset_index(drop=True)

# Simpan ulang
df_cleaned.to_excel('D:\\Python\\STORAGE_TANK\\BATCH_44\\output_course3\\ST121_cleaned10.xlsx', index=False, header=False)