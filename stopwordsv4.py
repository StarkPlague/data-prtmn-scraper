import os
import pandas as pd

# Folder input & output
BATCH = 6
INPUT_DIR = rF"D:\Python\STORAGE_TANK\BATCH_{BATCH}\rapihv2-body"
OUTPUT_DIR = rF"D:\Python\STORAGE_TANK\BATCH_{BATCH}\stopwordsv4-body"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Keyword yang ingin dihapus persis di baris itu saja
keywords = ['MAX', 'MIN', 'LONG', 'UT', 'EXTERNAL']

# Fungsi untuk cek apakah baris mengandung keyword
def is_exact_keyword_row(row):
    str_row = row.astype(str)
    return any(kw in cell for cell in str_row for kw in keywords)

def clean_excel_file(file_path, output_dir):
    # Load file tanpa header
    df = pd.read_excel(file_path, sheet_name=0, header=None)

    # Filter baris
    mask = [not is_exact_keyword_row(row) for _, row in df.iterrows()]
    df_cleaned = df[mask].reset_index(drop=True)

    # Nama file output
    fname = os.path.basename(file_path)
    name, ext = os.path.splitext(fname)
    outpath = os.path.join(output_dir, f"{name}_cleaned{ext}")

    # Simpan ulang
    df_cleaned.to_excel(outpath, index=False, header=False)
    print(f"Saved cleaned file: {outpath}")

def process_all_excels(input_dir, output_dir):
    for fname in os.listdir(input_dir):
        if fname.lower().endswith(".xlsx"):
            fpath = os.path.join(input_dir, fname)
            try:
                clean_excel_file(fpath, output_dir)
            except Exception as e:
                print(f"Skip {fname}: {e}")

if __name__ == "__main__":
    process_all_excels(INPUT_DIR, OUTPUT_DIR)