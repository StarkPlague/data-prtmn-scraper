import os
import pandas as pd

# Folder input & output
BATCH = 42
INPUT_DIR = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\rapihv2-body"
OUTPUT_FILE = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\merged_body.xlsx"

# Keyword yang ingin dihapus persis di baris itu saja
KEYWORDS = ['MAX', 'MIN', 'LONG', 'UT', 'EXTERNAL']

# Fungsi untuk cek apakah baris mengandung keyword
def is_exact_keyword_row(row):
    str_row = row.astype(str)
    return any(kw in cell for cell in str_row for kw in KEYWORDS)

def clean_and_merge_excels(input_dir, output_file):
    all_dfs = []
    for fname in os.listdir(input_dir):
        if fname.lower().endswith(".xlsx"):
            fpath = os.path.join(input_dir, fname)
            try:
                # Load tanpa header
                df = pd.read_excel(fpath, sheet_name=0, header=None)

                # Filter baris
                mask = [not is_exact_keyword_row(row) for _, row in df.iterrows()]
                df_cleaned = df[mask].reset_index(drop=True)

                # Tambahkan kolom nama file asal (tanpa ekstensi)
                base_name = os.path.splitext(fname)[0]
                df_cleaned.insert(0, "Source_File", base_name)

                all_dfs.append(df_cleaned)
                print(f"Loaded & cleaned: {fname} ({df_cleaned.shape[0]} rows)")
            except Exception as e:
                print(f"Skip {fname}: {e}")

    if not all_dfs:
        print("No Excel files found.")
        return

    merged = pd.concat(all_dfs, ignore_index=True)
    merged.to_excel(output_file, index=False, header=False)
    print(f"\n✅ Merged {len(all_dfs)} cleaned files into: {output_file}")
    print(f"Total rows: {merged.shape[0]}")

if __name__ == "__main__":
    clean_and_merge_excels(INPUT_DIR, OUTPUT_FILE)