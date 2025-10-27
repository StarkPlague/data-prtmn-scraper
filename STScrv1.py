import pandas as pd
import os
import re
import gdown
import time

BATCH = 5
input_excel = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\BATCH_{BATCH}-SOURCE.xlsx"
output_excel = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\BATCH_{BATCH}-LOCAL-DIR.xlsx"
download_folder = rf"D:\Python\STORAGE_TANK\BATCH_{BATCH}\files"
max_retry = 1

os.makedirs(download_folder, exist_ok=True)

df = pd.read_excel(input_excel)

def download_from_drive(link, folder):
    if pd.isna(link):
        return None

    match = re.search(r"/d/([a-zA-Z0-9_-]+)", str(link))
    if not match:
        return None

    file_id = match.group(1)
    output_path = os.path.join(folder, f"{file_id}.pdf")
    url = f"https://drive.google.com/uc?id={file_id}"

    try:
        gdown.download(url, output_path, quiet=True, use_cookies=False)
        return output_path if os.path.exists(output_path) and os.path.getsize(output_path) > 1000 else None
    except:
        return None

failed = []
print("🚀 Mulai download...")

for i, link in enumerate(df['File']):
    local_path = download_from_drive(link, download_folder)
    if local_path:
        df.at[i, 'File'] = local_path
        print(f"[OK] {i+1}")
    else:
        failed.append((i, link))
        print(f"[FAIL] {i+1}")

attempt = 1
while failed and attempt <= max_retry:
    print(f"\n🔁 Retry ke-{attempt}")
    new_fail = []

    for i, link in failed:
        time.sleep(2)
        local_path = download_from_drive(link, download_folder)
        if local_path:
            df.at[i, 'File'] = local_path
            print(f"[Retry OK] index {i}")
        else:
            new_fail.append((i, link))
            print(f"[Retry FAIL] index {i}")
    failed = new_fail
    attempt += 1

# --- PAKSA SIMPAN HASIL AKHIR ---
df.to_excel(output_excel, index=False)
print("\n✅ DONE")
print(f"📁 Output saved: {output_excel}")
print(f"✅ Success: {len(df)-len(failed)}")
print(f"❌ Fail: {len(failed)}")
