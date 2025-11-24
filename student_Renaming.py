import os
import re
import io
import requests
import zipfile
import pandas as pd

# ---------- Config ----------
EXCEL_URL = "https://github.com/amitgh88/StudentNameTransfer/raw/main/Guaridan_Call_Nov2025.xlsx"
ZIP_URL   = "https://github.com/amitgh88/StudentNameTransfer/raw/main/latter_all.zip"
EXCEL_FILE = "Guaridan_Call_Nov2025.xlsx"
ZIP_FILE   = "latter_all.zip"
EXTRACT_DIR = "letters"

# ---------- Helpers ----------
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]

def to_clean_str(x):
    # Convert numbers like 27900124001.0 → '27900124001'; keep strings intact; strip spaces
    if pd.isna(x):
        return None
    if isinstance(x, (int,)):
        return str(x)
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        else:
            return str(x).strip()
    return str(x).strip()

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

# ---------- Step 1: Download Excel ----------
print("Downloading Excel...")
resp = requests.get(EXCEL_URL)
resp.raise_for_status()
with open(EXCEL_FILE, "wb") as f:
    f.write(resp.content)
print(f"Saved: {EXCEL_FILE}")

# ---------- Step 2: Read registration numbers B2:B44 ----------
print("Reading registration numbers from B2:B44...")
df = pd.read_excel(EXCEL_FILE, engine="openpyxl", usecols="B", header=None)
# df rows start at 1 for Excel row 1; we want rows 2..44 → index 1..43
regs_raw = df.iloc[1:44, 0].tolist()
registration_numbers = [to_clean_str(x) for x in regs_raw if to_clean_str(x)]

print(f"Found {len(registration_numbers)} registration numbers:")
for i, rn in enumerate(registration_numbers, start=2):
    print(f"  B{i}: {rn}")

# ---------- Step 3: Download ZIP ----------
print("\nDownloading ZIP...")
resp = requests.get(ZIP_URL)
resp.raise_for_status()
with open(ZIP_FILE, "wb") as f:
    f.write(resp.content)
print(f"Saved: {ZIP_FILE}")

# ---------- Step 4: Inspect ZIP contents ----------
with zipfile.ZipFile(ZIP_FILE, "r") as zf:
    names = zf.namelist()
    print("\nZIP contents:")
    for n in names:
        print(f"  {n}")

# ---------- Step 5: Extract PDFs (preserve structure) ----------
print("\nExtracting ZIP to:", EXTRACT_DIR)
ensure_dir(EXTRACT_DIR)
with zipfile.ZipFile(ZIP_FILE, "r") as zf:
    zf.extractall(EXTRACT_DIR)

# Walk the extract dir to find PDFs and match pattern anywhere
print("\nScanning extracted PDFs...")
pdf_paths = []
pattern = re.compile(r"latter_all_(\d+)\.pdf$", re.IGNORECASE)
for root, dirs, files in os.walk(EXTRACT_DIR):
    for fname in files:
        if fname.lower().endswith(".pdf"):
            full = os.path.join(root, fname)
            pdf_paths.append(full)

print(f"Total PDFs found: {len(pdf_paths)}")
for p in pdf_paths:
    print("  ", p)

# Filter those that match expected naming
matched = []
for p in pdf_paths:
    fname = os.path.basename(p)
    m = pattern.search(fname)
    if m:
        num = int(m.group(1))
        matched.append((p, num))

# Sort by the numeric suffix (01..43)
matched.sort(key=lambda x: x[1])

print("\nMatched PDFs (ordered by suffix):")
for p, num in matched:
    print(f"  {os.path.relpath(p)}  → suffix {num}")

# ---------- Step 6: Validate counts ----------
if len(matched) != len(registration_numbers):
    print(f"\nERROR: Count mismatch — {len(matched)} matched PDFs vs {len(registration_numbers)} registration numbers.")
    print("Please verify the Excel range B2:B44 and the ZIP contents.")
    print("No renaming performed.")
else:
    # ---------- Step 7: Preview mapping ----------
    print("\nPreview mapping (old → new):")
    preview_map = []
    for (old_path, _), reg_no in zip(matched, registration_numbers):
        new_path = os.path.join(os.path.dirname(old_path), f"{reg_no}.pdf")
        preview_map.append((old_path, new_path))
        print(f"  {os.path.relpath(old_path)} → {os.path.relpath(new_path)}")

    # ---------- Step 8: Confirm and rename ----------
    proceed = input("\nProceed with renaming? (yes/no): ").strip().lower()
    if proceed == "yes":
        errors = 0
        for old_path, new_path in preview_map:
            try:
                # If target exists, replace it
                os.replace(old_path, new_path)
            except Exception as e:
                print(f"Failed: {old_path} → {new_path} ({e})")
                errors += 1

        # ---------- Step 9: Verify ----------
        print("\nVerification:")
        ok = 0
        for _, new_path in preview_map:
            if os.path.exists(new_path):
                print(f"  OK: {os.path.relpath(new_path)}")
                ok += 1
            else:
                print(f"  MISSING: {os.path.relpath(new_path)}")

        if errors == 0 and ok == len(preview_map):
            print("\nRenaming complete ✅")
        else:
            print(f"\nRenaming finished with {errors} errors; {ok}/{len(preview_map)} files verified.")

    else:
        print("\nRenaming cancelled.")
