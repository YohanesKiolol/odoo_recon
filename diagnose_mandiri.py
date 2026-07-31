"""Inspect Mandiri CSV columns. Run: source .venv/bin/activate && python3 diagnose_mandiri.py"""
import pyzipper, csv
from pathlib import Path
from dotenv import load_dotenv
import os

load_dotenv()

zip_dir = Path(os.getenv("MANDIRI_ZIP_DIR", "input/mandiri"))
pwd = os.getenv("MANDIRI_ZIP_PASSWORD", "").encode()
num_col = os.getenv("MANDIRI_NUMBER_COLUMN", "Reff ID/Invoice No")

print(f"Looking for number column: {num_col!r}")

for zip_path in sorted(zip_dir.glob("*.zip"))[:1]:
    print(f"\nZIP: {zip_path.name}")
    with pyzipper.AESZipFile(zip_path, "r") as zf:
        zf.setpassword(pwd)
        for name in zf.namelist():
            if not name.lower().endswith(".csv"):
                continue
            data = zf.read(name).decode("utf-8", errors="replace")
            lines = data.splitlines()

            print(f"CSV rows 1-7 (raw):")
            for i, line in enumerate(lines[:7], 1):
                print(f"  Row {i}: {line[:130]}")

            if len(lines) >= 6:
                header_line = lines[5]
                data_line   = lines[6] if len(lines) > 6 else ""
                reader = csv.DictReader([header_line, data_line], skipinitialspace=True)
                for row in reader:
                    print(f"\nAll column names (exact):")
                    for k in row.keys():
                        if k is not None:
                            marker = " ← MATCH" if k.strip().upper() == num_col.strip().upper() else ""
                            print(f"  {k.strip()!r}{marker}")
                    row_clean = {k.strip(): v.strip() for k, v in row.items() if k}
                    print(f"\nFirst data row values (non-empty):")
                    for k, v in row_clean.items():
                        if v: print(f"  {k!r}: {v!r}")
                    break
            break
