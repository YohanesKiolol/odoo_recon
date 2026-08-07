import csv
import glob
from pathlib import Path
from datetime import datetime
from decimal import Decimal
from config import MUTATION_DIR, BANK_ACCOUNTS

def read_all_mutations():
    """
    Reads all mutation files across all banks and aliases.
    Returns:
        mutations: list of dicts with recognized transactions
        unknowns: list of dicts with unmapped transactions
    """
    mutations = []
    unknowns = []

    # 1. BCA
    bca_dir = MUTATION_DIR / "bca"
    if bca_dir.exists():
        for alias in bca_dir.iterdir():
            if alias.is_dir():
                for f in alias.glob("*.csv"):
                    m, u = read_mutation_bca(f, alias.name)
                    mutations.extend(m)
                    unknowns.extend(u)

    # 2. Mandiri
    mandiri_dir = MUTATION_DIR / "mandiri"
    if mandiri_dir.exists():
        for alias in mandiri_dir.iterdir():
            if alias.is_dir():
                for f in alias.glob("*.csv"):
                    m, u = read_mutation_mandiri(f, alias.name)
                    mutations.extend(m)
                    unknowns.extend(u)

    # 3. BRI
    bri_dir = MUTATION_DIR / "bri"
    if bri_dir.exists():
        for alias in bri_dir.iterdir():
            if alias.is_dir():
                for f in alias.glob("*.csv"):
                    m, u = read_mutation_bri(f, alias.name)
                    mutations.extend(m)
                    unknowns.extend(u)

    return mutations, unknowns


def _clean_amount(val: str) -> float:
    if not val:
        return 0.0
    val = val.replace("Rp", "").replace(" ", "").replace(",", "")
    if val == ".00":
        return 0.0
    try:
        return float(val)
    except ValueError:
        return 0.0


def read_mutation_bca(filepath: Path, alias: str):
    mutations = []
    unknowns = []
    
    # Fallback year from "Periode" line or current year
    year = datetime.today().year
    
    with open(filepath, "r", encoding="utf-8-sig", errors="replace") as f:
        reader = csv.reader(f, delimiter=",")
        
        # 1. Dynamically read metadata until header row is found
        for row in reader:
            if not row:
                continue
            first_cell = row[0].strip()
            if "Periode" in first_cell:
                parts = first_cell.split("/")
                if len(parts) >= 3:
                    try:
                        year = int(parts[2].split(" ")[0][:4])
                    except Exception:
                        pass
            # Header row starts with "Tanggal" or "Date"
            if "Tanggal" in first_cell or "Date" in first_cell:
                break
                
        # 2. Read transaction data rows
        for row in reader:
            if len(row) < 4:
                continue
                
            date_str = row[0].strip()
            desc = row[1].strip()
            
            # Amount with CR/DB indicator may be in column 3 or column 4
            jumlah = row[3].strip() if len(row) > 3 else ""
            if not (jumlah.endswith(" CR") or jumlah.endswith(" DB")):
                if len(row) > 4 and (row[4].strip().endswith(" CR") or row[4].strip().endswith(" DB")):
                    jumlah = row[4].strip()
                else:
                    continue
            
            if jumlah.endswith(" CR"):
                cr_db_type = "Credit"
                amount = _clean_amount(jumlah.replace(" CR", ""))
            elif jumlah.endswith(" DB"):
                cr_db_type = "Debit"
                amount = _clean_amount(jumlah.replace(" DB", ""))
            else:
                continue
            
            # Parse Date (BCA format can be "DD/MM/YYYY", "YYYY-MM-DD", or "DD/MM")
            dt = None
            if len(date_str) == 10 and date_str.count("/") == 2:
                try:
                    dt = datetime.strptime(date_str, "%d/%m/%Y")
                except ValueError:
                    pass
            if not dt and len(date_str) == 10 and date_str.count("-") == 2:
                try:
                    dt = datetime.strptime(date_str, "%Y-%m-%d")
                except ValueError:
                    pass
            if not dt:
                try:
                    dt = datetime.strptime(f"{date_str}/{year}", "%d/%m/%Y")
                except ValueError:
                    continue

            # Determine Type
            desc_upper = desc.upper()
            
            # MID Validation
            txn_type = "Unknown"
            expected_mid = BANK_ACCOUNTS.get("bca", {}).get(alias, {}).get("mid", "")
            if expected_mid and expected_mid.upper() not in desc_upper:
                # If a MID is configured and it's not in the description, this is not an EDC transaction
                pass
            else:
                # EDC transactions are strictly Credit
                if cr_db_type == "Credit":
                    if "QR" in desc_upper:
                        txn_type = "QR"
                    elif "KARTU KREDIT" in desc_upper:
                        txn_type = "Credit Card"
                    elif "KR OTOMATIS" in desc_upper:
                        txn_type = "Debit Card"
                
            # Extract Admin Fee from description
            admin_fee = Decimal("0")
            if cr_db_type == "Credit":
                import re
                # Match DDR: <amount> or ADM: <amount> (amount can have commas or spaces)
                m = re.search(r'(?:DDR|ADM)\s*:\s*([\d\.,]+)', desc_upper)
                if m:
                    fee_str = m.group(1).replace(",", "")
                    try:
                        admin_fee = Decimal(fee_str)
                    except Exception:
                        pass
                        
            record = {
                "bank": "bca",
                "alias": alias,
                "date": dt,
                "desc": desc,
                "amount": amount,
                "admin_fee": admin_fee,
                "type": txn_type,
                "cr_db_type": cr_db_type
            }
            
            if txn_type == "Unknown":
                unknowns.append(record)
            else:
                mutations.append(record)

    return mutations, unknowns



def _read_mutation_mandiri_raw(filepath: Path, alias: str):
    mutations = []
    unknowns = []
    
    with open(filepath, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        
        if not header:
            return mutations, unknowns
            
        try:
            date_idx = header.index("Date")
            desc1_idx = header.index("Description")
            desc2_idx = header.index("Description", desc1_idx + 1)
            credit_idx = header.index("Credit")
            debit_idx = header.index("Debit")
        except ValueError:
            return mutations, unknowns
            
        for row in reader:
            if len(row) <= credit_idx:
                continue
                
            credit_str = row[credit_idx].strip()
            debit_str = row[debit_idx].strip() if len(row) > debit_idx else ""
            
            amount = 0
            cr_db_type = "Unknown"
            
            if credit_str and credit_str != ".00":
                amount = _clean_amount(credit_str)
                cr_db_type = "Credit"
            elif debit_str and debit_str != ".00":
                amount = _clean_amount(debit_str)
                cr_db_type = "Debit"
                
            if amount == 0:
                continue
                
            date_str = row[date_idx].strip()
            try:
                dt = datetime.strptime(date_str, "%d/%m/%y")
            except ValueError:
                continue
                
            desc1 = row[desc1_idx].strip()
            desc2 = row[desc2_idx].strip() if len(row) > desc2_idx else ""
            desc = f"{desc1} | {desc2}"
            desc_upper = desc.upper()
            
            # MID Validation
            txn_type = "Unknown"
            expected_mid = BANK_ACCOUNTS.get("mandiri", {}).get(alias, {}).get("mid", "")
            if expected_mid and expected_mid.upper() not in desc_upper:
                # If a MID is configured and it's not in the description, this is not an EDC transaction
                pass
            else:
                if cr_db_type == "Credit":
                    if "QREYERIZZ" in desc_upper or "QR" in desc_upper:
                        txn_type = "QR"
                    elif "BDN" in desc_upper:
                        txn_type = "Debit Card"
                
            record = {
                "bank": "mandiri",
                "alias": alias,
                "date": dt,
                "desc": desc,
                "amount": amount,
                "type": txn_type,
                "cr_db_type": cr_db_type
            }
            
            if txn_type == "Unknown":
                unknowns.append(record)
            else:
                mutations.append(record)

    return mutations, unknowns

def read_mutation_mandiri(filepath: Path, alias: str):
    return _read_mutation_mandiri_raw(filepath, alias)

def read_mutation_bri(filepath: Path, alias: str):
    mutations = []
    unknowns = []
    
    with open(filepath, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            credit_str = row.get("MUTASI_KREDIT", ".00").strip()
            debit_str = row.get("MUTASI_DEBET", ".00").strip()
            
            amount = 0
            cr_db_type = "Unknown"
            
            if credit_str and credit_str != ".00":
                amount = _clean_amount(credit_str)
                cr_db_type = "Credit"
            elif debit_str and debit_str != ".00":
                amount = _clean_amount(debit_str)
                cr_db_type = "Debit"
                
            if amount == 0:
                continue
                
            date_str = row.get("TGL_TRAN", "").strip()
            # BRI Date format: "2026-07-01 03:22:44"
            try:
                dt = datetime.strptime(date_str.split(" ")[0], "%Y-%m-%d")
            except ValueError:
                continue
                
            desc = row.get("DESK_TRAN", "").strip()
            desc_upper = desc.upper()
            
            # MID Validation
            txn_type = "Unknown"
            expected_mid = BANK_ACCOUNTS.get("bri", {}).get(alias, {}).get("mid", "")
            if expected_mid and expected_mid.upper() not in desc_upper:
                # If a MID is configured and it's not in the description, this is not an EDC transaction
                pass
            else:
                if cr_db_type == "Credit":
                    if "QRIS" in desc_upper:
                        txn_type = "QR"
                    elif "OFFUS" in desc_upper or "ONUS" in desc_upper:
                        txn_type = "Debit Card"
                
            # Extract Admin Fee from MDR in description
            admin_fee = Decimal("0")
            if cr_db_type == "Credit":
                import re
                m = re.search(r'MDR:([\d.,]+)', desc_upper)
                if m:
                    fee_str = m.group(1).replace(".", "").replace(",", ".")
                    try:
                        admin_fee = Decimal(fee_str)
                    except:
                        pass
                    
            record = {
                "bank": "bri",
                "alias": alias,
                "date": dt,
                "desc": desc,
                "amount": amount,
                "admin_fee": admin_fee,
                "type": txn_type,
                "cr_db_type": cr_db_type
            }
            
            if txn_type == "Unknown":
                unknowns.append(record)
            else:
                mutations.append(record)

    return mutations, unknowns
