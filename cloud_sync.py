from __future__ import annotations

import json
import os
import socket
import hmac
import hashlib
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime, date, timedelta

from pathlib import Path
from decimal import Decimal

# ── Encryption ────────────────────────────────────────────────────────────────
try:
    from cryptography.fernet import Fernet, InvalidToken
    _HAS_FERNET = True
except ImportError:
    _HAS_FERNET = False

# Hardcoded fallback for sales portal (no .env on sales laptops).
# anon key alone gives ZERO data access — RLS blocks everything.
# Only useful for Supabase Auth login.
_FALLBACK_SUPABASE_URL = "https://drauzroctexgfqxouswt.supabase.co"
_FALLBACK_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRyYXV6cm9jdGV4Z2ZxeG91c3d0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODY3MjU4NzUsImV4cCI6MjEwMjMwMTg3NX0.DLQ7WKb4ykshyjS1a_dHggQCM3vFpRDjxf0SXB77-qc"


def get_supabase_url() -> str:
    try:
        import config
        return getattr(config, "SUPABASE_URL", "") or os.environ.get("SUPABASE_URL", "") or _FALLBACK_SUPABASE_URL
    except Exception:
        return os.environ.get("SUPABASE_URL", "") or _FALLBACK_SUPABASE_URL


def get_supabase_key() -> str:
    """Return best available key: SERVICE_ROLE > SUPABASE_KEY > anon key > fallback."""
    try:
        import config
        return (
            getattr(config, "SERVICE_ROLE", "")
            or getattr(config, "SUPABASE_KEY", "")
            or os.environ.get("SERVICE_ROLE", "")
            or os.environ.get("SUPABASE_KEY", "")
            or _FALLBACK_ANON_KEY
        )
    except Exception:
        return os.environ.get("SERVICE_ROLE", "") or os.environ.get("SUPABASE_KEY", "") or _FALLBACK_ANON_KEY


def get_anon_key() -> str:
    """Return anon key specifically (for Supabase Auth login)."""
    try:
        import config
        return getattr(config, "SUPABASE_ANON_KEY", "") or os.environ.get("SUPABASE_ANON_KEY", "") or _FALLBACK_ANON_KEY
    except Exception:
        return os.environ.get("SUPABASE_ANON_KEY", "") or _FALLBACK_ANON_KEY


def get_encryption_key() -> str:
    """Return Fernet encryption key from config/.env. Empty string if not set."""
    try:
        import config
        return getattr(config, "ENCRYPTION_KEY", "") or os.environ.get("ENCRYPTION_KEY", "")
    except Exception:
        return os.environ.get("ENCRYPTION_KEY", "")


def get_company_key() -> str:
    try:
        import config
        return getattr(config, "COMPANY_KEY", "eyerizz") or "eyerizz"
    except Exception:
        return "eyerizz"


def get_company_name() -> str:
    try:
        import config
        return getattr(config, "COMPANY_NAME", "Eyerizz Eyewear") or "Eyerizz Eyewear"
    except Exception:
        return "Eyerizz Eyewear"


def get_device_id() -> str:
    """Return local machine hostname for multi-device sync audit trail."""
    try:
        return socket.gethostname()
    except Exception:
        return "Desktop-Client"


def is_cloud_configured() -> bool:
    """Check if active company environment has Supabase URL and Key configured."""
    url = get_supabase_url()
    key = get_supabase_key()
    return bool(url and key and "yourproject" not in url and "your-project" not in url)


# ── Field-Level Encryption ────────────────────────────────────────────────────

def encrypt_field(value: str) -> str:
    """Encrypt a string field using Fernet (AES). Returns base64 ciphertext or original if no key."""
    if not value or not _HAS_FERNET:
        return value
    key = get_encryption_key()
    if not key:
        return value
    try:
        f = Fernet(key.encode("utf-8") if isinstance(key, str) else key)
        return f.encrypt(value.encode("utf-8")).decode("utf-8")
    except Exception:
        return value


def decrypt_field(value: str) -> str:
    """Decrypt a Fernet-encrypted field. Returns original if decryption fails or no key."""
    if not value or not _HAS_FERNET:
        return value
    key = get_encryption_key()
    if not key:
        return value
    try:
        f = Fernet(key.encode("utf-8") if isinstance(key, str) else key)
        return f.decrypt(value.encode("utf-8")).decode("utf-8")
    except (InvalidToken, Exception):
        # Not encrypted or wrong key — return as-is
        return value


def _decrypt_rows(rows: list[dict], fields: list[str]) -> list[dict]:
    """Decrypt specified fields in a list of row dicts."""
    if not get_encryption_key():
        return rows
    for row in rows:
        for f in fields:
            if f in row and row[f]:
                row[f] = decrypt_field(str(row[f]))
    return rows


# ── Supabase Auth ─────────────────────────────────────────────────────────────

def supabase_auth_login(email: str, password: str) -> tuple[bool, dict | str]:
    """Authenticate via Supabase GoTrue (email/password). Returns (ok, result_or_error)."""
    url = get_supabase_url().rstrip("/")
    anon = get_anon_key()
    if not url or not anon:
        return False, "Supabase URL or anon key not configured."

    auth_url = f"{url}/auth/v1/token?grant_type=password"
    body = json.dumps({"email": email, "password": password}).encode("utf-8")
    headers = {
        "apikey": anon,
        "Content-Type": "application/json",
    }
    req = urllib.request.Request(auth_url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return True, data
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="ignore")
        try:
            msg = json.loads(err).get("error_description") or json.loads(err).get("msg") or err
        except Exception:
            msg = err
        return False, str(msg)
    except Exception as e:
        return False, str(e)


def check_internet_connectivity(timeout: float = 1.0) -> bool:
    """Fast check to determine if the computer has active internet connection (1s timeout)."""
    try:
        urllib.request.urlopen("https://1.1.1.1", timeout=timeout)
        return True
    except Exception:
        pass
    try:
        url = get_supabase_url()
        if url:
            urllib.request.urlopen(url, timeout=timeout)
            return True
    except Exception:
        pass
    return False


def generate_recon_hash(company_key: str, bank_name: str, txn_date: str, bank_number: str, amount: float, disc_type: str = "bank_only") -> str:
    """Generate deterministic HMAC-SHA256 fingerprint for a transaction / discrepancy."""
    secret = (get_supabase_key() or "bank_recon_secure_hash").encode("utf-8")
    normalized_amt = f"{float(amount):.2f}"
    raw_str = f"{company_key.lower().strip()}:{bank_name.upper().strip()}:{str(txn_date).strip()}:{str(bank_number).strip()}:{normalized_amt}:{disc_type.lower().strip()}"
    return hmac.new(secret, raw_str.encode("utf-8"), hashlib.sha256).hexdigest()


def generate_mutation_hash(company_key: str, bank_name: str, acc_num: str, txn_date: str, amount: float, mut_type: str, desc: str = "") -> str:
    """Generate deterministic HMAC-SHA256 fingerprint for a bank mutation record."""
    secret = (get_supabase_key() or "bank_recon_secure_hash").encode("utf-8")
    normalized_amt = f"{float(amount):.2f}"
    short_desc = str(desc).strip()[:40]
    raw_str = f"{company_key.lower().strip()}:{bank_name.upper().strip()}:{str(acc_num).strip()}:{str(txn_date).strip()}:{normalized_amt}:{mut_type.upper().strip()}:{short_desc}"
    return hmac.new(secret, raw_str.encode("utf-8"), hashlib.sha256).hexdigest()


def _api_headers(access_token: str | None = None) -> dict[str, str]:
    """Build standard headers for Supabase PostgREST requests."""
    key = get_supabase_key()
    token = access_token if access_token else key
    return {
        "apikey": key,
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _make_request(endpoint: str, method: str = "GET", data: dict | list | None = None, headers: dict | None = None, timeout: int = 15):
    """Execute an HTTP request to Supabase PostgREST endpoint."""
    if not is_cloud_configured():
        raise ConnectionError("Supabase is not configured in current company .env.")

    base_url = get_supabase_url().rstrip("/")
    if not endpoint.startswith("/"):
        endpoint = f"/{endpoint}"
    full_url = f"{base_url}/rest/v1{endpoint}"

    req_headers = _api_headers()
    if headers:
        req_headers.update(headers)

    body_bytes = None
    if data is not None:
        body_bytes = json.dumps(data, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(full_url, data=body_bytes, headers=req_headers, method=method)
    
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            resp_body = resp.read().decode("utf-8")
            if not resp_body:
                return {}
            try:
                return json.loads(resp_body)
            except Exception:
                return resp_body
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Cloud API HTTP Error {e.code}: {err_msg}")
    except urllib.error.URLError as e:
        raise ConnectionError(f"Cloud Connection Error: {e.reason}")


def test_connection() -> tuple[bool, str]:
    """Test connection to the Supabase database for active company."""
    if not is_cloud_configured():
        return False, "SUPABASE_URL or SUPABASE_KEY not set in active company .env"
    try:
        _make_request("/bank_discrepancies?limit=1", method="GET", timeout=5)
        return True, "Connected successfully"
    except Exception as e:
        return False, str(e)


def _get_active_uploader(user_profile: str = "") -> str:
    """Resolve active user/profile name for Supabase uploaded_by."""
    u_prof = str(user_profile or "").strip()
    if u_prof and u_prof not in ("Desktop-User", "odoo@example.com", "—", "-"):
        return u_prof

    # 1. Try to get active username/profile from odoo_inspector
    try:
        import odoo_inspector
        _, _, active_u, _ = odoo_inspector._get_credentials()
        if active_u and active_u.strip():
            active_u = active_u.strip()
            # If active_u matches username in PREDEFINED_ACCOUNTS, return the friendly profile display name
            from config import PREDEFINED_ACCOUNTS
            for display_name, acc_info in PREDEFINED_ACCOUNTS.items():
                if acc_info.get("username") == active_u or display_name.lower() == active_u.lower():
                    return display_name
            return active_u
    except Exception:
        pass

    # 2. Try PREDEFINED_ACCOUNTS first display name
    try:
        from config import PREDEFINED_ACCOUNTS
        if PREDEFINED_ACCOUNTS:
            return next(iter(PREDEFINED_ACCOUNTS.keys()))
    except Exception:
        pass

    # 3. Try environment variables
    for env_k in ("ODOO_USER", "ODOO_EMAIL", "USER"):
        val = os.environ.get(env_k)
        if val and val.strip():
            return val.strip()

    return "Desktop-User"


# ─── 1. Merchant Transactions Sync ────────────────────────────────────────────

def push_merchant_transactions(transactions: list[dict], user_profile: str = "") -> dict:
    """
    Push EDC / Merchant Report transactions to Supabase bank_merchant_transactions table.
    Uses atomic HMAC-SHA256 on_conflict=recon_hash to eliminate duplicates across devices.
    """
    if not is_cloud_configured():
        return {"success": False, "error": "Cloud not configured in .env", "count": 0}

    if not transactions:
        return {"success": True, "count": 0, "message": "No merchant transactions to upload"}

    ckey = get_company_key()
    cname = get_company_name()
    dev_id = get_device_id()
    uploader = _get_active_uploader(user_profile)
    
    records = []
    now_iso = datetime.now().isoformat()

    for item in transactions:
        t_date = item.get("date", item.get("transaction_date", ""))
        if isinstance(t_date, (datetime, date)):
            t_date = t_date.strftime("%Y-%m-%d")
        else:
            s = str(t_date).strip()
            parsed_d = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d %b %Y"):
                try:
                    parsed_d = datetime.strptime(s[:10] if len(s) >= 10 else s, fmt).strftime("%Y-%m-%d")
                    break
                except Exception:
                    pass
            t_date = parsed_d or s

        b_name = str(item.get("bank", item.get("bank_name", ""))).upper().strip()
        trace_no = str(item.get("number", item.get("trace_number", item.get("bank_number", "")))).strip()
        gross_amt = float(item.get("amount", item.get("gross_amount", 0.0)))
        net_amt = float(item.get("net_amount", gross_amt))
        fee_amt = float(item.get("fee_amount", item.get("mdr_fee", 0.0)))
        raw_store = str(item.get("store", "")).strip()
        alias_hint = str(item.get("alias", "")).strip().lower()
        item_mid = str(item.get("mid", "")).strip().lstrip("'")
        store = ""

        # 1. If explicit store name provided (not generic "main" / "other")
        if raw_store and raw_store.lower() not in ("main", "general", "other", "—", "-"):
            store = raw_store
        else:
            from config import BANK_ACCOUNTS
            b_low = b_name.lower()
            accs = BANK_ACCOUNTS.get(b_low, {})
            if alias_hint and alias_hint in accs and accs[alias_hint].get("store"):
                store = accs[alias_hint]["store"]
            elif raw_store and raw_store.lower() in accs and accs[raw_store.lower()].get("store"):
                store = accs[raw_store.lower()]["store"]
            elif item_mid:
                for ak, adata in accs.items():
                    if (adata.get("mid") or "").strip().lstrip("'") == item_mid and adata.get("store"):
                        store = adata["store"]
                        break
            if not store:
                if "main" in accs and accs["main"].get("store"):
                    store = accs["main"]["store"]
                else:
                    for ak, adata in accs.items():
                        if adata.get("store"):
                            store = adata["store"]
                            break
            if not store:
                store = "Sanur"

        card_t = str(item.get("card_type", item.get("category", item.get("type", "Credit Card")))).strip()
        if not card_t or card_t.upper() in ("EDC", "BANK", "NONE"):
            card_t = "Credit Card"
        fname = str(item.get("filename", "")).strip()

        # ponytail: hash computed on plaintext BEFORE encryption — ensures dedup consistency
        r_hash = generate_recon_hash(ckey, b_name, t_date, trace_no, gross_amt, card_t)
        trace_no = encrypt_field(trace_no)
        enc_bank = encrypt_field(b_name)
        enc_store = encrypt_field(store)
        enc_company = encrypt_field(cname)

        records.append({
            "company_key": ckey,
            "company": enc_company,
            "bank_name": enc_bank,
            "store": enc_store,
            "transaction_date": t_date,
            "trace_number": trace_no,
            "card_type": card_t,
            "gross_amount": gross_amt,
            "net_amount": net_amt,
            "fee_amount": fee_amt,
            "recon_hash": r_hash,
            "uploaded_by": uploader,
            "device_id": dev_id,
            "created_at": now_iso,
        })


    # Deduplicate in-memory by recon_hash to avoid PostgreSQL error 21000
    deduped_records = list({r["recon_hash"]: r for r in records}.values())

    headers = {"Prefer": "resolution=merge-duplicates,return=representation"}
    total_synced = 0
    chunk_size = 150
    try:
        for i in range(0, len(deduped_records), chunk_size):
            chunk = deduped_records[i:i + chunk_size]
            resp = _make_request(
                "/bank_merchant_transactions?on_conflict=recon_hash",
                method="POST",
                data=chunk,
                headers=headers
            )
            total_synced += len(resp) if isinstance(resp, list) else len(chunk)
        return {"success": True, "count": total_synced}
    except Exception as e:
        return {"success": False, "error": str(e), "count": total_synced}


# ─── 2. Bank Mutation Transactions Sync ───────────────────────────────────────

def push_mutation_transactions(mutations: list[dict], user_profile: str = "") -> dict:
    """
    Push Bank Mutation statement rows to Supabase bank_mutation_transactions table.
    Uses atomic HMAC-SHA256 on_conflict=recon_hash to eliminate duplicates across devices.
    """
    if not is_cloud_configured():
        return {"success": False, "error": "Cloud not configured in .env", "count": 0}

    if not mutations:
        return {"success": True, "count": 0, "message": "No mutations to upload"}

    ckey = get_company_key()
    cname = get_company_name()
    dev_id = get_device_id()
    uploader = _get_active_uploader(user_profile)
    
    records = []
    now_iso = datetime.now().isoformat()

    for item in mutations:
        t_date = item.get("date", item.get("transaction_date", ""))
        if isinstance(t_date, (datetime, date)):
            t_date = t_date.strftime("%Y-%m-%d")
        else:
            s = str(t_date).strip()
            parsed_d = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d %b %Y"):
                try:
                    parsed_d = datetime.strptime(s[:10] if len(s) >= 10 else s, fmt).strftime("%Y-%m-%d")
                    break
                except Exception:
                    pass
            t_date = parsed_d or s

        b_name = str(item.get("bank", item.get("bank_name", ""))).upper().strip()
        acc_num = str(item.get("account_number", item.get("acc", ""))).strip()
        amt = float(item.get("amount", 0.0))
        m_type = str(item.get("mutation_type", item.get("type", "CR"))).upper().strip()[:50]
        desc = str(item.get("description", item.get("remark", ""))).strip()

        # ponytail: hash on plaintext before encryption
        r_hash = generate_mutation_hash(ckey, b_name, acc_num, t_date, amt, m_type, desc)
        enc_bank = encrypt_field(b_name)
        enc_company = encrypt_field(cname)

        records.append({
            "company_key": ckey,
            "company": enc_company,
            "bank_name": enc_bank,
            "transaction_date": t_date,
            "amount": amt,
            "mutation_type": m_type,
            "recon_hash": r_hash,
            "uploaded_by": uploader,
            "device_id": dev_id,
            "created_at": now_iso,
            "updated_at": now_iso,
        })

    # Deduplicate in-memory by recon_hash to avoid PostgreSQL error 21000
    deduped_records = list({r["recon_hash"]: r for r in records}.values())

    headers = {"Prefer": "resolution=merge-duplicates,return=representation"}
    total_synced = 0
    chunk_size = 150
    try:
        for i in range(0, len(deduped_records), chunk_size):
            chunk = deduped_records[i:i + chunk_size]
            resp = _make_request(
                "/bank_mutation_transactions?on_conflict=recon_hash",
                method="POST",
                data=chunk,
                headers=headers
            )
            total_synced += len(resp) if isinstance(resp, list) else len(chunk)
        return {"success": True, "count": total_synced}
    except Exception as e:
        return {"success": False, "error": str(e), "count": total_synced}


# ─── 3. Discrepancies Sync ───────────────────────────────────────────────────

def push_bank_discrepancies(discrepancies: list[dict], recon_date: str = "", company: str = "", user_profile: str = "") -> dict:
    """Push transactions to Supabase bank_discrepancies table with hash deduplication."""
    if not is_cloud_configured():
        return {"success": False, "error": "Cloud not configured in .env", "count": 0}

    if not discrepancies:
        return {"success": True, "count": 0, "message": "No discrepancies to upload"}

    ckey = get_company_key()
    cname = company or get_company_name()
    dev_id = get_device_id()
    uploader = _get_active_uploader(user_profile)

    
    records = []
    now_iso = datetime.now().isoformat()

    for item in discrepancies:
        t_date = item.get("date", "")
        if isinstance(t_date, (datetime, date)):
            t_date = t_date.strftime("%Y-%m-%d")
        else:
            t_date = str(t_date).strip()

        r_date = recon_date or t_date
        if isinstance(r_date, (datetime, date)):
            r_date = r_date.strftime("%Y-%m-%d")
        else:
            r_date = str(r_date).strip()

        amt = float(item.get("amount", 0.0))
        disc_type = str(item.get("discrepancy_type") or "bank_only").strip()
        b_num = str(item.get("number_bank", item.get("bank_number", ""))).strip()
        o_num = str(item.get("number_odo", item.get("odoo_number", ""))).strip()
        o_ref = str(item.get("invoice_no", item.get("odoo_reference", ""))).strip()
        is_recon = str(item.get("is_reconciled", item.get("reconciled", "Yes"))).strip()
        b_name = str(item.get("bank", item.get("bank_name", ""))).upper().strip()

        # ponytail: hash on plaintext before encryption
        r_hash = generate_recon_hash(ckey, b_name, t_date, b_num or o_num, amt, disc_type)
        b_num = encrypt_field(b_num)
        enc_bank = encrypt_field(b_name)
        enc_company = encrypt_field(cname)

        records.append({
            "company_key": ckey,
            "company": enc_company,
            "recon_date": r_date,
            "bank_name": enc_bank,
            "journal": str(item.get("journal", "")).strip(),
            "transaction_date": t_date,
            "bank_number": b_num,
            "odoo_number": o_num,
            "odoo_reference": o_ref,
            "is_reconciled": is_recon,
            "discrepancy_type": disc_type,
            "filename": str(item.get("filename", item.get("filename_bank", ""))).strip(),
            "amount": amt,
            "recon_hash": r_hash,
            "status": "Pending",
            "uploaded_by": uploader,
            "device_id": dev_id,
            "created_at": now_iso,
            "updated_at": now_iso,
        })

    # Deduplicate in-memory by recon_hash to avoid PostgreSQL error 21000
    deduped_records = list({r["recon_hash"]: r for r in records}.values())

    headers = {"Prefer": "resolution=merge-duplicates,return=representation"}
    total_synced = 0
    chunk_size = 150
    try:
        for i in range(0, len(deduped_records), chunk_size):
            chunk = deduped_records[i:i + chunk_size]
            try:
                resp = _make_request(
                    "/bank_discrepancies?on_conflict=recon_hash",
                    method="POST",
                    data=chunk,
                    headers=headers
                )

            except Exception as e_post:
                if "recon_hash" in str(e_post) or "42703" in str(e_post) or "company_key" in str(e_post):
                    legacy_chunk = []
                    for r in chunk:
                        r_copy = dict(r)
                        r_copy.pop("company_key", None)
                        r_copy.pop("recon_hash", None)
                        r_copy.pop("uploaded_by", None)
                        r_copy.pop("device_id", None)
                        legacy_chunk.append(r_copy)
                    resp = _make_request(
                        "/bank_discrepancies?on_conflict=recon_date,bank_name,transaction_date,bank_number,odoo_number,discrepancy_type,amount",
                        method="POST",
                        data=legacy_chunk,
                        headers=headers
                    )
                else:
                    raise e_post
            total_synced += len(resp) if isinstance(resp, list) else len(chunk)
        return {"success": True, "count": total_synced}
    except Exception as e:
        return {"success": False, "error": str(e), "count": total_synced}


def fetch_discrepancies(
    status: str | None = None,
    bank: str | None = None,
    recon_date: str | None = None,
    limit: int = 250,
) -> list[dict]:
    """Fetch discrepancies from cloud scoped to active company with global status filters."""
    if not is_cloud_configured():
        return []

    ckey = get_company_key()
    # ponytail: bank_name encrypted → can't filter server-side, filter after decrypt
    params = [f"company_key=eq.{urllib.parse.quote(ckey)}", "order=transaction_date.desc,id.desc", f"limit={limit}"]

    if status and status.upper() != "ALL":
        st_norm = "Pending" if "pending" in status.lower() else ("Resolve" if "resolve" in status.lower() else status)
        params.append(f"status=eq.{urllib.parse.quote(st_norm)}")

    if recon_date:
        params.append(f"recon_date=eq.{urllib.parse.quote(recon_date)}")

    query_str = "&".join(params)
    endpoint = f"/bank_discrepancies?{query_str}"

    try:
        res = _make_request(endpoint, method="GET")
        items = res if isinstance(res, list) else []
        _decrypt_rows(items, ["bank_number", "bank_name", "company"])
        # Client-side bank filter after decryption
        if bank and bank.upper() != "ALL":
            items = [r for r in items if str(r.get("bank_name", "")).upper() == bank.upper()]
        return items
    except Exception as e:
        err_str = str(e)
        if "company_key" in err_str or "42703" in err_str:
            try:
                legacy_params = [p for p in params if not p.startswith("company_key=")]
                legacy_endpoint = f"/bank_discrepancies?{'&'.join(legacy_params)}"
                res = _make_request(legacy_endpoint, method="GET")
                return res if isinstance(res, list) else []
            except Exception:
                pass
        print(f"[CloudSync] Fetch failed: {e}")
        return []


def resolve_discrepancy(
    item_id: int,
    sales_person: str,
    action_type: str,
    sales_notes: str,
    odoo_reference: str = "",
) -> bool:
    """Update a discrepancy record as resolved with global status 'Resolve'."""
    if not is_cloud_configured():
        return False

    payload = {
        "status": "Resolve",
        "resolved_by": sales_person.strip(),
        "action_type": action_type.strip(),
        "sales_notes": sales_notes.strip(),
        "odoo_reference": odoo_reference.strip(),
        "updated_at": datetime.now().isoformat(),
    }

    headers = {"Prefer": "return=representation"}
    try:
        _make_request(f"/bank_discrepancies?id=eq.{item_id}", method="PATCH", data=payload, headers=headers)
        return True
    except Exception as e:
        print(f"[CloudSync] Resolve failed for id={item_id}: {e}")
        return False


def reopen_discrepancy(item_id: int) -> bool:
    """Reset discrepancy status back to Pending."""
    if not is_cloud_configured():
        return False

    payload = {
        "status": "Pending",
        "updated_at": datetime.now().isoformat(),
    }

    try:
        _make_request(f"/bank_discrepancies?id=eq.{item_id}", method="PATCH", data=payload)
        return True
    except Exception as e:
        print(f"[CloudSync] Reopen failed for id={item_id}: {e}")
        return False


# ─── 4. Cloud Reconciliation & Query Engine ───────────────────────────────────

def fetch_cloud_count(data_type: str = "merchant", bank: str = "", date_from: str = "", date_to: str = "") -> int:
    """Row count for cloud data. Bank filter applied client-side (bank_name encrypted)."""
    if not is_cloud_configured():
        return 0

    table = "bank_merchant_transactions" if data_type == "merchant" else "bank_mutation_transactions"
    ckey = get_company_key()
    params = [f"company_key=eq.{urllib.parse.quote(ckey)}", "select=id,bank_name"]

    if date_from:
        params.append(f"transaction_date=gte.{urllib.parse.quote(date_from)}")
    if date_to:
        params.append(f"transaction_date=lte.{urllib.parse.quote(date_to)}")

    need_bank_filter = bank and bank.upper() != "ALL"

    if not need_bank_filter:
        # No bank filter needed → use fast HEAD count
        headers = {"Prefer": "count=exact", "Range": "0-0"}
        try:
            endpoint = f"/{table}?{'&'.join(params)}"
            base_url = get_supabase_url().rstrip("/")
            full_url = f"{base_url}/rest/v1{endpoint}"
            req = urllib.request.Request(full_url, headers=_api_headers(), method="HEAD")
            with urllib.request.urlopen(req, timeout=5) as resp:
                crange = resp.headers.get("Content-Range", "")
                if "/" in crange:
                    total = crange.split("/")[-1]
                    return int(total) if total.isdigit() else 0
            return 0
        except Exception:
            pass

    # Fallback / bank filter: fetch minimal rows, decrypt bank_name, count matching
    try:
        res = _make_request(f"/{table}?{'&'.join(params)}&limit=5000", method="GET", timeout=8)
        if not isinstance(res, list):
            return 0
        if need_bank_filter:
            _decrypt_rows(res, ["bank_name"])
            return sum(1 for r in res if str(r.get("bank_name", "")).upper() == bank.upper())
        return len(res)
    except Exception:
        return 0


def fetch_cloud_transactions(data_type: str = "merchant", bank: str = "", date_from: str = "", date_to: str = "", limit: int = 50000) -> list[dict]:
    """
    Fetch merchant or mutation transactions from Supabase cloud with safe streaming pagination.
    Overcomes Supabase PostgREST default 1000 row page limit.
    Returns list of dict rows for direct conversion into pandas DataFrames.
    """
    if not is_cloud_configured():
        return []

    table = "bank_merchant_transactions" if data_type == "merchant" else "bank_mutation_transactions"
    ckey = get_company_key()
    # ponytail: bank_name encrypted → removed from server filter, applied after decrypt
    base_params = [f"company_key=eq.{urllib.parse.quote(ckey)}", "order=transaction_date.asc,id.asc"]

    if date_from:
        base_params.append(f"transaction_date=gte.{urllib.parse.quote(date_from)}")
    if date_to:
        base_params.append(f"transaction_date=lte.{urllib.parse.quote(date_to)}")

    all_rows = []
    page_size = 1000
    offset = 0
    decrypt_fields = ["trace_number", "bank_name", "store", "company"] if data_type == "merchant" else ["bank_name", "company"]

    while len(all_rows) < limit:
        chunk_limit = min(page_size, limit - len(all_rows))
        params = list(base_params) + [f"limit={chunk_limit}", f"offset={offset}"]
        endpoint = f"/{table}?{'&'.join(params)}"
        try:
            res = _make_request(endpoint, method="GET", timeout=15)
            if not isinstance(res, list) or not res:
                break
            _decrypt_rows(res, decrypt_fields)
            all_rows.extend(res)
            if len(res) < chunk_limit:
                break
            offset += len(res)
        except Exception as e:
            print(f"[CloudSync] fetch_cloud_transactions failed for {table} at offset {offset}: {e}")
            break

    # Client-side bank filter after decryption
    if bank and bank.upper() != "ALL":
        all_rows = [r for r in all_rows if str(r.get("bank_name", "")).upper() == bank.upper()]

    return all_rows



def fetch_cloud_dashboard_summary() -> dict:
    """
    Fetch live aggregate statistics from Supabase for the Cloud Dashboard tab.
    Computes total merchant volume, total mutation volume, date span, and recent uploaders.
    """
    if not is_cloud_configured():
        return {
            "configured": False,
            "merchant_count": 0,
            "merchant_volume": 0.0,
            "mutation_count": 0,
            "mutation_volume": 0.0,
            "discrepancy_count": 0,
            "date_span": "—",
            "last_uploader": "—",
            "last_device": "—",
            "last_updated": "—",
        }

    ckey = get_company_key()
    summary = {
        "configured": True,
        "company_key": ckey,
        "company_name": get_company_name(),
        "merchant_count": 0,
        "merchant_volume": 0.0,
        "mutation_count": 0,
        "mutation_volume": 0.0,
        "discrepancy_count": 0,
        "date_span": "—",
        "last_uploader": "—",
        "last_device": "—",
        "last_updated": "—",
    }

    try:
        # Fetch recent merchant records
        m_rows = _make_request(f"/bank_merchant_transactions?company_key=eq.{urllib.parse.quote(ckey)}&order=transaction_date.desc,id.desc&limit=500", method="GET", timeout=6)
        if isinstance(m_rows, list) and m_rows:
            _decrypt_rows(m_rows, ["bank_name", "store", "company", "trace_number"])
            summary["merchant_count"] = len(m_rows)
            summary["merchant_volume"] = sum(float(r.get("gross_amount", 0.0)) for r in m_rows)
            dates = [r["transaction_date"] for r in m_rows if r.get("transaction_date")]
            if dates:
                summary["date_span"] = f"{min(dates)} to {max(dates)}"
            latest = m_rows[0]
            summary["last_uploader"] = latest.get("uploaded_by", "—")
            summary["last_device"] = latest.get("device_id", "—")
            summary["last_updated"] = str(latest.get("created_at", "—"))[:19].replace("T", " ")

        # Fetch recent mutation records
        mut_rows = _make_request(f"/bank_mutation_transactions?company_key=eq.{urllib.parse.quote(ckey)}&order=transaction_date.desc,id.desc&limit=500", method="GET", timeout=6)
        if isinstance(mut_rows, list) and mut_rows:
            _decrypt_rows(mut_rows, ["bank_name", "company"])
            summary["mutation_count"] = len(mut_rows)
            summary["mutation_volume"] = sum(float(r.get("amount", 0.0)) for r in mut_rows)

        # Fetch discrepancy count
        disc_rows = _make_request(f"/bank_discrepancies?company_key=eq.{urllib.parse.quote(ckey)}&status=eq.Pending&limit=500", method="GET", timeout=6)
        if isinstance(disc_rows, list):
            summary["discrepancy_count"] = len(disc_rows)

    except Exception as e:
        print(f"[CloudSync] fetch_cloud_dashboard_summary failed: {e}")

    return summary


def fetch_cloud_analytics(bank: str = "ALL", period: str = "3d", custom_from: str = "", custom_to: str = "", data_type: str = "merchant") -> dict:
    """
    Compute comprehensive executive financial analytics from Supabase Cloud data.
    Separates EDC / Merchant transactions and Bank Account Mutation transactions into distinct streams.
    Provides volume totals, ATV, store distribution, bank share, and daily/monthly performance run-rate.
    Accurately computes missing settlement dates against the requested period calendar.
    """
    if not is_cloud_configured():
        return {
            "configured": False,
            "data_type": data_type,
            "total_gross": 0.0,
            "total_net": 0.0,
            "total_fee": 0.0,
            "total_txns": 0,
            "atv": 0.0,
            "daily_run_rate": 0.0,
            "active_days_count": 0,
            "missing_dates": [],
            "missing_dates_count": 0,
            "expected_days_count": 0,
            "date_span": "—",
            "peak_day": ("—", 0.0, 0),
            "granularity": "daily",
            "by_bank": {},
            "by_store": {},
            "by_card": {},
            "daily_stats": [],
            "last_uploader": "—",
            "last_device": "—",
            "last_updated": "—",
        }

    try:
        # Calculate date_from and date_to based on period filter
        date_from = None
        date_to = None
        today = datetime.now()

        if period == "custom" and custom_from and custom_to:
            try:
                f_dt = datetime.strptime(custom_from, "%Y-%m-%d")
                t_dt = datetime.strptime(custom_to, "%Y-%m-%d")
                if f_dt > t_dt:
                    f_dt, t_dt = t_dt, f_dt
                # Enforce 1 month maximum range (supports 28, 30, and 31 day months)
                import calendar
                y = f_dt.year + (f_dt.month // 12)
                m = (f_dt.month % 12) + 1
                max_d = min(f_dt.day, calendar.monthrange(y, m)[1])
                max_1mo = f_dt.replace(year=y, month=m, day=max_d)
                if t_dt > max_1mo:
                    t_dt = max_1mo
                date_from = f_dt.strftime("%Y-%m-%d")
                date_to = t_dt.strftime("%Y-%m-%d")
            except Exception:
                date_from = (today - timedelta(days=2)).strftime("%Y-%m-%d")
                date_to = today.strftime("%Y-%m-%d")

        elif period == "3d":
            date_from = (today - timedelta(days=2)).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")
        elif period == "7d":
            date_from = (today - timedelta(days=6)).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")
        elif period == "14d":
            date_from = (today - timedelta(days=13)).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")
        elif period == "30d":
            date_from = (today - timedelta(days=29)).strftime("%Y-%m-%d")
            date_to = today.strftime("%Y-%m-%d")

        # Query isolated data stream
        target_type = "mutation" if data_type == "mutation" else "merchant"
        raw_txns = fetch_cloud_transactions(data_type=target_type, bank=bank, date_from=date_from, date_to=date_to, limit=35000)

        txns = []
        for t in raw_txns:
            t_norm = dict(t)
            if target_type == "mutation":
                amt = float(t.get("amount", 0.0))
                t_norm["gross_amount"] = amt
                t_norm["net_amount"] = amt
                t_norm["fee_amount"] = 0.0
                t_norm["card_type"] = t.get("mutation_type") or "Mutation"
            else:
                g = float(t.get("gross_amount", 0.0))
                t_norm["gross_amount"] = g
                t_norm["net_amount"] = float(t.get("net_amount", g))
                t_norm["fee_amount"] = float(t.get("fee_amount", 0.0))
                t_norm["card_type"] = str(t.get("card_type") or "EDC").upper().strip()
            txns.append(t_norm)

        total_gross = 0.0
        total_net = 0.0
        total_fee = 0.0
        total_txns = len(txns)

        by_bank = {}
        by_store = {}
        by_card = {}
        by_date = {}

        last_u = "—"
        last_d = "—"
        last_up = "—"

        try:
            ckey = get_company_key()
            tbl_name = "bank_mutation_transactions" if target_type == "mutation" else "bank_merchant_transactions"
            latest_rows = _make_request(f"/{tbl_name}?company_key=eq.{urllib.parse.quote(ckey)}&order=created_at.desc&limit=1", method="GET", timeout=5)
            if isinstance(latest_rows, list) and latest_rows:
                lr = latest_rows[0]
                last_u = lr.get("uploaded_by", "—")
                last_d = lr.get("device_id", "—")
                last_up = str(lr.get("created_at", "—"))[:19].replace("T", " ")
        except Exception:
            if txns:
                latest = txns[-1]
                last_u = latest.get("uploaded_by", "—")
                last_d = latest.get("device_id", "—")
                last_up = str(latest.get("created_at", "—"))[:19].replace("T", " ")

        for t in txns:
            g = float(t.get("gross_amount", 0.0))
            n = float(t.get("net_amount", g))
            f = float(t.get("fee_amount", 0.0))
            total_gross += g
            total_net += n
            total_fee += f

            b = str(t.get("bank_name") or "BCA").upper().strip()
            st = str(t.get("store") or "General").capitalize().strip()
            card = str(t.get("card_type") or "EDC").upper().strip()
            dt = str(t.get("transaction_date") or "").strip()

            # Group by bank
            by_bank.setdefault(b, {"gross": 0.0, "count": 0, "fee": 0.0})
            by_bank[b]["gross"] += g
            by_bank[b]["count"] += 1
            by_bank[b]["fee"] += f

            # Group by store
            by_store.setdefault(st, {"gross": 0.0, "count": 0})
            by_store[st]["gross"] += g
            by_store[st]["count"] += 1

            # Group by card type
            by_card.setdefault(card, {"gross": 0.0, "count": 0})
            by_card[card]["gross"] += g
            by_card[card]["count"] += 1

            # Group by date with bank breakdowns
            if dt:
                by_date.setdefault(dt, {"gross": 0.0, "count": 0, "banks": {}})
                by_date[dt]["gross"] += g
                by_date[dt]["count"] += 1
                by_date[dt]["banks"].setdefault(b, 0.0)
                by_date[dt]["banks"][b] += g

        active_days = sorted(by_date.keys())
        active_days_count = len(active_days)
        missing_dates = []
        expected_days_count = active_days_count

        if date_from and date_to:
            try:
                start_dt = datetime.strptime(date_from, "%Y-%m-%d")
                end_dt = datetime.strptime(date_to, "%Y-%m-%d")
                curr = start_dt
                active_set = set(active_days)
                expected_days = []
                while curr <= end_dt:
                    c_str = curr.strftime("%Y-%m-%d")
                    expected_days.append(c_str)
                    if c_str not in active_set:
                        missing_dates.append(c_str)
                    curr += timedelta(days=1)
                expected_days_count = len(expected_days)
            except Exception:
                pass


        atv = (total_gross / total_txns) if total_txns > 0 else 0.0
        daily_run_rate = (total_gross / active_days_count) if active_days_count > 0 else 0.0

        # Automatic rollup to Monthly if dataset has > 40 unique dates (e.g. 1 year of data)
        granularity = "daily"
        daily_stats = []

        if active_days_count > 40 and period in ("all", "1y"):
            granularity = "monthly"
            by_month = {}
            for d, info in by_date.items():
                m_key = d[:7]  # YYYY-MM
                by_month.setdefault(m_key, {"gross": 0.0, "count": 0, "banks": {}})
                by_month[m_key]["gross"] += info["gross"]
                by_month[m_key]["count"] += info["count"]
                for b_name, b_vol in info["banks"].items():
                    by_month[m_key]["banks"].setdefault(b_name, 0.0)
                    by_month[m_key]["banks"][b_name] += b_vol

            for m in sorted(by_month.keys()):
                info = by_month[m]
                daily_stats.append({
                    "date": m,
                    "gross": info["gross"],
                    "count": info["count"],
                    "avg": (info["gross"] / info["count"]) if info["count"] > 0 else 0.0,
                    "banks": info["banks"]
                })
        else:
            timeline_days = expected_days if (date_from and date_to and expected_days) else active_days
            for d in timeline_days:
                if d in by_date:
                    info = by_date[d]
                    daily_stats.append({
                        "date": d,
                        "gross": info["gross"],
                        "count": info["count"],
                        "avg": (info["gross"] / info["count"]) if info["count"] > 0 else 0.0,
                        "banks": info["banks"],
                        "is_missing": False
                    })
                else:
                    daily_stats.append({
                        "date": d,
                        "gross": 0.0,
                        "count": 0,
                        "avg": 0.0,
                        "banks": {},
                        "is_missing": True
                    })

        peak_day = ("—", 0.0, 0)
        if daily_stats:
            p = max(daily_stats, key=lambda x: x["gross"])
            peak_day = (p["date"], p["gross"], p["count"])

        if date_from and date_to:
            date_span = f"{date_from} to {date_to}"
        elif active_days:
            date_span = f"{active_days[0]} to {active_days[-1]}"
        else:
            date_span = "—"


        return {
            "configured": True,
            "total_gross": total_gross,
            "total_net": total_net,
            "total_fee": total_fee,
            "total_txns": total_txns,
            "atv": atv,
            "daily_run_rate": daily_run_rate,
            "active_days_count": active_days_count,
            "expected_days_count": expected_days_count,
            "missing_dates": missing_dates,
            "missing_dates_count": len(missing_dates),
            "date_span": date_span,
            "peak_day": peak_day,
            "granularity": granularity,
            "by_bank": by_bank,
            "by_store": by_store,
            "by_card": by_card,
            "daily_stats": daily_stats,
            "last_uploader": last_u,
            "last_device": last_d,
            "last_updated": last_up,
        }

    except Exception as e:
        print(f"[CloudSync] fetch_cloud_analytics error: {e}")
        return {
            "configured": False,
            "total_gross": 0.0,
            "total_net": 0.0,
            "total_fee": 0.0,
            "total_txns": 0,
            "atv": 0.0,
            "daily_run_rate": 0.0,
            "active_days_count": 0,
            "date_span": "—",
            "peak_day": ("—", 0.0, 0),
            "granularity": "daily",
            "by_bank": {},
            "by_store": {},
            "by_card": {},
            "daily_stats": [],
            "last_uploader": "—",
            "last_device": "—",
            "last_updated": "—",
        }



def fetch_cloud_coverage_matrix() -> list[dict]:
    """
    Fetch distinct date ranges for Merchant Settlements and Bank Mutations per bank account.
    Returns structured data for the Date Coverage by Account & Mutation table in Cloud Dashboard.
    """
    if not is_cloud_configured():
        return []

    # Query distinct dates for merchant
    m_rows = fetch_cloud_transactions(data_type="merchant", limit=50000)
    # Query distinct dates for mutations
    mut_rows = fetch_cloud_transactions(data_type="mutation", limit=50000)

    from collections import defaultdict
    bank_m_dates = defaultdict(set)
    bank_mut_dates = defaultdict(set)

    for r in m_rows:
        b = str(r.get("bank_name") or "BCA").upper().strip()
        d = str(r.get("transaction_date") or "").strip()
        if b and d:
            bank_m_dates[b].add(d)

    for r in mut_rows:
        b = str(r.get("bank_name") or "BCA").upper().strip()
        d = str(r.get("transaction_date") or "").strip()
        if b and d:
            bank_mut_dates[b].add(d)

    def _fmt_dates(d_set):
        clean = sorted(d_set)
        if not clean:
            return "—"
        try:
            d_start = datetime.strptime(clean[0], "%Y-%m-%d").strftime("%d/%m/%y")
            d_end = datetime.strptime(clean[-1], "%Y-%m-%d").strftime("%d/%m/%y")
            if len(clean) == 1:
                return f"{d_start} (1 day)"
            return f"{d_start} – {d_end} ({len(clean)} days)"
        except Exception:
            return f"{clean[0]} – {clean[-1]} ({len(clean)} days)"

    all_banks = sorted(set(list(bank_m_dates.keys()) + list(bank_mut_dates.keys())))
    if not all_banks:
        all_banks = ["BCA", "MANDIRI", "BRI"]

    matrix = []
    for b in all_banks:
        matrix.append({
            "bank": b,
            "merchant_dates": _fmt_dates(bank_m_dates[b]),
            "merchant_count": len(bank_m_dates[b]),
            "mutation_dates": _fmt_dates(bank_mut_dates[b]),
            "mutation_count": len(bank_mut_dates[b]),
        })

    return matrix


# Backward compatibility alias
get_cloud_summary = fetch_cloud_dashboard_summary



def sync_local_to_cloud(user_profile: str = "") -> dict:
    """
    Parse all local statement files (input/...) and mutation CSV files (mutation/...)
    and sync them to Supabase in batch with deduplication.
    """
    if not is_cloud_configured():
        return {"success": False, "error": "Cloud not configured in .env", "merchant_count": 0, "mutation_count": 0}

    uploader = _get_active_uploader(user_profile)
    all_merchants = []

    # 1. Parse BCA Merchant
    try:
        from config import BCA_EXCEL_DIR, BCA_EXCEL_PASSWORD, BCA_AMOUNT_COLUMN, BCA_DATE_COLUMN, BCA_NUMBER_COLUMN, BANK_ACCOUNTS
        from readers.bca_reader import _read_one_bca
        bca_store = BANK_ACCOUNTS.get("bca", {}).get("main", {}).get("store", "Sanur")
        if BCA_EXCEL_DIR.exists():
            for f in sorted(BCA_EXCEL_DIR.rglob("*.xlsx")):
                if not f.name.startswith((".", "~$")):
                    try:
                        rows = _read_one_bca(f, BCA_EXCEL_PASSWORD, BCA_AMOUNT_COLUMN, BCA_DATE_COLUMN, BCA_NUMBER_COLUMN)
                        for r in rows:
                            r_copy = dict(r)
                            r_copy["bank"] = "BCA"
                            r_copy["store"] = bca_store
                            r_copy["card_type"] = r.get("card_type") or r.get("category") or "Credit Card"
                            all_merchants.append(r_copy)
                    except Exception as e:
                        print(f"  [CloudSync] Error reading BCA file {f.name}: {e}")
    except Exception as e:
        print(f"  [CloudSync] BCA setup error: {e}")

    # 2. Parse Mandiri Merchant
    try:
        from config import MANDIRI_ZIP_DIR, MANDIRI_ZIP_PASSWORD, MANDIRI_AMOUNT_COLUMN, MANDIRI_NUMBER_COLUMN, BANK_ACCOUNTS
        from readers.mandiri_reader import _read_csv_from_bytes
        import pyzipper
        man_store = BANK_ACCOUNTS.get("mandiri", {}).get("main", {}).get("store", "Seminyak")
        if MANDIRI_ZIP_DIR.exists():
            for f in MANDIRI_ZIP_DIR.rglob("*"):
                if f.is_file() and not f.name.startswith((".", "~$")):
                    if f.suffix.lower() == ".csv":
                        for r in _read_csv_from_bytes(f.read_bytes(), MANDIRI_AMOUNT_COLUMN, MANDIRI_NUMBER_COLUMN):
                            r_copy = dict(r)
                            r_copy["bank"] = "MANDIRI"
                            r_copy["store"] = man_store
                            r_copy["card_type"] = r.get("category") or "Debit Card"
                            all_merchants.append(r_copy)
                    elif f.suffix.lower() == ".zip":
                        with pyzipper.AESZipFile(f, "r") as zf:
                            if MANDIRI_ZIP_PASSWORD: zf.setpassword(MANDIRI_ZIP_PASSWORD.encode("utf-8"))
                            for name in zf.namelist():
                                for r in _read_csv_from_bytes(zf.read(name), MANDIRI_AMOUNT_COLUMN, MANDIRI_NUMBER_COLUMN):
                                    r_copy = dict(r)
                                    r_copy["bank"] = "MANDIRI"
                                    r_copy["store"] = man_store
                                    r_copy["card_type"] = r.get("category") or "Debit Card"
                                    all_merchants.append(r_copy)
    except Exception:
        pass

    # 3. Parse BRI Merchant
    try:
        from config import BRI_ZIP_DIR, BRI_PDF_PATTERN, BRI_AMOUNT_COLUMN, BRI_NUMBER_COLUMN, BANK_ACCOUNTS
        from readers.bri_reader import _extract_detail_pdf, _parse_pdf_table
        if BRI_ZIP_DIR.exists():
            for f in BRI_ZIP_DIR.rglob("*.zip"):
                if not f.name.startswith((".", "~$")):
                    alias = f.parent.name.lower()
                    b_store = BANK_ACCOUNTS.get("bri", {}).get(alias, {}).get("store", "Sanur")
                    pdf_bytes = _extract_detail_pdf(f, BRI_PDF_PATTERN)
                    for r in _parse_pdf_table(pdf_bytes, BRI_AMOUNT_COLUMN, BRI_NUMBER_COLUMN):
                        r_copy = dict(r)
                        r_copy["bank"] = "BRI"
                        r_copy["alias"] = alias
                        r_copy["store"] = b_store
                        r_copy["card_type"] = r.get("category") or r.get("card_type") or "Credit Card"
                        all_merchants.append(r_copy)
    except Exception:
        pass

    # 4. Parse Mutations
    all_muts = []
    try:
        from readers.mutation_reader import read_all_mutations
        m_list, u_list = read_all_mutations()
        all_muts = (m_list or []) + (u_list or [])
    except Exception:
        pass

    res_m = push_merchant_transactions(all_merchants, user_profile=uploader) if all_merchants else {"success": True, "count": 0}
    res_mut = push_mutation_transactions(all_muts, user_profile=uploader) if all_muts else {"success": True, "count": 0}

    m_count = res_m.get("count", 0)
    mut_count = res_mut.get("count", 0)
    success = res_m.get("success", False) and res_mut.get("success", False)
    err = res_m.get("error") or res_mut.get("error")

    return {
        "success": success,
        "error": err,
        "merchant_count": m_count,
        "mutation_count": mut_count
    }


def sync_cloud_to_local(banks: list[str] | None = None) -> dict:
    """
    Download cloud transactions from Supabase into local input and mutation directories
    so offline/local reconciler engine can run directly on cloud data.
    """
    if not is_cloud_configured():
        return {"success": False, "error": "Cloud not configured in .env", "merchant_files": 0, "mutation_files": 0}

    import csv
    import openpyxl
    from config import INPUT_DIR, MUTATION_DIR

    m_files = 0
    mut_files = 0

    try:
        # 1. Fetch Merchant Transactions from Cloud
        merchant_txns = fetch_cloud_transactions(data_type="merchant", limit=10000)
        if merchant_txns:
            by_bank = {}
            for t in merchant_txns:
                b = str(t.get("bank_name") or "BCA").upper().strip()
                if banks and "all" not in [x.lower() for x in banks] and b.lower() not in [x.lower() for x in banks]:
                    continue
                by_bank.setdefault(b, []).append(t)

            for b_name, txns in by_bank.items():
                b_key = b_name.lower()
                b_dir = INPUT_DIR / b_key
                b_dir.mkdir(parents=True, exist_ok=True)

                if b_key == "bca":
                    # Create BCA Excel file formatted to match BCA statements
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Report Merchant"
                    ws.append(["PT. BANK CENTRAL ASIA TBK"])
                    ws.append(["MERCHANT REPORT"])
                    ws.append([])
                    ws.append([])
                    headers = ["Merchant Name", "MID", "Terminal ID", "Batch No", "Card Number", "Transaction Date", "Transaction Time", "Original Amount", "Net Amount", "MDR", "Trace Number"]
                    ws.append(headers)
                    for r in txns:
                        ws.append([
                            r.get("store") or "Store",
                            "002669303",
                            "001",
                            "001",
                            r.get("trace_number") or "-",
                            r.get("transaction_date") or "",
                            "12:00:00",
                            float(r.get("gross_amount", 0.0)),
                            float(r.get("net_amount", 0.0)),
                            float(r.get("fee_amount", 0.0)),
                            r.get("trace_number") or "-"
                        ])
                    out_f = b_dir / "BCA_Cloud_Transactions.xlsx"
                    wb.save(out_f)
                    m_files += 1

                elif b_key == "mandiri":
                    by_store = {}
                    for r in txns:
                        st = str(r.get("store") or "main").lower()
                        by_store.setdefault(st, []).append(r)
                    for st, s_txns in by_store.items():
                        st_dir = b_dir / st
                        st_dir.mkdir(parents=True, exist_ok=True)
                        csv_p = st_dir / f"Mandiri_Cloud_{st}.csv"
                        with open(csv_p, "w", newline="", encoding="utf-8") as cf:
                            cw = csv.writer(cf)
                            cw.writerow(["MERCHANT STATEMENT REPORT", "", "", "", "", ""])
                            cw.writerow(["TRANSACTION DATE", "AUTHCODE", "CARD NUMBER", "AMOUNT", "MDR", "NET AMOUNT"])
                            for r in s_txns:
                                cw.writerow([
                                    r.get("transaction_date", ""),
                                    r.get("trace_number", ""),
                                    r.get("trace_number", ""),
                                    f"{float(r.get('gross_amount', 0.0)):.2f}",
                                    f"{float(r.get('fee_amount', 0.0)):.2f}",
                                    f"{float(r.get('net_amount', 0.0)):.2f}",
                                ])
                        m_files += 1

                elif b_key == "bri":
                    by_store = {}
                    for r in txns:
                        st = str(r.get("store") or "main").lower()
                        by_store.setdefault(st, []).append(r)
                    for st, s_txns in by_store.items():
                        st_dir = b_dir / st
                        st_dir.mkdir(parents=True, exist_ok=True)
                        csv_p = st_dir / f"BRI_Cloud_{st}.csv"
                        with open(csv_p, "w", newline="", encoding="utf-8") as cf:
                            cw = csv.writer(cf)
                            cw.writerow(["TRANSACTION DATE", "REMARK_RK", "AMT_TRX", "CARD_TYPE"])
                            for r in s_txns:
                                cw.writerow([
                                    r.get("transaction_date", ""),
                                    r.get("trace_number", ""),
                                    f"{float(r.get('gross_amount', 0.0)):.2f}",
                                    r.get("card_type", "EDC")
                                ])
                        m_files += 1

        # 2. Fetch Mutation Transactions from Cloud
        mut_txns = fetch_cloud_transactions(data_type="mutation", limit=10000)
        if mut_txns:
            by_b_mut = {}
            for t in mut_txns:
                b = str(t.get("bank_name") or "BCA").lower().strip()
                acc = str(t.get("account_number") or "main").strip()
                by_b_mut.setdefault((b, acc), []).append(t)

            for (b_k, acc_num), m_rows in by_b_mut.items():
                m_t_dir = MUTATION_DIR / b_k / "main"
                m_t_dir.mkdir(parents=True, exist_ok=True)
                csv_p = m_t_dir / f"Mutation_Cloud_{b_k}.csv"
                with open(csv_p, "w", newline="", encoding="utf-8") as cf:
                    cw = csv.writer(cf)
                    cw.writerow(["Date", "Description", "Amount", "Type", "Balance"])
                    for r in m_rows:
                        cw.writerow([
                            r.get("transaction_date", ""),
                            r.get("description", ""),
                            f"{float(r.get('amount', 0.0)):.2f}",
                            r.get("mutation_type", "CR"),
                            f"{float(r.get('balance', 0.0)):.2f}"
                        ])
                mut_files += 1

        return {
            "success": True,
            "merchant_count": len(merchant_txns),
            "mutation_count": len(mut_txns),
            "merchant_files": m_files,
            "mutation_files": mut_files
        }
    except Exception as e:
        return {"success": False, "error": str(e), "merchant_files": 0, "mutation_files": 0}



