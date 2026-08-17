from __future__ import annotations

import json
import os
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime, date
from pathlib import Path
from decimal import Decimal

# Default Supabase configuration for standalone Sales Portal (no .env needed)
DEFAULT_SUPABASE_URL = "https://drauzroctexgfqxouswt.supabase.co"
DEFAULT_SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRyYXV6cm9jdGV4Z2ZxeG91c3d0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODY3MjU4NzUsImV4cCI6MjEwMjMwMTg3NX0.DLQ7WKb4ykshyjS1a_dHggQCM3vFpRDjxf0SXB77-qc"
DEFAULT_COMPANY_NAME = "Eyerizz Eyewear"

try:
    from config import SUPABASE_URL as _CFG_URL, SUPABASE_KEY as _CFG_KEY, ODOO_COMPANY_NAME as _CFG_COMPANY
    SUPABASE_URL = _CFG_URL if _CFG_URL else DEFAULT_SUPABASE_URL
    SUPABASE_KEY = _CFG_KEY if _CFG_KEY else DEFAULT_SUPABASE_KEY
    ODOO_COMPANY_NAME = _CFG_COMPANY if _CFG_COMPANY else DEFAULT_COMPANY_NAME
except Exception:
    SUPABASE_URL = os.environ.get("SUPABASE_URL", DEFAULT_SUPABASE_URL)
    SUPABASE_KEY = os.environ.get("SUPABASE_KEY", DEFAULT_SUPABASE_KEY)
    ODOO_COMPANY_NAME = os.environ.get("ODOO_COMPANY_NAME", DEFAULT_COMPANY_NAME)


def is_cloud_configured() -> bool:
    """Check if Supabase URL and Key are configured."""
    return bool(SUPABASE_URL and SUPABASE_KEY and "your-project" not in SUPABASE_URL)


def supabase_auth_login(email: str, password: str) -> tuple[bool, dict | str]:
    """Authenticate with Supabase Auth using email and password."""
    if not is_cloud_configured():
        return False, "Supabase is not configured"

    base_url = SUPABASE_URL.rstrip("/")
    url = f"{base_url}/auth/v1/token?grant_type=password"
    headers = {
        "apikey": SUPABASE_KEY,
        "Content-Type": "application/json"
    }
    payload = {
        "email": email.strip(),
        "password": password.strip()
    }
    body_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=body_bytes, headers=headers, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return True, data
    except urllib.error.HTTPError as e:
        try:
            err_data = json.loads(e.read().decode("utf-8"))
            msg = err_data.get("msg") or err_data.get("error_description") or err_data.get("message") or f"Invalid login credentials (code {e.code})"
            return False, msg
        except Exception:
            return False, f"Login failed (HTTP {e.code})"
    except Exception as e:
        return False, str(e)


def _api_headers(access_token: str | None = None) -> dict[str, str]:
    """Build standard headers for Supabase PostgREST requests."""
    token = access_token if access_token else SUPABASE_KEY
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _make_request(endpoint: str, method: str = "GET", data: dict | list | None = None, headers: dict | None = None, timeout: int = 15):
    """Execute an HTTP request to Supabase PostgREST endpoint."""
    if not is_cloud_configured():
        raise ConnectionError("Supabase is not configured. Please set SUPABASE_URL and SUPABASE_KEY in .env.")

    base_url = SUPABASE_URL.rstrip("/")
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
    """Test connection to the Supabase bank_discrepancies table."""
    if not is_cloud_configured():
        return False, "SUPABASE_URL or SUPABASE_KEY not set in .env"
    try:
        _make_request("/bank_discrepancies?limit=1", method="GET", timeout=6)
        return True, "Connected successfully"
    except Exception as e:
        return False, str(e)


def push_bank_discrepancies(discrepancies: list[dict], recon_date: str = "", company: str = "") -> dict:
    """
    Push 'Only in Bank' transactions to Supabase bank_discrepancies table.
    Uses upsert with on_conflict merge to avoid duplicates on multiple syncs.
    """
    if not is_cloud_configured():
        return {"success": False, "error": "Cloud not configured in .env", "count": 0}

    if not discrepancies:
        return {"success": True, "count": 0, "message": "No discrepancies to upload"}

    company_name = company or ODOO_COMPANY_NAME or "Eyerizz Eyewear"
    
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

        records.append({
            "recon_date": r_date,
            "company": company_name,
            "bank_name": str(item.get("bank", item.get("bank_name", ""))).upper(),
            "journal": str(item.get("journal", "")).strip(),
            "transaction_date": t_date,
            "bank_number": b_num,
            "odoo_number": o_num,
            "odoo_reference": o_ref,
            "is_reconciled": is_recon,
            "discrepancy_type": disc_type,
            "filename": str(item.get("filename", item.get("filename_bank", ""))).strip(),
            "amount": amt,
            "status": "pending_sales",
            "created_at": now_iso,
            "updated_at": now_iso,
        })

    # PostgREST Upsert header: resolution=merge-duplicates
    headers = {
        "Prefer": "resolution=merge-duplicates,return=representation"
    }

    total_synced = 0
    chunk_size = 150
    try:
        for i in range(0, len(records), chunk_size):
            chunk = records[i:i + chunk_size]
            resp = _make_request(
                "/bank_discrepancies?on_conflict=recon_date,bank_name,transaction_date,bank_number,odoo_number,discrepancy_type,amount",
                method="POST",
                data=chunk,
                headers=headers
            )
            total_synced += len(resp) if isinstance(resp, list) else len(chunk)
        return {"success": True, "count": total_synced}
    except Exception as e:
        return {"success": False, "error": str(e), "count": total_synced}


def fetch_discrepancies(
    status: str | None = None,
    bank: str | None = None,
    recon_date: str | None = None,
    limit: int = 200,
) -> list[dict]:
    """Fetch discrepancies from cloud with optional filters."""
    if not is_cloud_configured():
        return []

    params = ["order=bank_name.asc,transaction_date.asc,id.desc", f"limit={limit}"]

    if status and status.lower() != "all":
        params.append(f"status=eq.{urllib.parse.quote(status)}")
    if bank and bank.lower() != "all":
        params.append(f"bank_name=eq.{urllib.parse.quote(bank.upper())}")
    if recon_date:
        params.append(f"recon_date=eq.{urllib.parse.quote(recon_date)}")

    query_str = "&".join(params)
    endpoint = f"/bank_discrepancies?{query_str}"

    try:
        res = _make_request(endpoint, method="GET")
        return res if isinstance(res, list) else []
    except Exception as e:
        print(f"[CloudSync] Fetch failed: {e}")
        return []


def resolve_discrepancy(
    item_id: int,
    sales_person: str,
    action_type: str,
    sales_notes: str,
    odoo_reference: str = "",
) -> bool:
    """Update a discrepancy record as resolved by the Sales team."""
    if not is_cloud_configured():
        return False

    payload = {
        "status": "resolved_by_sales",
        "sales_person": sales_person.strip(),
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
    """Reset discrepancy status back to pending_sales."""
    if not is_cloud_configured():
        return False

    payload = {
        "status": "pending_sales",
        "updated_at": datetime.now().isoformat(),
    }

    try:
        _make_request(f"/bank_discrepancies?id=eq.{item_id}", method="PATCH", data=payload)
        return True
    except Exception as e:
        print(f"[CloudSync] Reopen failed for id={item_id}: {e}")
        return False
