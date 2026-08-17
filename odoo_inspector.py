from __future__ import annotations

import os
import re
import xmlrpc.client
import urllib.parse
from datetime import datetime, date
from typing import Any

# Import configs
try:
    from config import ODOO_URL, ODOO_DB, ODOO_API_KEY, PREDEFINED_ACCOUNTS
except Exception:
    ODOO_URL = os.environ.get("ODOO_URL", "https://eyerizz.raytech.id")
    ODOO_DB = os.environ.get("ODOO_DB", "production")
    ODOO_API_KEY = os.environ.get("ODOO_API_KEY", "")
    PREDEFINED_ACCOUNTS = {}

DEFAULT_ODOO_URL = "https://eyerizz.raytech.id"
DEFAULT_ODOO_DB = "production"

_cached_auth: dict[str, Any] = {
    "url": None,
    "db": None,
    "uid": None,
    "pwd": None
}


def _get_base_url() -> str:
    """Extract scheme + netloc base URL from ODOO_URL."""
    raw = ODOO_URL or DEFAULT_ODOO_URL
    parsed = urllib.parse.urlparse(raw)
    if parsed.scheme and parsed.netloc:
        return f"{parsed.scheme}://{parsed.netloc}"
    return raw.rstrip("/")


def _get_credentials() -> tuple[str, str, str, str]:
    """Retrieve url, db, username, password/api_key for Odoo RPC authentication."""
    url = _get_base_url()
    db = ODOO_DB or DEFAULT_ODOO_DB

    username = ""
    password = ""

    if PREDEFINED_ACCOUNTS:
        first_acc = next(iter(PREDEFINED_ACCOUNTS.values()))
        username = first_acc.get("username", "")
        password = first_acc.get("password", "")

    if not username:
        username = os.environ.get("ODOO_USER", "fransisca")
    if not password:
        password = os.environ.get("ODOO_PASSWORD", "250201")

    if ODOO_API_KEY:
        password = ODOO_API_KEY.strip()

    return url, db, username, password


def authenticate(force_refresh: bool = False) -> tuple[int | None, str | None]:
    """
    Authenticate with Odoo via XML-RPC.
    Returns (uid, error_message).
    """
    global _cached_auth
    url, db, username, password = _get_credentials()

    if not force_refresh and _cached_auth.get("uid") and _cached_auth.get("url") == url and _cached_auth.get("db") == db:
        return _cached_auth["uid"], None

    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        uid = common.authenticate(db, username, password, {})
        if uid:
            _cached_auth = {
                "url": url,
                "db": db,
                "uid": uid,
                "pwd": password
            }
            return uid, None
        return None, f"Invalid Odoo credentials for user '{username}' on database '{db}'"
    except Exception as e:
        return None, f"Odoo Connection Error: {str(e)}"


_cached_models_proxy = None

def _execute_kw(model: str, method: str, args: list, kwargs: dict | None = None) -> Any:
    """Execute a read-only RPC call on an Odoo model."""
    global _cached_models_proxy
    uid, err = authenticate()
    if not uid:
        raise ConnectionError(err or "Failed to authenticate with Odoo")

    url, db, _, password = _get_credentials()
    if _cached_models_proxy is None or getattr(_cached_models_proxy, "_url", "") != f"{url}/xmlrpc/2/object":
        _cached_models_proxy = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
        _cached_models_proxy._url = f"{url}/xmlrpc/2/object"
    return _cached_models_proxy.execute_kw(db, uid, password, model, method, args, kwargs or {})


def _normalize_date_to_iso(val: Any) -> str:
    """Safely convert any date string, datetime, or date object to ISO 'YYYY-MM-DD'."""
    if not val:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")

    s = str(val).strip()
    for fmt in [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%m/%d/%Y",
        "%Y/%m/%d",
        "%d.%m.%Y",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    ]:
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        return s[:10]
    return s


# =============================================================================
# 1. 🏦 Bank Only — Possible Candidate Invoices (Draft or Unpaid on that date)
# =============================================================================
def inspect_bank_only(amount: float, date_str: str) -> dict:
    """
    Diagnose '🏦 Bank Only' discrepancy.
    Rule: Search invoices on the exact date and amount.
    Filter criteria: Only return possible candidates:
      - 🟡 Invoices still in Draft (state == 'draft')
      - 🟢 Invoices posted but unpaid (state == 'posted' and payment_state == 'not_paid')
    Already-paid / linked invoices are filtered out.
    """
    try:
        clean_amt = float(amount)
    except (ValueError, TypeError):
        clean_amt = 0.0

    t_date = _normalize_date_to_iso(date_str)
    if not t_date or clean_amt <= 0:
        return {
            "success": False,
            "error": "Invalid date or amount provided.",
            "invoices": []
        }

    domain = [
        ('move_type', '=', 'out_invoice'),
        ('invoice_date', '=', t_date),
        ('amount_total', '>=', round(clean_amt - 0.01, 2)),
        ('amount_total', '<=', round(clean_amt + 0.01, 2)),
    ]

    try:
        invoices = _execute_kw(
            'account.move',
            'search_read',
            [domain],
            {
                'fields': ['name', 'invoice_date', 'amount_total', 'payment_state', 'state', 'partner_id'],
                'limit': 25
            }
        )
    except Exception as e:
        return {
            "success": False,
            "error": f"Odoo Query Error: {str(e)}",
            "invoices": []
        }

    candidates = []
    for inv in invoices:
        state = inv.get('state', '')
        pay_state = inv.get('payment_state', '')
        p_name = inv['partner_id'][1] if inv.get('partner_id') else "Walk-in Customer"

        if state == 'draft':
            candidates.append({
                "id": inv['id'],
                "name": inv['name'],
                "customer": p_name,
                "partner": p_name,
                "amount": inv['amount_total'],
                "date": inv['invoice_date'],
                "status_code": "DRAFT_INVOICE",
                "badge": "🟡 Draft Invoice",
                "state": "Draft",
                "detail": f"Invoice {inv['name']} is in Draft state (Needs to be posted in Odoo)"
            })
        elif state == 'posted' and pay_state == 'not_paid':
            candidates.append({
                "id": inv['id'],
                "name": inv['name'],
                "customer": p_name,
                "partner": p_name,
                "amount": inv['amount_total'],
                "date": inv['invoice_date'],
                "status_code": "AVAILABLE_OPEN",
                "badge": "🟢 Unpaid Open Invoice",
                "state": "Unpaid",
                "detail": f"Invoice {inv['name']} is posted & unpaid (Ready to match with bank payment)"
            })

    if not candidates:
        summary = f"No draft or unpaid invoices found in Odoo on {t_date} matching Rp {clean_amt:,.2f}."
    else:
        summary = f"Found {len(candidates)} candidate invoice(s) on {t_date} matching Rp {clean_amt:,.2f}."

    return {
        "success": True,
        "discrepancy_type": "bank_only",
        "date": t_date,
        "amount": clean_amt,
        "found_count": len(candidates),
        "invoices": candidates,
        "summary": summary
    }


# =============================================================================
# 2. 📦 Odoo Only — Linked Invoice Number & Details
# =============================================================================
def inspect_odoo_only(odoo_number: str, amount: float = 0.0) -> dict:
    """
    Diagnose '📦 Odoo Only' discrepancy.
    Rule: Finds the exact Invoice Number that this payment was linked to in Odoo.
    """
    doc_num = str(odoo_number).strip()
    if not doc_num:
        return {
            "success": False,
            "error": "No Odoo payment reference number provided.",
            "payment": None,
            "linked_invoices": []
        }

    domain = [('name', '=', doc_num)]
    try:
        payments = _execute_kw(
            'account.payment',
            'search_read',
            [domain],
            {
                'fields': ['name', 'date', 'amount', 'state', 'partner_id', 'journal_id', 'ref', 'reconciled_invoice_ids', 'move_id'],
                'limit': 1
            }
        )
    except Exception as e:
        return {
            "success": False,
            "error": f"Odoo Query Error: {str(e)}",
            "payment": None,
            "linked_invoices": []
        }

    if not payments:
        return {
            "success": True,
            "discrepancy_type": "odoo_only",
            "found": False,
            "odoo_number": doc_num,
            "payment": None,
            "linked_invoices": [],
            "summary": f"Payment document '{doc_num}' was not found in Odoo."
        }

    p = payments[0]
    p_partner = p['partner_id'][1] if p.get('partner_id') else "No Customer"
    p_journal = p['journal_id'][1] if p.get('journal_id') else "No Journal"
    p_state = str(p.get('state', '')).capitalize()

    # Find linked invoice(s)
    linked_invoices = []
    if p.get('reconciled_invoice_ids'):
        try:
            invs = _execute_kw(
                'account.move',
                'search_read',
                [[('id', 'in', p['reconciled_invoice_ids'])]],
                {'fields': ['name', 'invoice_date', 'amount_total', 'state', 'payment_state', 'partner_id']}
            )
            linked_invoices.extend(invs)
        except Exception:
            pass

    if not linked_invoices:
        inv_name_cand = None
        if p.get('ref') and 'INV/' in p['ref']:
            m = re.search(r'INV/\d+/\d+', p['ref'])
            if m:
                inv_name_cand = m.group(0)
        elif p.get('move_id') and len(p['move_id']) > 1:
            m = re.search(r'INV/\d+/\d+', p['move_id'][1])
            if m:
                inv_name_cand = m.group(0)

        if inv_name_cand:
            try:
                invs = _execute_kw(
                    'account.move',
                    'search_read',
                    [[('name', '=', inv_name_cand)]],
                    {'fields': ['name', 'invoice_date', 'amount_total', 'state', 'payment_state', 'partner_id']}
                )
                linked_invoices.extend(invs)
            except Exception:
                pass

    formatted_invoices = []
    for li in linked_invoices:
        c_name = li['partner_id'][1] if li.get('partner_id') else p_partner
        st = str(li.get('state', '')).capitalize()
        pst = str(li.get('payment_state', '')).capitalize()
        formatted_invoices.append({
            "name": li['name'],
            "invoice_number": li['name'],
            "date": li.get('invoice_date'),
            "customer": c_name,
            "amount": li.get('amount_total', 0.0),
            "state": st,
            "payment_state": pst,
            "detail": f"Invoice: {li['name']} • Date: {li.get('invoice_date')} • Customer: {c_name} • Total: Rp {li.get('amount_total', 0.0):,.2f}"
        })

    if formatted_invoices:
        inv_nums = ", ".join(i['name'] for i in formatted_invoices)
        summary = f"Payment {p['name']} is linked to Invoice: {inv_nums}."
    else:
        summary = f"Payment {p['name']} has no linked invoice in Odoo."

    return {
        "success": True,
        "discrepancy_type": "odoo_only",
        "found": True,
        "id": p['id'],
        "name": p['name'],
        "date": p.get('date'),
        "amount": p.get('amount', 0.0),
        "state": p_state,
        "partner": p_partner,
        "journal": p_journal,
        "linked_invoices": formatted_invoices,
        "summary": summary
    }


# =============================================================================
# 3. ⚠️ Unreconciled — Referenced Invoice & Other Payment Linkage Check
# =============================================================================
def inspect_unreconciled(odoo_number: str, amount: float, date_str: str) -> dict:
    """
    Diagnose '⚠️ Unreconciled' discrepancy.
    Rule:
      1. Shows the referenced/linked invoice information (Invoice Number, Date, Customer, Amount, Status).
      2. Checks if that referenced invoice is already linked to another payment in Odoo.
         If yes, displays that other payment's Number, Date, Journal, and Amount.
    """
    doc_num = str(odoo_number).strip()
    t_date = _normalize_date_to_iso(date_str)
    try:
        clean_amt = float(amount)
    except (ValueError, TypeError):
        clean_amt = 0.0

    # 1. Fetch payment info
    payment_info = None
    p = None
    if doc_num:
        try:
            payments = _execute_kw(
                'account.payment',
                'search_read',
                [[('name', '=', doc_num)]],
                {
                    'fields': ['name', 'date', 'amount', 'state', 'partner_id', 'journal_id', 'ref', 'reconciled_invoice_ids', 'move_id'],
                    'limit': 1
                }
            )
            if payments:
                p = payments[0]
                payment_info = {
                    "id": p['id'],
                    "name": p['name'],
                    "date": p.get('date'),
                    "amount": p.get('amount', 0.0),
                    "state": str(p.get('state', '')).capitalize(),
                    "partner": p['partner_id'][1] if p.get('partner_id') else "No Customer",
                    "journal": p['journal_id'][1] if p.get('journal_id') else "No Journal",
                }
        except Exception:
            pass

    # 2. Extract referenced invoice names from payment
    inv_names = []
    if p:
        if p.get('ref') and 'INV/' in p['ref']:
            m = re.findall(r'INV/\d+/\d+', p['ref'])
            inv_names.extend(m)
        if p.get('move_id') and len(p['move_id']) > 1:
            m = re.findall(r'INV/\d+/\d+', p['move_id'][1])
            inv_names.extend(m)

    found_invoices = []
    if inv_names:
        try:
            found_invoices = _execute_kw(
                'account.move',
                'search_read',
                [[('name', 'in', list(set(inv_names)))]],
                {'fields': ['name', 'invoice_date', 'amount_total', 'payment_state', 'state', 'partner_id', 'invoice_payments_widget']}
            )
        except Exception:
            pass

    # Fallback: search by date and amount if no explicit reference on payment
    if not found_invoices and t_date:
        try:
            found_invoices = _execute_kw(
                'account.move',
                'search_read',
                [[('move_type', '=', 'out_invoice'), ('invoice_date', '=', t_date), ('amount_total', '>=', round(clean_amt - 0.01, 2)), ('amount_total', '<=', round(clean_amt + 0.01, 2))]],
                {'fields': ['name', 'invoice_date', 'amount_total', 'payment_state', 'state', 'partner_id', 'invoice_payments_widget'], 'limit': 5}
            )
        except Exception:
            pass

    invoice_cards = []
    has_other_payment = False

    for inv in found_invoices:
        p_name = inv['partner_id'][1] if inv.get('partner_id') else "Walk-in Customer"
        widget = inv.get('invoice_payments_widget') or {}
        content = widget.get('content', []) if isinstance(widget, dict) else []

        other_payments = []
        for pay in content:
            p_ref = pay.get('ref') or pay.get('name') or f"Payment #{pay.get('account_payment_id')}"
            # Check if this is a different payment from the current unreconciled payment
            if doc_num not in p_ref:
                other_payments.append({
                    "ref": p_ref,
                    "date": str(pay.get('date') or "").strip(),
                    "journal": str(pay.get('journal_name') or "").strip(),
                    "amount": float(pay.get('amount') or 0.0)
                })

        is_already_linked = len(other_payments) > 0
        if is_already_linked:
            has_other_payment = True
            badge = "🔴 Linked to Another Payment"
            status_code = "ALREADY_LINKED"
        elif inv.get('state') == 'draft':
            badge = "🟡 Draft Invoice"
            status_code = "DRAFT_INVOICE"
        elif inv.get('state') == 'posted' and inv.get('payment_state') == 'not_paid':
            badge = "🟢 Unpaid Open Invoice"
            status_code = "AVAILABLE_OPEN"
        else:
            badge = f"⚪ {str(inv.get('state', '')).capitalize()}"
            status_code = "OTHER"

        invoice_cards.append({
            "name": inv['name'],
            "invoice_number": inv['name'],
            "date": inv.get('invoice_date'),
            "customer": p_name,
            "amount": inv.get('amount_total', 0.0),
            "state": inv.get('state'),
            "payment_state": inv.get('payment_state'),
            "badge": badge,
            "status_code": status_code,
            "other_payments": other_payments
        })

    if not invoice_cards:
        summary = f"No referenced invoice found in Odoo for payment {doc_num}."
    elif has_other_payment:
        summary = f"Referenced invoice is already linked to another customer payment in Odoo."
    else:
        summary = f"Found referenced invoice in Odoo ({len(invoice_cards)} invoice)."

    return {
        "success": True,
        "discrepancy_type": "unreconciled_odoo",
        "payment": payment_info,
        "date": t_date,
        "amount": clean_amt,
        "invoices_found": len(invoice_cards),
        "invoices": invoice_cards,
        "has_other_payment": has_other_payment,
        "summary": summary
    }


def inspect_discrepancy(item: dict) -> dict:
    """
    Unified entry point to inspect any discrepancy item dictionary.
    """
    dtype = str(item.get("discrepancy_type") or "bank_only").strip()
    amt = float(item.get("amount", 0.0))
    t_date = str(item.get("transaction_date", item.get("date", ""))).strip()
    o_num = str(item.get("odoo_number", item.get("number_odo", item.get("invoice_no", "")))).strip()

    if dtype == "bank_only":
        return inspect_bank_only(amount=amt, date_str=t_date)
    elif dtype == "odoo_only":
        return inspect_odoo_only(odoo_number=o_num, amount=amt)
    elif dtype == "unreconciled_odoo":
        return inspect_unreconciled(odoo_number=o_num, amount=amt, date_str=t_date)
    else:
        return {
            "success": True,
            "summary": f"Discrepancy type '{dtype}' requires manual review.",
            "invoices": []
        }
