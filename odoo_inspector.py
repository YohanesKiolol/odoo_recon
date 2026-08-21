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
    ODOO_URL = os.environ.get("ODOO_URL", "")
    ODOO_DB = os.environ.get("ODOO_DB", "")
    ODOO_API_KEY = os.environ.get("ODOO_API_KEY", "")
    PREDEFINED_ACCOUNTS = {}

DEFAULT_ODOO_URL = ""
DEFAULT_ODOO_DB = ""

_cached_auth: dict[str, Any] = {
    "url": None,
    "db": None,
    "uid": None,
    "pwd": None
}


def _get_base_url() -> str:
    """Extract scheme + netloc base URL from ODOO_URL."""
    raw = (ODOO_URL or DEFAULT_ODOO_URL or "").strip()
    if not raw:
        return ""
    parsed = urllib.parse.urlparse(raw)
    if parsed.scheme and parsed.netloc:
        return f"{parsed.scheme}://{parsed.netloc}"
    return raw.rstrip("/")


_active_credentials: dict[str, str] = {
    "username": "",
    "password": ""
}


def set_active_credentials(username: str, password: str) -> None:
    """Set active credentials selected in GUI and reset cached session."""
    global _active_credentials, _cached_auth
    u = str(username or "").strip()
    p = str(password or "").strip()
    if _active_credentials.get("username") != u or _active_credentials.get("password") != p:
        _active_credentials["username"] = u
        _active_credentials["password"] = p
        _cached_auth = {
            "url": None,
            "db": None,
            "uid": None,
            "pwd": None
        }


def _get_credentials() -> tuple[str, str, str, str]:
    """Retrieve url, db, username, password/api_key for Odoo RPC authentication."""
    url = _get_base_url()
    db = (ODOO_DB or DEFAULT_ODOO_DB or "").strip()

    username = _active_credentials.get("username", "")
    password = _active_credentials.get("password", "")

    if not username and PREDEFINED_ACCOUNTS:
        first_acc = next(iter(PREDEFINED_ACCOUNTS.values()))
        username = first_acc.get("username", "")
        password = first_acc.get("api_key") or first_acc.get("password", "")

    if not username:
        username = os.environ.get("ODOO_USER", "")
    if not password:
        password = os.environ.get("ODOO_API_KEY", os.environ.get("ODOO_PASSWORD", ""))

    if not _active_credentials.get("password") and ODOO_API_KEY:
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


def get_odoo_user_profile() -> dict:

    """Retrieve full name, login, email of the authenticated user from Odoo server."""
    try:
        uid, err = authenticate()
        if not uid:
            return {}
        info = _execute_kw('res.users', 'read', [[uid]], {'fields': ['name', 'login', 'email']})
        if info and isinstance(info, list) and len(info) > 0:
            return info[0]
    except Exception:
        pass
    return {}


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
def inspect_bank_only(amount: float, date_str: str, bank_name: str = "", journal_name: str = "") -> dict:
    """
    Diagnose '🏦 Bank Only' discrepancy.
    Rule:
      1. Search invoices on the exact date and amount (Draft or Unpaid).
      2. Search account.payment on date (+/-1 day) for exact amount to detect:
         - 🔄 Cross-Journal Misclassification (e.g. recorded under Mandiri/Cash instead of BCA).
         - 💳 Existing Odoo payments in the same journal.
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
            "invoices": [],
            "cross_journal_payments": [],
            "same_journal_payments": []
        }

    # 1. Check account.move for candidate draft/unpaid invoices
    inv_domain = [
        ('move_type', '=', 'out_invoice'),
        ('invoice_date', '=', t_date),
        ('amount_total', '>=', round(clean_amt - 0.01, 2)),
        ('amount_total', '<=', round(clean_amt + 0.01, 2)),
    ]

    candidates = []
    try:
        invoices = _execute_kw(
            'account.move',
            'search_read',
            [inv_domain],
            {
                'fields': ['name', 'invoice_date', 'amount_total', 'payment_state', 'state', 'partner_id'],
                'limit': 25
            }
        )
        for inv in (invoices or []):
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
    except Exception:
        invoices = []

    # 2. Check account.payment across ALL journals (date +/- 1 day)
    from datetime import datetime as _dt, timedelta as _tmd
    try:
        dt_obj = _dt.strptime(t_date, "%Y-%m-%d").date()
        d_min = (dt_obj - _tmd(days=1)).strftime("%Y-%m-%d")
        d_max = (dt_obj + _tmd(days=1)).strftime("%Y-%m-%d")
    except Exception:
        d_min = t_date
        d_max = t_date

    pay_domain = [
        ('date', '>=', d_min),
        ('date', '<=', d_max),
        ('amount', '>=', round(clean_amt - 0.01, 2)),
        ('amount', '<=', round(clean_amt + 0.01, 2)),
    ]

    cross_journal_payments = []
    same_journal_payments = []

    try:
        payments = _execute_kw(
            'account.payment',
            'search_read',
            [pay_domain],
            {
                'fields': ['name', 'date', 'amount', 'journal_id', 'state', 'partner_id', 'ref'],
                'limit': 25
            }
        )
        from config import get_journal_store
        expected_store = get_journal_store(journal_name or bank_name)

        for p in (payments or []):
            j_info = p.get('journal_id') or [0, "Unknown Journal"]
            actual_jrn_name = j_info[1] if isinstance(j_info, (list, tuple)) and len(j_info) > 1 else str(j_info)
            act_jrn_upper = actual_jrn_name.upper()

            is_same = False
            if norm_expected_jrn and norm_expected_jrn in act_jrn_upper:
                is_same = True
            elif norm_expected_bank and norm_expected_bank in act_jrn_upper:
                is_same = True

            actual_store = get_journal_store(actual_jrn_name)

            # Store Location Isolation: Cross-journal must belong to the same physical store
            if expected_store and actual_store and expected_store != actual_store:
                continue

            p_partner = p['partner_id'][1] if p.get('partner_id') else "Walk-in Customer"
            p_dict = {
                "id": p['id'],
                "name": p.get('name', '-'),
                "date": p.get('date', '-'),
                "amount": float(p.get('amount', 0.0)),
                "actual_journal": actual_jrn_name,
                "expected_journal": journal_name or bank_name,
                "store": actual_store or expected_store,
                "state": p.get('state', 'posted'),
                "customer": p_partner,
                "ref": p.get('ref') or "-",
                "is_same_journal": is_same
            }

            if not is_same:
                cross_journal_payments.append(p_dict)
            else:
                same_journal_payments.append(p_dict)
    except Exception:
        pass

    summary_parts = []
    if cross_journal_payments:
        j_names = ", ".join(sorted(set(p['actual_journal'] for p in cross_journal_payments)))
        summary_parts.append(f"🚨 Possible Wrong Journal: Found {len(cross_journal_payments)} payment(s) in Odoo recorded under '{j_names}' matching Rp {clean_amt:,.2f}.")
    if candidates:
        summary_parts.append(f"Found {len(candidates)} candidate draft/unpaid invoice(s) in Odoo matching Rp {clean_amt:,.2f}.")
    if not summary_parts:
        if same_journal_payments:
            summary_parts.append(f"Found existing payment in '{same_journal_payments[0]['actual_journal']}' ({same_journal_payments[0]['name']}), but not reconciled.")
        else:
            summary_parts.append(f"No draft/unpaid invoices or cross-journal payments found in Odoo on {t_date} matching Rp {clean_amt:,.2f}.")

    return {
        "success": True,
        "discrepancy_type": "bank_only",
        "date": t_date,
        "amount": clean_amt,
        "found_count": len(candidates),
        "invoices": candidates,
        "cross_journal_payments": cross_journal_payments,
        "same_journal_payments": same_journal_payments,
        "summary": " ".join(summary_parts)
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
    b_name = str(item.get("bank", item.get("bank_name", ""))).strip()
    j_name = str(item.get("journal", "")).strip()

    if dtype == "bank_only":
        return inspect_bank_only(amount=amt, date_str=t_date, bank_name=b_name, journal_name=j_name)
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
