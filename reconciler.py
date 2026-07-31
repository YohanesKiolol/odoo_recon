"""
reconciler.py — compares bank vs Odoo transactions by (date, amount).

Matching logic:
  1. Try to match bank txn to an ODO txn with same date AND amount (strict).
  2. If no ODO txn on that date → Cuma ada di Bank.
  Remaining unmatched ODO txns → Cuma ada di ODO.

This prevents a bank transaction on July 6 from matching an ODO transaction
on July 10 just because they share the same amount.

Result dicts carry two separate number fields:
  number_odo  : reference from ODO  (Number column)
  number_bank : reference from Bank (Trace Number / REMARK_RK / Reff ID)
Both are populated for Done rows; one is empty for Bank-only / ODO-only.
"""

from collections import Counter, defaultdict
from decimal import Decimal


STATUS_DONE        = "Done"
STATUS_BANK_ONLY   = "Cuma ada di Bank"
STATUS_ODO_ONLY    = "Cuma ada di ODO"


def reconcile(bank_txns: list[dict], odo_txns: list[dict]) -> list[dict]:
    """
    Match bank and Odoo transactions by (date, amount) — date-aware multiset.

    Returns a list of result dicts with fields:
        amount, amount_raw, date, description, source, status,
        number_odo, number_bank, filename_bank
    """
    results = []

    # ── Build per-date ODO index ──────────────────────────────────────────────
    # odo_by_date[date][amount] = [txn, txn, ...]
    odo_by_date: dict[str, dict[Decimal, list[dict]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for txn in odo_txns:
        d = txn.get("date", "") or ""
        odo_by_date[d][txn["amount"]].append(txn)

    # Counter per date to track remaining matches
    odo_remaining: dict[str, Counter] = {
        d: Counter({amt: len(lst) for amt, lst in amts.items()})
        for d, amts in odo_by_date.items()
    }
    # Track consumed index per (date, amount) for number_odo lookup
    odo_used: dict[str, dict[Decimal, int]] = defaultdict(lambda: defaultdict(int))

    # ── Bank transactions ─────────────────────────────────────────────────────
    for txn in bank_txns:
        amt      = txn["amount"]
        bank_d   = txn.get("date", "") or ""
        bank_num = txn.get("number", "")

        day_counter = odo_remaining.get(bank_d, Counter())

        if day_counter[amt] > 0:
            day_counter[amt] -= 1
            status = STATUS_DONE

            # Grab the matched ODO txn's number (FIFO)
            used       = odo_used[bank_d][amt]
            candidates = odo_by_date.get(bank_d, {}).get(amt, [])
            odo_num    = candidates[used]["number"] if used < len(candidates) else ""
            odo_used[bank_d][amt] = used + 1
        else:
            status  = STATUS_BANK_ONLY
            odo_num = ""

        results.append({
            "amount":        amt,
            "amount_raw":    txn.get("amount_raw", ""),
            "date":          bank_d,
            "description":   txn.get("description", ""),
            "source":        "Bank",
            "status":        status,
            "number_odo":    odo_num,
            "number_bank":   bank_num,
            "filename_bank": txn.get("filename", ""),
        })

    # ── Odoo-only transactions ────────────────────────────────────────────────
    for d, day_counter in odo_remaining.items():
        for amt, count in day_counter.items():
            if count <= 0:
                continue
            candidates = odo_by_date.get(d, {}).get(amt, [])
            used       = odo_used[d][amt]
            unmatched  = candidates[used:used + count]

            for txn in unmatched:
                results.append({
                    "amount":        amt,
                    "amount_raw":    txn.get("amount_raw", ""),
                    "date":          txn.get("date", ""),
                    "description":   txn.get("description", ""),
                    "source":        "Odoo",
                    "status":        STATUS_ODO_ONLY,
                    "number_odo":    txn.get("number", ""),
                    "number_bank":   "",
                    "filename_bank": "",
                })

            # Safety: fill if count > available candidates
            for _ in range(count - len(unmatched)):
                results.append({
                    "amount":        amt,
                    "amount_raw":    str(amt),
                    "date":          d,
                    "description":   "",
                    "source":        "Odoo",
                    "status":        STATUS_ODO_ONLY,
                    "number_odo":    "",
                    "number_bank":   "",
                    "filename_bank": "",
                })

    return results


def summary(results: list[dict]) -> dict:
    done      = sum(1 for r in results if r["status"] == STATUS_DONE)
    bank_only = sum(1 for r in results if r["status"] == STATUS_BANK_ONLY)
    odo_only  = sum(1 for r in results if r["status"] == STATUS_ODO_ONLY)
    return {"done": done, "bank_only": bank_only, "odo_only": odo_only, "total": len(results)}
