"""
reconciler.py — compares bank vs Odoo transactions by amount (multiset).

Uses collections.Counter so duplicate amounts are matched 1-to-1.
Example:
  Bank:  [500, 500, 1000]
  Odoo:  [500, 1000, 2000]
  →  500(Done), 500(Cuma ada di Bank), 1000(Done), 2000(Cuma ada di ODO)

Result dicts carry two separate number fields:
  number_odo  : reference from ODO  (Number column)
  number_bank : reference from Bank (Trace Number / REMARK_RK / Reff ID)
Both are populated for Done rows; one is empty for Bank-only / ODO-only.
"""

from collections import Counter
from decimal import Decimal


STATUS_DONE        = "Done"
STATUS_BANK_ONLY   = "Cuma ada di Bank"
STATUS_ODO_ONLY    = "Cuma ada di ODO"


def reconcile(bank_txns: list[dict], odo_txns: list[dict]) -> list[dict]:
    """
    Match bank and Odoo transactions by amount using multiset logic.

    Returns a list of result dicts with fields:
        amount, amount_raw, date, description, source, status,
        number_odo, number_bank
    """
    results = []

    # Build lookup: amount → list of odo txns (to grab number_odo on match)
    odo_by_amount: dict[Decimal, list[dict]] = {}
    for txn in odo_txns:
        odo_by_amount.setdefault(txn["amount"], []).append(txn)

    # Remaining ODO counter — decremented as matches are found
    remaining_odo = Counter(t["amount"] for t in odo_txns)
    # Track which odo txns have already been used (by index)
    odo_used_count: dict[Decimal, int] = {}

    # ── Bank transactions ──────────────────────────────────────────────────────
    for txn in bank_txns:
        amt = txn["amount"]
        bank_num = txn.get("number", "")

        if remaining_odo[amt] > 0:
            remaining_odo[amt] -= 1
            status = STATUS_DONE

            # Grab the matched ODO txn's number (FIFO)
            used = odo_used_count.get(amt, 0)
            candidates = odo_by_amount.get(amt, [])
            odo_num = candidates[used]["number"] if used < len(candidates) else ""
            odo_used_count[amt] = used + 1
        else:
            status = STATUS_BANK_ONLY
            odo_num = ""

        results.append({
            "amount":      amt,
            "amount_raw":  txn.get("amount_raw", ""),
            "date":        txn.get("date", ""),
            "description": txn.get("description", ""),
            "source":      "Bank",
            "status":      status,
            "number_odo":  odo_num,
            "number_bank": bank_num,
        })

    # ── Odoo-only transactions ─────────────────────────────────────────────────
    for amt, count in remaining_odo.items():
        if count <= 0:
            continue
        candidates = odo_by_amount.get(amt, [])
        # The unmatched ones are the last `count` that were never consumed
        used = odo_used_count.get(amt, 0)
        unmatched = candidates[used:used + count]

        for txn in unmatched:
            results.append({
                "amount":      amt,
                "amount_raw":  txn.get("amount_raw", ""),
                "date":        txn.get("date", ""),
                "description": txn.get("description", ""),
                "source":      "Odoo",
                "status":      STATUS_ODO_ONLY,
                "number_odo":  txn.get("number", ""),
                "number_bank": "",
            })

        # Safety: fill if somehow count > available candidates
        for _ in range(count - len(unmatched)):
            results.append({
                "amount":      amt,
                "amount_raw":  str(amt),
                "date":        "",
                "description": "",
                "source":      "Odoo",
                "status":      STATUS_ODO_ONLY,
                "number_odo":  "",
                "number_bank": "",
            })

    return results


def summary(results: list[dict]) -> dict:
    done      = sum(1 for r in results if r["status"] == STATUS_DONE)
    bank_only = sum(1 for r in results if r["status"] == STATUS_BANK_ONLY)
    odo_only  = sum(1 for r in results if r["status"] == STATUS_ODO_ONLY)
    return {"done": done, "bank_only": bank_only, "odo_only": odo_only, "total": len(results)}
