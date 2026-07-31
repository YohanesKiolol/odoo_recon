"""
amount_utils.py — robust amount parsing that handles mixed formats.

Supported inputs:
  "1.500.000"        → Indonesian thousand-separated, no decimal
  "1.500.000,50"     → Indonesian with decimal comma
  "1,500,000"        → US thousand-separated, no decimal
  "1,500,000.50"     → US with decimal dot
  "Rp 1.500.000,00"  → with currency prefix
  "1500000"          → plain integer string
  1500000            → numeric cell (int/float)
  1500000.5          → numeric cell with decimal
"""

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


def parse_amount(raw) -> Decimal:
    """
    Parse any reasonable amount format into Decimal.
    Raises ValueError if parsing fails.
    """
    if raw is None:
        raise ValueError("Amount is None")

    # Already numeric (Excel stored as number)
    if isinstance(raw, (int, float)):
        return Decimal(str(raw))

    text = str(raw).strip()

    # Strip currency symbols and spaces
    text = re.sub(r"[Rp\s$€£¥]", "", text)

    if not text:
        raise ValueError(f"Empty amount after stripping: {raw!r}")

    # Detect format:
    # Case 1: has both . and ,
    if "." in text and "," in text:
        last_dot  = text.rfind(".")
        last_comma = text.rfind(",")

        if last_dot > last_comma:
            # "1,500,000.50" → US format → remove commas, dot is decimal
            text = text.replace(",", "")
        else:
            # "1.500.000,50" → Indonesian format → remove dots, replace comma with dot
            text = text.replace(".", "").replace(",", ".")

    # Case 2: only commas
    elif "," in text and "." not in text:
        # Could be: "1,500,000" (US thousands) or "1,50" (decimal comma, unlikely w/o dot)
        # Heuristic: if there are multiple commas OR comma position is not at position -3/-7
        # treat as thousand separator; else treat as decimal
        parts = text.split(",")
        if len(parts) > 2:
            # Multiple commas → thousand separator (e.g. 1,500,000)
            text = text.replace(",", "")
        elif len(parts) == 2 and len(parts[-1]) == 2:
            # e.g. "1500000,50" → decimal comma (Indonesian)
            text = text.replace(",", ".")
        elif len(parts) == 2 and len(parts[-1]) == 3:
            # e.g. "1,500" → ambiguous, treat as thousand separator
            text = text.replace(",", "")
        else:
            # Fallback: treat comma as decimal
            text = text.replace(",", ".")

    # Case 3: only dots
    elif "." in text and "," not in text:
        parts = text.split(".")
        if len(parts) > 2:
            # Multiple dots → thousand separator (e.g. 1.500.000)
            text = text.replace(".", "")
        elif len(parts) == 2 and len(parts[-1]) == 3:
            # e.g. "1.500" → ambiguous, treat as thousand separator (Indonesian common)
            text = text.replace(".", "")
        else:
            # e.g. "1500.50" → decimal dot
            pass  # keep as-is

    # Strip any remaining non-numeric except dot and minus
    text = re.sub(r"[^\d.\-]", "", text)

    try:
        return Decimal(text)
    except InvalidOperation:
        raise ValueError(f"Cannot parse amount: {raw!r} → cleaned: {text!r}")


def normalize_for_compare(amount: Decimal) -> Decimal:
    """
    Round to 0 decimal places for comparison.
    Indonesian bank transactions are typically whole numbers (no cents).
    ponytail: assumes integer amounts; if cents matter, remove this rounding.
    """
    return amount.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
