"""
test_amount_utils.py — quick self-test for amount parser.
Run: python test_amount_utils.py
No frameworks needed.
"""

from decimal import Decimal
from amount_utils import parse_amount, normalize_for_compare

cases = [
    # (input, expected_decimal)
    ("1.500.000",       Decimal("1500000")),
    ("1.500.000,50",    Decimal("1500001")),   # rounded
    ("1,500,000",       Decimal("1500000")),
    ("1,500,000.50",    Decimal("1500001")),   # rounded
    ("Rp 1.500.000",    Decimal("1500000")),
    ("Rp 1.500.000,00", Decimal("1500000")),
    ("1500000",         Decimal("1500000")),
    (1500000,           Decimal("1500000")),
    (1500000.5,         Decimal("1500001")),   # rounded
    ("500",             Decimal("500")),
    ("1.500",           Decimal("1500")),      # ambiguous → treated as thousands
]

passed = 0
failed = 0
for raw, expected in cases:
    try:
        result = normalize_for_compare(parse_amount(raw))
        if result == expected:
            print(f"  ✅  {str(raw)!r:25} → {result}")
            passed += 1
        else:
            print(f"  ❌  {str(raw)!r:25} → {result}  (expected {expected})")
            failed += 1
    except Exception as e:
        print(f"  ❌  {str(raw)!r:25} → ERROR: {e}")
        failed += 1

print(f"\n{passed} passed, {failed} failed")
assert failed == 0, "Some amount parsing tests failed!"
print("All good ✅")
