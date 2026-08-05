"""
excel_writer.py — writes reconciliation results to a formatted Excel report.

Output: output/reconciliation_YYYYMMDD_HHMMSS.xlsx
  - One sheet per bank: "BCA", "Mandiri", "BRI"
  - Sheet "Selisih (Semua)"  — all discrepancies combined
  - Sheet "Legenda"          — color legend
"""

from datetime import datetime, date
from decimal import Decimal
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reconciler import STATUS_DONE, STATUS_BANK_ONLY, STATUS_ODO_ONLY
from config import BANK_ACCOUNTS, ODOO_COMPANY_NAME

def _get_account_info(acc_key: str) -> tuple[str, str]:
    if acc_key == "other":
        return "Lainnya", "Belum Teridentifikasi"
    parts = acc_key.split("_", 1)
    if len(parts) == 2:
        bank = parts[0]
        alias = parts[1]
        acc_info = BANK_ACCOUNTS.get(bank, {}).get(alias, {})
        acc_name = acc_info.get("group") or alias
        return bank.upper(), acc_name
    return acc_key.upper(), acc_key


# ── Constants ─────────────────────────────────────────────────────────────
STATUS_DONE      = "Match"
STATUS_BANK_ONLY = "Only in Bank"
STATUS_ODO_ONLY  = "Only in Odoo"

COLOR_HEADER    = "1F4E79"
COLOR_DONE      = "E2EFDA"
COLOR_BANK_ONLY = "FCE4D6"
COLOR_ODO_ONLY  = "DDEBF7"
COLOR_WHITE     = "FFFFFF"

STATUS_COLOR = {
    STATUS_DONE:       COLOR_DONE,
    STATUS_BANK_ONLY:  COLOR_BANK_ONLY,
    STATUS_ODO_ONLY:   COLOR_ODO_ONLY,
}

THIN   = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLOR_HEADER    = "4F81BD"

def _rp(amount: Decimal) -> str:
    return f"Rp {int(amount):,}".replace(",", ".")


def _hdr(cell, text: str, bg: str = COLOR_HEADER):
    cell.value = text
    cell.font = Font(bold=True, color=COLOR_WHITE, size=11)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def _cell(cell, value, status: str, align="left"):
    cell.value = value
    cell.fill = PatternFill("solid", fgColor=STATUS_COLOR.get(status, COLOR_WHITE))
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = BORDER
    cell.font = Font(size=10)

def _merge_title(ws, text: str, ncols: int, row: int, bg: str = COLOR_HEADER):
    """Write a full-width merged title cell."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    
    for col in range(1, ncols + 1):
        cell = ws.cell(row, col)
        if col == 1:
            cell.value = text
            cell.font = Font(size=14, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.fill = PatternFill("solid", fgColor=bg)
        
    ws.row_dimensions[row].height = 28


def _write_info_row(ws, parts: list[tuple[str, str]], ncols: int, row: int, bg: str = "DDEBF7"):
    """
    Write a merged info row containing key: value pairs separated by pipes.
    parts = [(label, value), ...]
    """
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    
    segments = "   |   ".join(f"{lbl}: {val}" for lbl, val in parts)
    
    for col in range(1, ncols + 1):
        cell = ws.cell(row, col)
        if col == 1:
            cell.value = segments
            cell.font = Font(size=10, color="1F4E79", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.border = Border(bottom=Side(style="medium", color="4F81BD"))
        
    ws.row_dimensions[row].height = 22

def _fmt_date(date_val) -> str:
    if not date_val:
        return ""
    if isinstance(date_val, (date, datetime)):
        return date_val.strftime("%d/%m/%Y")
    s = str(date_val).strip()
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        try:
            return datetime.strptime(s[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            pass
    return s

def _fmt_amount_str(val) -> str:
    if not val:
        return "0,00"
    try:
        s = f"{float(val):,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(val)


def _write_bank_sheet(ws, rows: list[dict], bank_name: str, odo_date=None):
    ws.title = bank_name[:31]
    ncols = 9

    done_count  = sum(1 for r in rows if r["status"] == STATUS_DONE)
    bank_count  = sum(1 for r in rows if r["status"] == STATUS_BANK_ONLY)
    odo_count   = sum(1 for r in rows if r["status"] == STATUS_ODO_ONLY)
    date_str    = odo_date.strftime("%d %B %Y") if odo_date else "-"
    total_amount = sum(float(r.get("amount", 0)) for r in rows)

    # ── Title block (rows 1-2) ──────────────────────────────────────────
    _merge_title(ws, f"RECONCILIATION {bank_name.upper()}", ncols, row=1)
    _write_info_row(ws, [
        ("Date",           date_str),
        ("Total Amount",   _fmt_amount_str(total_amount)),
        ("Total",          str(len(rows))),
        ("Match",          str(done_count)),
    ], ncols, row=2)

    # ── Column headers (row 3) ───────────────────────────────────────
    COL_HEADERS = ["No", "Date", "Payment Date", "Odoo Number", "Bank Number", "Filename", "Amount", "Source", "Status"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(row=3, column=col), h)
    ws.row_dimensions[3].height = 28

    # ── Data rows (start row 4) ───────────────────────────────────────
    rows_sorted = sorted(rows, key=lambda r: r.get("date", "") or "")
    for idx, r in enumerate(rows_sorted, 1):
        rn = idx + 3  # offset by 3 header rows
        st = r["status"]
        _cell(ws.cell(rn, 1), idx,                              st, "center")
        _cell(ws.cell(rn, 2), _fmt_date(r.get("date", "")),         st, "center")
        _cell(ws.cell(rn, 3), _fmt_date(r.get("payment_date", "")), st, "center")
        _cell(ws.cell(rn, 4), r.get("number_odo",  ""),         st)
        _cell(ws.cell(rn, 5), r.get("number_bank", ""),         st)
        _cell(ws.cell(rn, 6), r.get("filename_bank", ""),       st)
        
        c_amt = ws.cell(rn, 7)
        _cell(c_amt, float(r["amount"]), st, "right")
        c_amt.number_format = '#,##0.00'
        
        _cell(ws.cell(rn, 8), r.get("source", ""),              st, "center")
        sc = ws.cell(rn, 9)
        _cell(sc, st, st, "center")
        sc.font = Font(bold=True, size=10)
        ws.row_dimensions[rn].height = 18

    for col, width in enumerate([6, 15, 15, 20, 20, 40, 20, 15, 15], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{len(rows) + 3}"
    ws.freeze_panes = "A4"  # freeze title + col-header rows


def _write_other_sheet(ws, rows: list[dict], odo_date=None):
    ws.title = "Other Payment"
    ncols = 6
    
    date_str    = odo_date.strftime("%d %B %Y") if odo_date else "-"
    total_amount = sum(float(r.get("amount", 0)) for r in rows)

    # ── Title block (rows 1-2) ──────────────────────────────────────────
    _merge_title(ws, "OTHER PAYMENT", ncols, row=1)
    _write_info_row(ws, [
        ("Date",           date_str),
        ("Total Amount",   _fmt_amount_str(total_amount)),
        ("Total",          str(len(rows))),
    ], ncols, row=2)

    # ── Column headers (row 3) ───────────────────────────────────────
    COL_HEADERS = ["No", "Date", "Journal", "Odoo Number", "Amount", "Source"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(row=3, column=col), h)
    ws.row_dimensions[3].height = 28

    # ── Data rows (start row 4) ───────────────────────────────────────
    rows_sorted = sorted(rows, key=lambda r: r.get("date", "") or "")
    for idx, r in enumerate(rows_sorted, 1):
        rn = idx + 3
        st = r["status"]
        _cell(ws.cell(rn, 1), idx,                              st, "center")
        _cell(ws.cell(rn, 2), _fmt_date(r.get("date", "")),     st, "center")
        _cell(ws.cell(rn, 3), r.get("description", ""),         st, "center")
        _cell(ws.cell(rn, 4), r.get("number_odo",  ""),         st)
        
        c_amt = ws.cell(rn, 5)
        _cell(c_amt, float(r["amount"]), st, "right")
        c_amt.number_format = '#,##0.00'
        
        _cell(ws.cell(rn, 6), r.get("source", ""),              st, "center")
        ws.row_dimensions[rn].height = 18

    for col, width in enumerate([6, 15, 20, 20, 20, 15], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{len(rows) + 3}"
    ws.freeze_panes = "A4"


def _write_discrepancy_sheet(ws, all_results: dict[str, list[dict]], odo_date=None):
    ws.title = "Differences"
    ncols = 10
    date_str   = odo_date.strftime("%d %B %Y") if odo_date else "-"
    
    rows_data = [
        r for rows in all_results.values()
        for r in rows if r["status"] != STATUS_DONE
    ]
    total_disc = len(rows_data)
    total_amount = sum(float(r.get("amount", 0)) for r in rows_data)

    # ── Title block ────────────────────────────────────────────────────
    _merge_title(ws, "DIFFERENCES — ALL BANKS", ncols, row=1)
    _write_info_row(ws, [
        ("Date",          date_str),
        ("Total Amount",  _fmt_amount_str(total_amount)),
        ("Total Differences", str(total_disc)),
    ], ncols, row=2)

    # ── Column headers (row 3) ─────────────────────────────────────
    COL_HEADERS = ["No", "Date", "Bank", "Journal", "Odoo Number", "Bank Number", "Filename", "Amount", "Source", "Status"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(3, col), h)
    ws.row_dimensions[3].height = 28

    idx = 0
    for acc_key, rows in all_results.items():
        if acc_key == "other":
            continue
            
        bank_label, acc_name = _get_account_info(acc_key)
        rows_sorted = sorted(
            (r for r in rows if r["status"] != STATUS_DONE),
            key=lambda r: r.get("date", "") or ""
        )
        for r in rows_sorted:
            idx += 1
            rn = idx + 3
            st = r["status"]
            _cell(ws.cell(rn, 1), idx,                            st, "center")
            _cell(ws.cell(rn, 2), _fmt_date(r.get("date", "")),   st, "center")
            _cell(ws.cell(rn, 3), bank_label,                     st, "center")
            _cell(ws.cell(rn, 4), acc_name,                       st, "center")
            _cell(ws.cell(rn, 5), r.get("number_odo",  ""),       st)
            _cell(ws.cell(rn, 6), r.get("number_bank", ""),       st)
            _cell(ws.cell(rn, 7), r.get("filename_bank", ""),     st)
            
            c_amt = ws.cell(rn, 8)
            _cell(c_amt, float(r["amount"]), st, "right")
            c_amt.number_format = '#,##0.00'
            
            _cell(ws.cell(rn, 9), r.get("source", ""),            st, "center")
            sc = ws.cell(rn, 10)
            _cell(sc, st, st, "center")
            sc.font = Font(bold=True, size=10)
            ws.row_dimensions[rn].height = 18

    for col, width in enumerate([6, 15, 10, 18, 20, 20, 40, 20, 15, 15], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{total_disc + 3}"
    ws.freeze_panes = "A4"


def _write_legend(ws):
    ws.title = "Legend"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 50
    entries = [
        ("Status", "Meaning", True),
        (STATUS_DONE,      "Transaction exists in both Bank and Odoo",          False),
        (STATUS_BANK_ONLY, "Transaction exists ONLY in Bank, not in Odoo", False),
        (STATUS_ODO_ONLY,  "Transaction exists ONLY in Odoo, not in Bank", False),
    ]
    for i, (lbl, desc, is_hdr) in enumerate(entries, 1):
        a = ws.cell(i, 1, lbl)
        b = ws.cell(i, 2, desc)
        if is_hdr:
            for c in [a, b]:
                c.font = Font(bold=True, color=COLOR_WHITE)
                c.fill = PatternFill("solid", fgColor=COLOR_HEADER)
                c.alignment = Alignment(horizontal="center")
        else:
            col = STATUS_COLOR.get(lbl, COLOR_WHITE)
            for c in [a, b]:
                c.fill = PatternFill("solid", fgColor=col)
                c.font = Font(size=10)
            a.font = Font(bold=True, size=10)
        for c in [a, b]:
            c.border = BORDER
            c.alignment = Alignment(vertical="center")
        ws.row_dimensions[i].height = 20


def _write_daily_summary_sheet(ws, all_results: dict, odo_date=None):
    from collections import defaultdict
    from decimal import Decimal

    ws.title = "Daily Summary"
    ncols = 10
    date_str = odo_date.strftime("%d %B %Y") if odo_date else "-"

    _merge_title(ws, "DAILY SUMMARY — BANK vs ODO", ncols, row=1)
    _write_info_row(ws, [("Date", date_str)], ncols, row=2)

    COL_HEADERS = ["No", "Date", "Payment Date", "Bank", "Journal", "Total Bank", "Total Odoo", "Difference", "Status", "Journal Status"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(3, col), h)
    ws.row_dimensions[3].height = 28

    bank_sums: dict = defaultdict(lambda: defaultdict(Decimal))
    odo_sums:  dict = defaultdict(lambda: defaultdict(Decimal))
    payment_dates = {}

    all_banks = sorted([k for k in all_results.keys() if k != "other"])

    for bank in all_banks:
        for r in all_results.get(bank, []):
            d   = r.get("date", "") or ""
            pd  = r.get("payment_date", "")
            if pd:
                payment_dates[(bank, d)] = pd
            amt = r.get("amount", Decimal(0))
            st  = r.get("status", "")
            if st == STATUS_DONE:
                bank_sums[bank][d] += amt
                odo_sums[bank][d]  += amt
            elif st == STATUS_BANK_ONLY:
                bank_sums[bank][d] += amt
            elif st == STATUS_ODO_ONLY:
                odo_sums[bank][d]  += amt

    rows = []
    for acc_key in all_banks:
        bank_label, acc_name = _get_account_info(acc_key)
        all_dates = sorted(
            set(list(bank_sums[acc_key].keys()) + list(odo_sums[acc_key].keys()))
        )
        for d in all_dates:
            b_sum = bank_sums[acc_key].get(d, Decimal(0))
            o_sum = odo_sums[acc_key].get(d, Decimal(0))
            pd = payment_dates.get((acc_key, d), "")
            rows.append((bank_label, acc_name, d, pd, b_sum, o_sum))

    COLOR_MATCH   = "E2EFDA"
    COLOR_DIFF    = "FCE4D6"
    COLOR_MISSING = "FFF2CC"

    for idx, (bank, acc_name, d, pd, b_sum, o_sum) in enumerate(rows, 1):
        rn  = idx + 3
        sel = b_sum - o_sum

        if b_sum == 0 or o_sum == 0:
            bg = COLOR_MISSING
            label = "Incomplete Data"
        elif sel == 0:
            bg = COLOR_MATCH
            label = "✅ Match"
        else:
            bg = COLOR_DIFF
            label = "⚠️ Difference"

        def _sc(cell, val, align="left", _bg=bg):
            cell.value = val
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.fill = PatternFill("solid", fgColor=_bg)
            cell.font = Font(size=10)
            cell.border = Border(
                bottom=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin", color="CCCCCC"),
            )

        def _num(cell, val, _bg=bg):
            cell.value = float(val)
            cell.number_format = "#,##0.00"
            cell.fill = PatternFill("solid", fgColor=_bg)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = Font(size=10)
            cell.border = Border(
                bottom=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin", color="CCCCCC"),
            )

        _sc(ws.cell(rn, 1), idx, "center")
        _sc(ws.cell(rn, 2), _fmt_date(d), "center")
        _sc(ws.cell(rn, 3), _fmt_date(pd), "center")
        _sc(ws.cell(rn, 4), bank, "center")
        _sc(ws.cell(rn, 5), acc_name, "center")
        _num(ws.cell(rn, 6), b_sum)
        _num(ws.cell(rn, 7), o_sum)
        _num(ws.cell(rn, 8), sel, _bg=COLOR_DIFF if sel != 0 else bg)
        lc = ws.cell(rn, 9)
        _sc(lc, label, "center")
        lc.font = Font(size=10, bold=True)
        
        jc = ws.cell(rn, 10)
        _sc(jc, "⏳ Not Yet", "center")
        jc.font = Font(size=10, bold=True)
        ws.row_dimensions[rn].height = 18

    for col, width in enumerate([6, 15, 15, 10, 18, 20, 20, 20, 15, 15], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{len(rows) + 3}"
    ws.freeze_panes = "A4"


def write_report(
    all_results: dict[str, list[dict]],
    odo_date: date,
    output_dir: Path,
    bank_txns: dict | None = None,
    odo_bank_txns: dict | None = None,
    mutations: list[dict] = None,
    unknown_mutations: list[dict] = None,
    excluded_txns: list[dict] = None,
) -> Path:
    date_str = odo_date.strftime("%d%m%Y") if odo_date else "unknown"
    company_str = ODOO_COMPANY_NAME.replace(" ", "_") if ODOO_COMPANY_NAME else "Company"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"Reconciliation_{company_str}_{date_str}_{timestamp}.xlsx"

    wb = openpyxl.Workbook()
    
    # 1. Create Daily Summary as the first sheet
    ws_daily = wb.active
    ws_daily.title = "Daily Summary"
    _write_daily_summary_sheet(ws_daily, all_results, odo_date=odo_date)
    
    # 2. Create Differences as the second sheet
    ws_disc = wb.create_sheet(title="Differences")
    _write_discrepancy_sheet(ws_disc, all_results, odo_date=odo_date)
    
    # 3. Mutation Summary
    if mutations is not None:
        ws_mut = wb.create_sheet(title="Mutation Summary")
        _write_mutation_summary(ws_mut, all_results, mutations, odo_date=odo_date)
        
    # 4. Admin Fee
    ws_admin_fee = wb.create_sheet(title="Admin Fee")
    _write_admin_fee_sheet(ws_admin_fee, all_results, mutations, odo_date=odo_date)

    # 5. Excluded Payment
    if excluded_txns is not None and excluded_txns:
        ws_excluded = wb.create_sheet(title="Excluded Payment")
        _write_excluded_sheet(ws_excluded, excluded_txns, odo_date=odo_date)
    
    other_rows = all_results.pop("other", [])
    
    # 6. Bank Sheets
    for acc_key, rows in all_results.items():
        bank_label, acc_name = _get_account_info(acc_key)
        sheet_name = f"{bank_label} ({acc_name})"[:31]
        ws = wb.create_sheet(title=sheet_name)
        _write_bank_sheet(ws, rows, f"{bank_label} ({acc_name})", odo_date=odo_date)

    # 7. Other Payment
    if other_rows:
        ws_other = wb.create_sheet(title="Other Payment")
        _write_other_sheet(ws_other, other_rows, odo_date=odo_date)
        
    # 8. Other Mutation
    if unknown_mutations is not None:
        ws_unmapped = wb.create_sheet(title="Other Mutation")
        _write_unmapped_mutations(ws_unmapped, all_results, unknown_mutations, odo_date=odo_date)
        
    ws_legend = wb.create_sheet(title="Legend")
    _write_legend(ws_legend)

    wb.save(out_path)
    return out_path


def _build_date_lookup(all_results):
    date_lookup = {}
    if all_results:
        for acc_key, rows in all_results.items():
            if acc_key == "other": continue
            bank, acc_name = _get_account_info(acc_key)
            for r in rows:
                if r["status"] == STATUS_DONE:
                    p_date = _fmt_date(r.get("payment_date", ""))
                    o_date = _fmt_date(r.get("date", ""))
                    if p_date and o_date:
                        key = (bank, acc_name, p_date)
                        if key not in date_lookup:
                            date_lookup[key] = set()
                        date_lookup[key].add(o_date)
    return date_lookup

def _write_mutation_summary(ws, all_results, mutations, odo_date=None):
    ws.title = "Mutation Summary"
    ncols = 8
    date_str = odo_date.strftime("%d %B %Y") if odo_date else "-"
    total_amount = sum(float(m.get("amount", 0)) for m in mutations) if mutations else 0.0
    
    _merge_title(ws, "MUTATION SUMMARY", ncols, row=1)
    _write_info_row(ws, [
        ("Date", date_str),
        ("Total Amount", _fmt_amount_str(total_amount))
    ], ncols, row=2)
    
    _hdr(ws.cell(3, 1), "No")
    _hdr(ws.cell(3, 2), "Date")
    _hdr(ws.cell(3, 3), "Payment Date")
    _hdr(ws.cell(3, 4), "Bank")
    _hdr(ws.cell(3, 5), "Journal")
    _hdr(ws.cell(3, 6), "Description")
    _hdr(ws.cell(3, 7), "Transaction Category")
    _hdr(ws.cell(3, 8), "Total Amount")
    ws.row_dimensions[3].height = 28
    
    date_lookup = _build_date_lookup(all_results)
    
    from collections import defaultdict
    summary = defaultdict(float)
    for m in mutations:
        p_str = _fmt_date(m["date"])
        bank = m["bank"].upper()
        _, acc_name = _get_account_info(f"{m['bank']}_{m['alias']}")
        
        o_dates = date_lookup.get((bank, acc_name, p_str), set())
        o_str = ", ".join(sorted(o_dates)) if o_dates else "-"
        
        desc = m.get("desc", "")
        key = (o_str, p_str, bank, acc_name, desc, m.get("type", "Unknown"))
        summary[key] += m["amount"]
        
    row = 4
    idx = 1
    for key in sorted(summary.keys()):
        o_str, p_str, bank, acc_name, desc, txn_type = key
        amount = summary[key]
        
        _cell(ws.cell(row, 1), idx, STATUS_DONE, "center")
        _cell(ws.cell(row, 2), o_str, STATUS_DONE, "center")
        _cell(ws.cell(row, 3), p_str, STATUS_DONE, "center")
        _cell(ws.cell(row, 4), bank, STATUS_DONE, "center")
        _cell(ws.cell(row, 5), acc_name, STATUS_DONE, "center")
        _cell(ws.cell(row, 6), desc, STATUS_DONE, "left")
        _cell(ws.cell(row, 7), txn_type, STATUS_DONE, "center")
        
        c_amt = ws.cell(row, 8)
        _cell(c_amt, amount, STATUS_DONE, "right")
        c_amt.number_format = '#,##0.00'
        
        row += 1
        idx += 1

    for col, width in enumerate([6, 15, 15, 12, 20, 40, 20, 20], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{row - 1}"
    ws.freeze_panes = "A4"


def _write_admin_fee_sheet(ws, all_results, mutations, odo_date=None):
    ws.title = "Admin Fee"
    ncols = 8
    date_str = odo_date.strftime("%d %B %Y") if odo_date else "-"
    
    total_admin = 0.0
    if mutations:
        total_admin += sum(float(m.get("admin_fee") or 0) for m in mutations)
    for acc_key, rows in all_results.items():
        if acc_key.startswith("bca") or acc_key.startswith("bri") or acc_key == "other":
            continue
        total_admin += sum(float(r.get("admin_fee") or 0) for r in rows)
    
    _merge_title(ws, "ADMIN FEE", ncols, row=1)
    _write_info_row(ws, [
        ("Date", date_str),
        ("Total Amount", _fmt_amount_str(total_admin))
    ], ncols, row=2)
    
    _hdr(ws.cell(3, 1), "No")
    _hdr(ws.cell(3, 2), "Date")
    _hdr(ws.cell(3, 3), "Payment Date")
    _hdr(ws.cell(3, 4), "Bank")
    _hdr(ws.cell(3, 5), "Journal")
    _hdr(ws.cell(3, 6), "Description")
    _hdr(ws.cell(3, 7), "Transaction Category")
    _hdr(ws.cell(3, 8), "Total Admin Fee")
    ws.row_dimensions[3].height = 28
    
    date_lookup = _build_date_lookup(all_results)
    
    from collections import defaultdict
    summary = defaultdict(float)
    
    if mutations:
        for m in mutations:
            adm = float(m.get("admin_fee") or 0)
            if adm > 0:
                d_str = _fmt_date(m["date"])
                bank = m["bank"].upper()
                parts = _get_account_info(f"{m['bank']}_{m['alias']}")
                acc_name = parts[1]
                
                o_dates = date_lookup.get((bank, acc_name, d_str), set())
                o_str = ", ".join(sorted(o_dates)) if o_dates else "-"
                
                desc = m.get("desc", "")
                key = (o_str, d_str, bank, acc_name, desc, m.get("type", "Unknown"))
                summary[key] += adm
                
    for acc_key, rows in all_results.items():
        if acc_key.startswith("bca") or acc_key.startswith("bri") or acc_key == "other":
            continue
            
        bank_label, acc_name = _get_account_info(acc_key)
        
        for r in rows:
            adm = float(r.get("admin_fee") or 0)
            if adm > 0:
                p_str = _fmt_date(r.get("payment_date", ""))
                o_str = _fmt_date(r.get("date", ""))
                if not p_str:
                    p_str = "Unknown"
                if not o_str:
                    o_str = "-"
                
                cat = r.get("category") or "Unknown"
                desc = r.get("description") or "-"
                key = (o_str, p_str, bank_label, acc_name, desc, cat)
                summary[key] += adm

    row = 4
    idx = 1
    for key in sorted(summary.keys()):
        o_str, p_str, bank, acc_name, desc, txn_type = key
        amount = summary[key]
        
        _cell(ws.cell(row, 1), idx, STATUS_DONE, "center")
        _cell(ws.cell(row, 2), o_str, STATUS_DONE, "center")
        _cell(ws.cell(row, 3), p_str, STATUS_DONE, "center")
        _cell(ws.cell(row, 4), bank, STATUS_DONE, "center")
        _cell(ws.cell(row, 5), acc_name, STATUS_DONE, "center")
        _cell(ws.cell(row, 6), desc, STATUS_DONE, "left")
        _cell(ws.cell(row, 7), txn_type, STATUS_DONE, "center")
        
        c_amt = ws.cell(row, 8)
        _cell(c_amt, amount, STATUS_DONE, "right")
        c_amt.number_format = '#,##0.00'
        row += 1
        idx += 1

    for col, width in enumerate([6, 15, 15, 12, 20, 30, 20, 20], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(8)}{row - 1}"
    ws.freeze_panes = "A4"


def _write_unmapped_mutations(ws, all_results, unknowns, odo_date=None):
    ws.title = "Other Mutation"
    ncols = 8
    date_str = odo_date.strftime("%d %B %Y") if odo_date else "-"
    total_amount = sum(float(m.get("amount", 0)) for m in unknowns) if unknowns else 0.0
    
    _merge_title(ws, "OTHER MUTATION", ncols, row=1)
    _write_info_row(ws, [
        ("Date", date_str),
        ("Total Amount", _fmt_amount_str(total_amount))
    ], ncols, row=2)
    
    _hdr(ws.cell(3, 1), "No")
    _hdr(ws.cell(3, 2), "Date")
    _hdr(ws.cell(3, 3), "Payment Date")
    _hdr(ws.cell(3, 4), "Bank")
    _hdr(ws.cell(3, 5), "Journal")
    _hdr(ws.cell(3, 6), "Description")
    _hdr(ws.cell(3, 7), "Transaction Type")
    _hdr(ws.cell(3, 8), "Amount")
    ws.row_dimensions[3].height = 28
    
    date_lookup = _build_date_lookup(all_results)
    
    row = 4
    idx = 1
    for m in unknowns:
        p_str = _fmt_date(m["date"])
        bank_raw = m["bank"].upper()
        _, acc_name = _get_account_info(f"{m['bank']}_{m['alias']}")
        
        o_dates = date_lookup.get((bank_raw, acc_name, p_str), set())
        o_str = ", ".join(sorted(o_dates)) if o_dates else "-"
        
        _cell(ws.cell(row, 1), idx, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 2), o_str, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 3), p_str, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 4), bank_raw, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 5), acc_name, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 6), m.get("desc", ""), STATUS_BANK_ONLY, "left")
        _cell(ws.cell(row, 7), m.get("cr_db_type", ""), STATUS_BANK_ONLY, "center")
        
        c_amt = ws.cell(row, 8)
        _cell(c_amt, float(m["amount"]), STATUS_BANK_ONLY, "right")
        c_amt.number_format = '#,##0.00'
        
        row += 1
        idx += 1

    for col, width in enumerate([6, 15, 15, 12, 18, 40, 14, 20], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{row - 1}"
    ws.freeze_panes = "A4"

def _write_excluded_sheet(ws, excluded_txns: list[dict], odo_date=None):
    ws.title = "Excluded Payment"
    ncols = 8
    date_str = odo_date.strftime("%d %B %Y") if odo_date else "-"
    total_amount = sum(float(t.get("amount", 0)) for t in excluded_txns)
    
    _merge_title(ws, "EXCLUDED TRANSACTIONS", ncols, row=1)
    _write_info_row(ws, [
        ("Date", date_str),
        ("Total Excluded", str(len(excluded_txns))),
        ("Total Amount", _fmt_amount_str(total_amount))
    ], ncols, row=2)
    
    _hdr(ws.cell(3, 1), "No")
    _hdr(ws.cell(3, 2), "Date")
    _hdr(ws.cell(3, 3), "Bank")
    _hdr(ws.cell(3, 4), "Trace Number")
    _hdr(ws.cell(3, 5), "Filename")
    _hdr(ws.cell(3, 6), "Description")
    _hdr(ws.cell(3, 7), "Amount")
    _hdr(ws.cell(3, 8), "Exclusion Reason")
    ws.row_dimensions[3].height = 28
    
    row = 4
    for idx, t in enumerate(excluded_txns, 1):
        _cell(ws.cell(row, 1), idx, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 2), _fmt_date(t.get("date", "")), STATUS_BANK_ONLY, "center")
        bank_name = t.get("source", "").replace("Bank (", "").replace(")", "")
        _cell(ws.cell(row, 3), bank_name, STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 4), t.get("number", ""), STATUS_BANK_ONLY, "center")
        _cell(ws.cell(row, 5), t.get("filename", ""), STATUS_BANK_ONLY, "left")
        _cell(ws.cell(row, 6), t.get("description", ""), STATUS_BANK_ONLY, "left")
        
        c_amt = ws.cell(row, 7)
        _cell(c_amt, float(t.get("amount", 0)), STATUS_BANK_ONLY, "right")
        c_amt.number_format = '#,##0.00'
        
        _cell(ws.cell(row, 8), t.get("exclusion_reason", "Voided"), STATUS_BANK_ONLY, "center")
        
        ws.row_dimensions[row].height = 18
        row += 1

    for col, width in enumerate([6, 15, 15, 20, 30, 30, 20, 25], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{row - 1}"
    ws.freeze_panes = "A4"
