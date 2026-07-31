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


# ── Colors ────────────────────────────────────────────────────────────────────
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

BANK_COLORS = {
    "BCA":     "2F75B6",
    "Mandiri": "375623",
    "BRI":     "833C00",
}


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
    cell = ws.cell(row, 1)
    cell.value = text
    cell.font = Font(bold=True, color=COLOR_WHITE, size=13)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 28


def _write_info_row(ws, parts: list[tuple[str, str]], ncols: int, row: int, bg: str = "F2F2F2"):
    """
    Write a merged info row containing key: value pairs separated by spaces.
    parts = [(label, value), ...]
    """
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row, 1)
    segments = "      ".join(f"{lbl}: {val}" for lbl, val in parts)
    cell.value = segments
    cell.font = Font(size=10, color="333333")
    cell.fill = PatternFill("solid", fgColor="F2F2F2")
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    cell.border = Border(bottom=Side(style="thin", color="CCCCCC"))
    ws.row_dimensions[row].height = 20



def _write_bank_sheet(ws, rows: list[dict], bank_name: str, odo_date=None):
    color = BANK_COLORS.get(bank_name, COLOR_HEADER)
    ws.title = bank_name
    ncols = 9

    done_count  = sum(1 for r in rows if r["status"] == STATUS_DONE)
    bank_count  = sum(1 for r in rows if r["status"] == STATUS_BANK_ONLY)
    odo_count   = sum(1 for r in rows if r["status"] == STATUS_ODO_ONLY)
    date_str    = odo_date.strftime("%d %B %Y") if odo_date else "-"

    # ── Title block (rows 1-2) ──────────────────────────────────────────
    _merge_title(ws, f"REKONSILIASI {bank_name.upper()}", ncols, row=1, bg=color)
    _write_info_row(ws, [
        ("Tanggal",        date_str),
        ("Total",          str(len(rows))),
        ("Done",           str(done_count)),
        ("Cuma di Bank",   str(bank_count)),
        ("Cuma di ODO",    str(odo_count)),
    ], ncols, row=2, bg=color)

    # ── Column headers (row 3) ───────────────────────────────────────
    COL_HEADERS = ["No", "Tanggal", "Nomor ODO", "Nomor Bank", "Filename", "Jumlah (Raw)", "Jumlah", "Sumber", "Status"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(row=3, column=col), h, bg=color)
    ws.row_dimensions[3].height = 28

    # ── Data rows (start row 4) ───────────────────────────────────────
    rows_sorted = sorted(rows, key=lambda r: r.get("date", "") or "")
    for idx, r in enumerate(rows_sorted, 1):
        rn = idx + 3  # offset by 3 header rows
        st = r["status"]
        _cell(ws.cell(rn, 1), idx,                              st, "center")
        _cell(ws.cell(rn, 2), r.get("date", ""),                st)
        _cell(ws.cell(rn, 3), r.get("number_odo",  ""),         st)
        _cell(ws.cell(rn, 4), r.get("number_bank", ""),         st)
        _cell(ws.cell(rn, 5), r.get("filename_bank", ""),       st)
        _cell(ws.cell(rn, 6), str(r.get("amount_raw", "")),     st, "right")
        _cell(ws.cell(rn, 7), _rp(r["amount"]),                 st, "right")
        _cell(ws.cell(rn, 8), r.get("source", ""),              st, "center")
        sc = ws.cell(rn, 9)
        _cell(sc, st, st, "center")
        sc.font = Font(bold=True, size=10)
        ws.row_dimensions[rn].height = 18

    for col, w in enumerate([6, 14, 22, 22, 30, 20, 20, 14, 22], 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = "A4"  # freeze title + col-header rows

    # ── Summary at bottom ─────────────────────────────────────────────
    gap = len(rows) + 5  # 3 header rows + 1 blank gap
    ws.cell(gap,     1, "RINGKASAN").font = Font(bold=True, size=11)
    ws.cell(gap + 1, 1, "Done (Matched)").font = Font(size=10)
    ws.cell(gap + 1, 2, done_count).font = Font(bold=True, size=10, color="375623")
    ws.cell(gap + 2, 1, "Cuma ada di Bank").font = Font(size=10)
    ws.cell(gap + 2, 2, bank_count).font = Font(bold=True, size=10, color="833C00")
    ws.cell(gap + 3, 1, "Cuma ada di ODO").font = Font(size=10)
    ws.cell(gap + 3, 2, odo_count).font = Font(bold=True, size=10, color="2F75B6")


def _write_discrepancy_sheet(ws, all_results: dict[str, list[dict]], odo_date=None):
    ws.title = "Selisih (Semua)"
    ncols = 9
    date_str   = odo_date.strftime("%d %B %Y") if odo_date else "-"
    total_disc = sum(
        1 for rows in all_results.values()
        for r in rows if r["status"] != STATUS_DONE
    )

    # ── Title block ────────────────────────────────────────────────────
    _merge_title(ws, "TRANSAKSI SELISIH — SEMUA BANK", ncols, row=1)
    _write_info_row(ws, [
        ("Tanggal",       date_str),
        ("Total Selisih", str(total_disc)),
    ], ncols, row=2)

    # ── Column headers (row 3) ─────────────────────────────────────
    COL_HEADERS = ["No", "Bank", "Tanggal", "Nomor ODO", "Nomor Bank", "Filename", "Jumlah", "Sumber", "Status"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(3, col), h)
    ws.row_dimensions[3].height = 28

    idx = 0
    for bank_name, rows in all_results.items():
        rows_sorted = sorted(
            (r for r in rows if r["status"] != STATUS_DONE),
            key=lambda r: r.get("date", "") or ""
        )
        for r in rows_sorted:
            idx += 1
            rn = idx + 3
            st = r["status"]
            _cell(ws.cell(rn, 1), idx,                            st, "center")
            _cell(ws.cell(rn, 2), bank_name,                      st, "center")
            _cell(ws.cell(rn, 3), r.get("date", ""),              st)
            _cell(ws.cell(rn, 4), r.get("number_odo",  ""),       st)
            _cell(ws.cell(rn, 5), r.get("number_bank", ""),       st)
            _cell(ws.cell(rn, 6), r.get("filename_bank", ""),     st)
            _cell(ws.cell(rn, 7), _rp(r["amount"]),               st, "right")
            _cell(ws.cell(rn, 8), r.get("source", ""),            st, "center")
            sc = ws.cell(rn, 9)
            _cell(sc, st, st, "center")
            sc.font = Font(bold=True, size=10)
            ws.row_dimensions[rn].height = 18

    for col, w in enumerate([6, 12, 14, 22, 22, 30, 20, 14, 22], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"

    if idx == 0:
        ws.cell(4, 1, "🎉 Semua transaksi matched!").font = Font(bold=True, color="375623")


def _write_legend(ws):
    ws.title = "Legenda"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 50
    entries = [
        ("Status", "Arti", True),
        (STATUS_DONE,      "Transaksi ada di Bank DAN di Odoo",          False),
        (STATUS_BANK_ONLY, "Transaksi HANYA ada di Bank, tidak di Odoo", False),
        (STATUS_ODO_ONLY,  "Transaksi HANYA ada di Odoo, tidak di Bank", False),
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


def _write_daily_summary_sheet(
    ws,
    all_results: dict,   # {bank_name: [reconciled result dicts]}
    odo_date=None,
):
    """
    Sheet 'Ringkasan Harian': per-bank per-date sum comparison.

    Uses reconciled results (not raw transactions) so that Done pairs
    always cancel out — selisih only reflects truly unmatched transactions.

    Logic per row:
      Done      → bank_sum[bank_date] += amt  AND  odo_sum[bank_date] += amt
                  (they cancel out → selisih = 0 for matched pairs)
      Bank Only → bank_sum[bank_date] += amt
      ODO Only  → odo_sum[odo_date]  += amt
    """
    from collections import defaultdict
    from decimal import Decimal

    ws.title = "Ringkasan Harian"
    ncols = 7
    date_str = odo_date.strftime("%d %B %Y") if odo_date else "-"

    _merge_title(ws, "RINGKASAN HARIAN — BANK vs ODO", ncols, row=1)
    _write_info_row(ws, [("Tanggal", date_str)], ncols, row=2)

    COL_HEADERS = ["No", "Bank", "Tanggal", "Total Bank", "Total ODO", "Selisih", "Status"]
    for col, h in enumerate(COL_HEADERS, 1):
        _hdr(ws.cell(3, col), h)
    ws.row_dimensions[3].height = 28

    # Build sums from reconciled results
    bank_sums: dict = defaultdict(lambda: defaultdict(Decimal))
    odo_sums:  dict = defaultdict(lambda: defaultdict(Decimal))

    all_banks = sorted(all_results.keys())

    for bank in all_banks:
        for r in all_results.get(bank, []):
            d   = r.get("date", "") or ""
            amt = r.get("amount", Decimal(0))
            st  = r.get("status", "")
            if st == STATUS_DONE:
                # Both sides attributed to same date → cancel out in selisih
                bank_sums[bank][d] += amt
                odo_sums[bank][d]  += amt
            elif st == STATUS_BANK_ONLY:
                bank_sums[bank][d] += amt
            elif st == STATUS_ODO_ONLY:
                odo_sums[bank][d]  += amt

    # Build sorted rows
    rows = []
    for bank in all_banks:
        all_dates = sorted(
            set(list(bank_sums[bank].keys()) + list(odo_sums[bank].keys()))
        )
        for d in all_dates:
            b_sum = bank_sums[bank].get(d, Decimal(0))
            o_sum = odo_sums[bank].get(d, Decimal(0))
            rows.append((bank, d, b_sum, o_sum))

    COLOR_MATCH   = "E2EFDA"   # light green  — sesuai
    COLOR_DIFF    = "FCE4D6"   # light orange — selisih
    COLOR_MISSING = "FFF2CC"   # light yellow — one side missing

    for idx, (bank, d, b_sum, o_sum) in enumerate(rows, 1):
        rn  = idx + 3
        sel = b_sum - o_sum

        if b_sum == 0 or o_sum == 0:
            bg = COLOR_MISSING
            label = "Data Tidak Lengkap"
        elif sel == 0:
            bg = COLOR_MATCH
            label = "✅ Sesuai"
        else:
            bg = COLOR_DIFF
            label = "⚠️ Selisih"

        def _sc(cell, val, align="left", _bg=bg):
            cell.value = val
            cell.fill = PatternFill("solid", fgColor=_bg)
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.font = Font(size=10)
            cell.border = Border(
                bottom=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin", color="CCCCCC"),
            )

        NUM_FMT = "#,##0"  # thousands separator, no decimals

        def _num(cell, val, _bg=bg):
            """Write a numeric cell with thousands-separator format."""
            cell.value = float(val)
            cell.number_format = NUM_FMT
            cell.fill = PatternFill("solid", fgColor=_bg)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = Font(size=10)
            cell.border = Border(
                bottom=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin", color="CCCCCC"),
            )

        _sc(ws.cell(rn, 1), idx,   "center")
        _sc(ws.cell(rn, 2), bank,  "center")
        _sc(ws.cell(rn, 3), d)
        _num(ws.cell(rn, 4), b_sum)
        _num(ws.cell(rn, 5), o_sum)
        _num(ws.cell(rn, 6), sel)
        lc = ws.cell(rn, 7)
        _sc(lc, label, "center")
        lc.font = Font(size=10, bold=(sel != 0))
        ws.row_dimensions[rn].height = 18

    for col, w in enumerate([6, 14, 14, 20, 20, 20, 22], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"


def write_report(
    all_results: dict[str, list[dict]],  # {bank_name: [result dicts]}
    odo_date: date,
    output_dir: Path,
    bank_txns: dict | None = None,       # unused, kept for backward compat
    odo_bank_txns: dict | None = None,   # unused, kept for backward compat
) -> Path:
    """Write full reconciliation report. Returns path to file."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    date_str  = odo_date.strftime("%d%m%Y") if odo_date else "unknown"
    out_path  = output_dir / f"reconciliation_{date_str}_{timestamp}.xlsx"

    wb = openpyxl.Workbook()
    first = True

    for bank_name, rows in all_results.items():
        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()
        _write_bank_sheet(ws, rows, bank_name, odo_date=odo_date)

    # Daily summary sheet (before discrepancy)
    ws_daily = wb.create_sheet()
    _write_daily_summary_sheet(ws_daily, all_results, odo_date=odo_date)

    ws_disc = wb.create_sheet()
    _write_discrepancy_sheet(ws_disc, all_results, odo_date=odo_date)

    ws_legend = wb.create_sheet()
    _write_legend(ws_legend)

    wb.save(out_path)
    return out_path
