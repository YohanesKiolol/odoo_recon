"""Confirm Journal Creation, Odoo Posting, and Export modal dialog."""
import glob
import os
import sys
import json
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from pathlib import Path
from collections import defaultdict
from excel_writer import safe_load_workbook
from journal_generator import generate_journal_import
from odoo_journal_creator import (
    get_draft_journal_details,
    post_draft_settlement_journals
)
from journal_checker import check_journals
import config
from ui.theme import (
    BG, PANEL, SIDEBAR_BG, PREVIEW_BG, BORDER, BORDER_DARK,
    ACCENT, ACCENT_DARK, SUCCESS, SUCCESS_DARK, ERROR, WARN,
    TEXT, MUTED, WHITE, FONT_FAMILY, IS_WINDOWS
)
from ui.widgets.window_utils import _center_modal_on_parent, _open_path

def open_journal_modal(
    parent,
    output_dir: Path,
    base_dir: Path,
    venv_python: str,
    log_write_fn=None,
    set_status_fn=None,
    on_run_fn=None,
    on_done_fn=None,
    on_manual_match_fn=None,
    update_dashboard_fn=None
):
    """Open Confirm Journal Creation modal."""
    output_files = glob.glob(str(output_dir / "[Rr]econciliation_*.xlsx"))
    if not output_files:
        if set_status_fn:
            set_status_fn("No reconciliation file found", ERROR)
        return None
    latest_file = max(output_files, key=os.path.getctime)

    top = ctk.CTkToplevel(parent)
    top.withdraw()
    top.title("Confirm Journal Creation")
    top.minsize(860, 480)
    top.resizable(True, True)
    top.configure(fg_color=BG)
    top.transient(parent)
    _center_modal_on_parent(top, parent)

    items = []
    journal_state = []
    mutation_totals = defaultdict(float)
    mutation_raw = []
    admin_totals = defaultdict(float)
    admin_raw = []

    def _load_data():
        nonlocal items, journal_state, mutation_totals, mutation_raw, admin_totals, admin_raw
        output_fls = glob.glob(str(output_dir / "[Rr]econciliation_*.xlsx"))
        if not output_fls: return
        cur_file = max(output_fls, key=os.path.getctime)
        try:
            wb = safe_load_workbook(cur_file, data_only=True)
            if "Daily Summary" not in wb.sheetnames:
                if set_status_fn:
                    set_status_fn("'Daily Summary' sheet not found", ERROR)
                wb.close()
                return

            def get_col_map(sheet, row_idx=3):
                return {
                    str(sheet.cell(row=row_idx, column=c).value).strip().lower(): c
                    for c in range(1, sheet.max_column + 1)
                    if sheet.cell(row=row_idx, column=c).value
                }

            if "Mutation Summary" in wb.sheetnames:
                ws_mut = wb["Mutation Summary"]
                col_map = get_col_map(ws_mut)
                c_date = col_map.get("payment date", 3)
                c_bank = col_map.get("bank", 4)
                c_group = col_map.get("journal", 5)
                c_cat = col_map.get("transaction category", 7)
                c_amt = col_map.get("total amount", 8)

                mutation_totals.clear()
                mutation_raw.clear()
                for row in range(4, ws_mut.max_row + 1):
                    tanggal = ws_mut.cell(row=row, column=c_date).value
                    bank = ws_mut.cell(row=row, column=c_bank).value
                    group = ws_mut.cell(row=row, column=c_group).value
                    cat = ws_mut.cell(row=row, column=c_cat).value
                    amount = ws_mut.cell(row=row, column=c_amt).value
                    if tanggal and group and amount:
                        mutation_raw.append({"payment_date": tanggal, "bank": bank, "group": group, "category": cat, "amount": float(amount)})
                        try:
                            mutation_totals[(str(tanggal).strip(), str(group).strip())] += float(amount)
                        except ValueError:
                            pass
            else:
                mutation_totals.clear()
                mutation_raw.clear()

            if "Admin Fee" in wb.sheetnames:
                ws_adm = wb["Admin Fee"]
                col_map = get_col_map(ws_adm)
                c_date = col_map.get("payment date", 3)
                c_bank = col_map.get("bank", 4)
                c_group = col_map.get("journal", 5)
                c_cat = col_map.get("transaction category", 7)
                c_amt = col_map.get("total amount", 8)

                admin_totals.clear()
                admin_raw.clear()
                for row in range(4, ws_adm.max_row + 1):
                    tanggal = ws_adm.cell(row=row, column=c_date).value
                    bank = ws_adm.cell(row=row, column=c_bank).value
                    group = ws_adm.cell(row=row, column=c_group).value
                    cat = ws_adm.cell(row=row, column=c_cat).value
                    amount = ws_adm.cell(row=row, column=c_amt).value
                    if tanggal and group and amount:
                        admin_raw.append({"payment_date": tanggal, "bank": bank, "group": group, "category": cat, "amount": float(amount)})
                        try:
                            admin_totals[(str(tanggal).strip(), str(group).strip())] += float(amount)
                        except ValueError:
                            pass
            else:
                admin_totals.clear()
                admin_raw.clear()

            ws = wb["Daily Summary"]
            col_map = get_col_map(ws)
            c_date = col_map.get("date", 2)
            c_pdate = col_map.get("payment date", 3)
            c_bank = col_map.get("bank", 4)
            c_group = col_map.get("journal", 5)
            c_tbank = col_map.get("total bank", 6)
            c_todoo = col_map.get("total odoo", 7)
            c_diff = col_map.get("difference", 8)
            c_recon = col_map.get("reconciled", 9)
            c_status = col_map.get("status", 10)
            c_jstatus = col_map.get("journal information", 11)
            c_edc_num = col_map.get("edc number", 12)
            c_ar_num = col_map.get("ar number", 13)

            items.clear()
            for row in range(4, ws.max_row + 1):
                bank = ws.cell(row=row, column=c_bank).value
                group = ws.cell(row=row, column=c_group).value
                tanggal = ws.cell(row=row, column=c_date).value
                payment_date = ws.cell(row=row, column=c_pdate).value
                total_bank = ws.cell(row=row, column=c_tbank).value
                total_odoo = ws.cell(row=row, column=c_todoo).value
                selisih = ws.cell(row=row, column=c_diff).value
                reconciled = ws.cell(row=row, column=c_recon).value
                status = ws.cell(row=row, column=c_status).value
                journal_status = ws.cell(row=row, column=c_jstatus).value
                edc_num = ws.cell(row=row, column=c_edc_num).value if c_edc_num <= ws.max_column else ""
                ar_num = ws.cell(row=row, column=c_ar_num).value if c_ar_num <= ws.max_column else ""

                if not bank or not status: continue
                if "incomplete" in str(status).lower(): continue

                try:
                    diff = abs(float(selisih)) if selisih is not None else 0
                except:
                    diff = 0

                status_str = str(status).strip()
                status_valid = ("Match" in status_str) or (diff <= config.JOURNAL_TOLERANCE)

                mutation_found = False
                mutation_matched = False
                mut_total = 0.0
                try:
                    d_mut_str = str(payment_date).strip()
                    mut_amount = mutation_totals.get((d_mut_str, str(group).strip()), 0.0)
                    adm_amount = admin_totals.get((d_mut_str, str(group).strip()), 0.0)
                    mut_total = mut_amount + adm_amount

                    if mut_amount > 0:
                        mutation_found = True
                        mut_diff = abs(mut_total - float(total_bank)) if total_bank else 0
                        if mut_diff <= config.JOURNAL_TOLERANCE:
                            mutation_matched = True
                except Exception:
                    pass

                items.append({
                    "row": row,
                    "bank": bank,
                    "group": group,
                    "tanggal": tanggal,
                    "payment_date": payment_date,
                    "merchant_amount": total_bank,
                    "odoo_amount": total_odoo,
                    "amount": total_odoo,
                    "selisih": selisih,
                    "mutation_found": mutation_found,
                    "mutation_matched": mutation_matched,
                    "mutation_amount": mut_total,
                    "reconciled": reconciled,
                    "journal_status": journal_status,
                    "edc_number": str(edc_num or "").strip(),
                    "ar_number": str(ar_num or "").strip(),
                    "status_valid": status_valid
                })

            if not items and set_status_fn:
                set_status_fn(f"No data with difference <= {config.JOURNAL_TOLERANCE}", ERROR)
                return
        except Exception as e:
            if set_status_fn:
                set_status_fn(f"Error reading excel: {e}", ERROR)
            return
        finally:
            try:
                wb.close()
            except Exception:
                pass

        journal_state.clear()
        for item in items:
            is_reconciled = str(item.get("reconciled", "")).strip().lower() == "yes"
            status_valid = item.get("status_valid", True)

            disabled_edc = False
            disabled_ar = False

            if not is_reconciled or not status_valid:
                disabled_edc = True
                disabled_ar = True

            if not item.get("mutation_matched", False):
                disabled_ar = True

            j_status = item.get("journal_status")
            if j_status:
                j_status_str = str(j_status).strip()
                if j_status_str not in ["", "None", "Not Yet", "-"]:
                    parts = [p.strip() for p in j_status_str.split("|")]
                    for p in parts:
                        if "(Both" in p:
                            if "Posted" in p or ("Draft" in p and "Difference" not in p):
                                disabled_edc = True
                                disabled_ar = True
                        elif "(EDC" in p:
                            if "Posted" in p or ("Draft" in p and "Difference" not in p):
                                disabled_edc = True
                        elif "(AR" in p:
                            if "Posted" in p or ("Draft" in p and "Difference" not in p):
                                disabled_ar = True

            var_item = tk.BooleanVar(value=not (disabled_edc and disabled_ar))
            var_edc = tk.BooleanVar(value=not disabled_edc)
            var_ar = tk.BooleanVar(value=not disabled_ar)

            journal_state.append({
                "item": item,
                "var_item": var_item,
                "var_edc": var_edc,
                "var_ar": var_ar,
                "disabled_edc": disabled_edc,
                "disabled_ar": disabled_ar
            })

    _load_data()

    ITEMS_PER_PAGE = 15
    current_page = [0]

    header_frame = ctk.CTkFrame(top, fg_color=PANEL, corner_radius=0, height=80, border_color=BORDER, border_width=1)
    header_frame.pack(fill="x", side="top")
    header_frame.pack_propagate(False)

    left_header = ctk.CTkFrame(header_frame, fg_color="transparent")
    left_header.pack(side="left", padx=24, pady=16)

    ctk.CTkLabel(left_header, text="Confirm Journal Creation", font=(FONT_FAMILY, 15, "bold"), text_color=TEXT).pack(anchor="w")
    ctk.CTkLabel(left_header, text="Review and select transactions to post. Expand any row to preview journal entries.", font=(FONT_FAMILY, 10, "bold"), text_color=MUTED).pack(anchor="w", pady=(2, 0))

    def _refresh_modal():
        _load_data()
        render_page(current_page[0])

    ctk.CTkButton(
        header_frame, text="🧩 Manual Match", height=32,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=ACCENT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=lambda: [top.destroy(), on_manual_match_fn() if on_manual_match_fn else None]
    ).pack(side="right", padx=(0, 12))

    ctk.CTkButton(
        header_frame, text="↻ Refresh Data", height=32,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=_refresh_modal
    ).pack(side="right", padx=(0, 24))

    footer_frame = ctk.CTkFrame(top, fg_color=PANEL, corner_radius=0, height=70, border_color=BORDER, border_width=1)
    footer_frame.pack(fill="x", side="bottom")
    footer_frame.pack_propagate(False)

    body_frame = ctk.CTkFrame(top, fg_color=BG, corner_radius=0)
    body_frame.pack(fill="both", expand=True, padx=24, pady=16)

    list_frame = ctk.CTkFrame(body_frame, fg_color=PANEL, corner_radius=8, border_color=BORDER, border_width=1)
    list_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(list_frame, bg=PANEL, highlightthickness=0)
    scrollbar = ctk.CTkScrollbar(list_frame, orientation="vertical", command=canvas.yview, button_color="#94A3B8", button_hover_color=ACCENT)
    scrollable_frame = tk.Frame(canvas, bg=PANEL)

    def _update_scrollregion(e=None):
        canvas.update_idletasks()
        bbox = canvas.bbox("all")
        if bbox:
            canvas.configure(scrollregion=(0, 0, bbox[2], bbox[3]))

    scrollable_frame.bind("<Configure>", _update_scrollregion)
    canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=4)

    footer_inner = ctk.CTkFrame(footer_frame, fg_color="transparent")
    footer_inner.pack(fill="both", expand=True, padx=24, pady=14)

    pagination_frame = ctk.CTkFrame(footer_inner, fg_color="transparent")
    pagination_frame.pack(side="left")

    btn_prev = ctk.CTkButton(
        pagination_frame, text="◄ Prev", width=80, height=34,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 11, "bold"), corner_radius=6
    )
    btn_prev.pack(side="left")

    lbl_page = ctk.CTkLabel(pagination_frame, text="", font=(FONT_FAMILY, 11, "bold"), text_color=TEXT)
    lbl_page.pack(side="left", padx=12)

    btn_next = ctk.CTkButton(
        pagination_frame, text="Next ►", width=80, height=34,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 11, "bold"), corner_radius=6
    )
    btn_next.pack(side="left")

    def _on_mousewheel(event):
        try:
            if not canvas.winfo_exists():
                return
            y_top, y_bot = canvas.yview()
            delta = event.delta
            if abs(delta) >= 120:
                delta = int(delta / 120)
            if delta > 0 and y_top <= 0.001:
                return
            if delta < 0 and y_bot >= 0.999:
                return
            if delta != 0:
                canvas.yview_scroll(int(-1 * delta), "units")
        except Exception:
            pass

    top.bind("<MouseWheel>", _on_mousewheel)
    canvas.bind("<MouseWheel>", _on_mousewheel)
    scrollable_frame.bind("<MouseWheel>", _on_mousewheel)

    def render_page(page_idx):
        active_det = {"btn": None, "frm": None}
        for widget in scrollable_frame.winfo_children():
            widget.destroy()

        col_widths = {
            0: 40,   1: 80,   2: 110,  3: 140,
            4: 130,  5: 130,  6: 145,  7: 100,
            8: 80,   9: 115,  10: 80,  11: 115
        }
        stretchy = {3, 9, 11}
        for col, w in col_widths.items():
            scrollable_frame.grid_columnconfigure(
                col, minsize=w, weight=2 if col in stretchy else 0
            )
        scrollable_frame.grid_columnconfigure(len(col_widths), weight=1)

        headers = ["", "Select", "Date", "Journal", "Merchant Amt", "Odoo Amt", "Mutation + Admin", "Difference", "EDC", "EDC Status", "AR", "AR Status"]
        for col, h in enumerate(headers):
            lbl_anchor = "w" if col in [2, 3] else "e" if col in [4, 5, 6, 7] else "center"
            tk.Label(
                scrollable_frame, text=h, bg=PREVIEW_BG, fg=MUTED,
                font=(FONT_FAMILY, 10, "bold"), anchor=lbl_anchor, padx=14
            ).grid(row=0, column=col, sticky="nsew", pady=(0, 4), ipady=7)

        tk.Frame(scrollable_frame, bg=BORDER_DARK, height=1).grid(row=1, column=0, columnspan=len(headers)+1, sticky="ew", pady=(0, 4))
        tk.Label(scrollable_frame, text="", bg=PREVIEW_BG).grid(row=0, column=len(headers), sticky="nsew", ipady=7)

        start_idx = page_idx * ITEMS_PER_PAGE
        end_idx = min(start_idx + ITEMS_PER_PAGE, len(journal_state))

        for i, state in enumerate(journal_state[start_idx:end_idx]):
            r_main = i * 3 + 2
            r_det  = i * 3 + 3
            r_sep  = i * 3 + 4
            bg_row = PANEL if i % 2 == 0 else "#FAFAFC"

            item = state["item"]
            var_item = state["var_item"]
            var_edc = state["var_edc"]
            var_ar = state["var_ar"]

            amt = float(item['amount']) if item['amount'] else 0
            sel = float(item['selisih']) if item['selisih'] else 0

            def _on_jurnal_toggle(vi=var_item, ve=var_edc, va=var_ar, de=state["disabled_edc"], da=state["disabled_ar"]):
                val = vi.get()
                if not de:
                    ve.set(val)
                if not da:
                    va.set(val)

            def _on_sub_toggle(vi=var_item, ve=var_edc, va=var_ar, de=state["disabled_edc"], da=state["disabled_ar"]):
                has_any = (bool(ve.get()) and not de) or (bool(va.get()) and not da)
                vi.set(1 if has_any else 0)

            b_name = str(item['bank']).lower()
            b_group = str(item['group']).lower()

            props = {}
            for a, p in config.BANK_ACCOUNTS.get(b_name, {}).items():
                if p.get("group", "").lower() == b_group:
                    props = p
                    break

            det_frame = tk.Frame(scrollable_frame, bg=PREVIEW_BG)

            edc_debit = props.get("edc_debit") or f"{str(item['bank']).upper()} EDC Debit"
            edc_credit = props.get("edc_credit") or f"{str(item['group'])} Credit"

            edc_frame = tk.Frame(det_frame, bg=PREVIEW_BG)
            edc_frame.pack(side="left", anchor="n", padx=20, pady=10)

            tk.Label(edc_frame, text="EDC Journal:", bg=PREVIEW_BG, fg=ACCENT, font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
            tk.Label(edc_frame, text="Debit:", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=1, column=0, sticky="w")
            tk.Label(edc_frame, text=edc_debit, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=1, column=1, sticky="w", padx=(10, 30))
            tk.Label(edc_frame, text=f"Rp {amt:,.0f}", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=1, column=2, sticky="e")
            tk.Label(edc_frame, text="Credit:", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=2, column=0, sticky="w")
            tk.Label(edc_frame, text=edc_credit, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=2, column=1, sticky="w", padx=(10, 30))
            tk.Label(edc_frame, text=f"Rp {amt:,.0f}", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=2, column=2, sticky="e")

            tk.Frame(edc_frame, bg=BORDER_DARK, height=1).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(4, 4))
            tk.Label(edc_frame, text=f"Total Debit: Rp {amt:,.0f}", bg=PREVIEW_BG, fg=SUCCESS, font=(FONT_FAMILY, 10, "bold")).grid(row=4, column=0, columnspan=2, sticky="w")
            tk.Label(edc_frame, text=f"Total Credit: Rp {amt:,.0f}", bg=PREVIEW_BG, fg=SUCCESS, font=(FONT_FAMILY, 10, "bold")).grid(row=4, column=2, sticky="e")

            if item.get("mutation_matched", False) or (item.get("mutation_found", False) and abs(sel) > 0.01):
                m_date = item['payment_date']
                m_group = item['group']
                m_raw = [m for m in mutation_raw if m["payment_date"] == m_date and m["group"] == m_group]
                a_raw = [a for a in admin_raw if a["payment_date"] == m_date and a["group"] == m_group]

                ar_debits = []
                for m in m_raw:
                    cat = (m["category"] or "").replace(" ", "").lower()
                    acc = props.get(f"ar_debit_{cat}") or props.get("ar_debit") or f"AR DEBIT ({cat})"
                    ar_debits.append({"account": acc, "amount": m["amount"]})

                admin_sums = defaultdict(float)
                for a in a_raw:
                    cat = (a["category"] or "").replace(" ", "").lower()
                    acc = props.get(f"admin_debit_{cat}") or props.get("admin_debit") or f"ADMIN DEBIT ({cat})"
                    admin_sums[acc] += a["amount"]

                for acc, a_amt in admin_sums.items():
                    if a_amt > 0:
                        ar_debits.append({"account": acc, "amount": a_amt})

                t_debit = sum(d["amount"] for d in ar_debits)
                t_credit = float(amt)
                t_diff = t_debit - t_credit

                ar_rows = []
                for d in ar_debits:
                    ar_rows.append(("Debit:", d['account'], d['amount']))

                if t_credit > 0:
                    ar_rows.append(("Credit:", edc_debit, t_credit))

                if round(t_diff, 2) > 0:
                    ar_rows.append(("Credit:", config.ODOO_ACCOUNT_BANK_DIFF_INCOME, abs(t_diff)))
                elif round(t_diff, 2) < 0:
                    ar_rows.append(("Debit:", config.ODOO_ACCOUNT_BANK_DIFF_LOSS, abs(t_diff)))

                ar_frame = tk.Frame(det_frame, bg=PREVIEW_BG)
                ar_frame.pack(side="left", anchor="n", padx=30, pady=10)

                tk.Label(ar_frame, text="AR Journal:", bg=PREVIEW_BG, fg=ACCENT, font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")

                tot_deb = sum(amt_v for typ, acc, amt_v in ar_rows if typ == "Debit:")
                tot_crd = sum(amt_v for typ, acc, amt_v in ar_rows if typ == "Credit:")

                for idx, (typ, acc, amt_val) in enumerate(ar_rows):
                    r = idx + 1
                    tk.Label(ar_frame, text=typ, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r, column=0, sticky="w")
                    tk.Label(ar_frame, text=acc, bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r, column=1, sticky="w", padx=(10, 30))
                    tk.Label(ar_frame, text=f"Rp {amt_val:,.0f}", bg=PREVIEW_BG, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r, column=2, sticky="e")

                last_r = len(ar_rows) + 1
                tot_color = SUCCESS if abs(tot_deb - tot_crd) < 0.01 else WARN
                tk.Frame(ar_frame, bg=BORDER_DARK, height=1).grid(row=last_r, column=0, columnspan=3, sticky="ew", pady=(4, 4))
                tk.Label(ar_frame, text=f"Total Debit: Rp {tot_deb:,.0f}", bg=PREVIEW_BG, fg=tot_color, font=(FONT_FAMILY, 10, "bold")).grid(row=last_r+1, column=0, columnspan=2, sticky="w")
                tk.Label(ar_frame, text=f"Total Credit: Rp {tot_crd:,.0f}", bg=PREVIEW_BG, fg=tot_color, font=(FONT_FAMILY, 10, "bold")).grid(row=last_r+1, column=2, sticky="e")

            def _toggle_det(btn, frm=det_frame, row_idx=r_det):
                if frm.winfo_ismapped():
                    frm.grid_remove()
                    btn.config(text="►")
                    if active_det["frm"] == frm:
                        active_det["btn"] = None
                        active_det["frm"] = None
                else:
                    prev_frm = active_det["frm"]
                    prev_btn = active_det["btn"]
                    if prev_frm and prev_frm.winfo_exists() and prev_frm.winfo_ismapped():
                        prev_frm.grid_remove()
                        if prev_btn and prev_btn.winfo_exists():
                            prev_btn.config(text="►")

                    frm.grid(row=row_idx, column=1, columnspan=10, sticky="w", pady=(8, 10))
                    btn.config(text="▼")
                    active_det["btn"] = btn
                    active_det["frm"] = frm

            btn_expand = tk.Label(scrollable_frame, text="►", bg=bg_row, fg=ACCENT, cursor="hand2", font=(FONT_FAMILY, 10, "bold"))
            btn_expand.bind("<Button-1>", lambda e, b=btn_expand, f=det_frame: _toggle_det(b, f))
            btn_expand.grid(row=r_main, column=0, padx=8, ipady=6, sticky="nsew")

            def _make_cb(parent_w, variable, bg_color, command=None):
                s = 18
                r = 3
                m = 1
                x, y = m, m
                w, h = s - 2*m, s - 2*m
                d = 2 * r

                cv = tk.Canvas(parent_w, width=s, height=s, bg=bg_color, highlightthickness=0, cursor="hand2")

                def _rrect(fill, outline, bw):
                    cv.delete("all")
                    if fill and fill != bg_color:
                        cv.create_rectangle(x+r, y, x+w-r, y+h, fill=fill, outline="")
                        cv.create_rectangle(x, y+r, x+w, y+h-r, fill=fill, outline="")
                        cv.create_oval(x, y, x+d, y+d, fill=fill, outline="")
                        cv.create_oval(x+w-d, y, x+w, y+d, fill=fill, outline="")
                        cv.create_oval(x, y+h-d, x+d, y+h, fill=fill, outline="")
                        cv.create_oval(x+w-d, y+h-d, x+w, y+h, fill=fill, outline="")
                    cv.create_arc(x, y, x+d, y+d, start=90, extent=90, style="arc", outline=outline, width=bw)
                    cv.create_arc(x+w-d, y, x+w, y+d, start=0, extent=90, style="arc", outline=outline, width=bw)
                    cv.create_arc(x, y+h-d, x+d, y+h, start=180, extent=90, style="arc", outline=outline, width=bw)
                    cv.create_arc(x+w-d, y+h-d, x+w, y+h, start=270, extent=90, style="arc", outline=outline, width=bw)
                    cv.create_line(x+r, y, x+w-r, y, fill=outline, width=bw)
                    cv.create_line(x+r, y+h, x+w-r, y+h, fill=outline, width=bw)
                    cv.create_line(x, y+r, x, y+h-r, fill=outline, width=bw)
                    cv.create_line(x+w, y+r, x+w, y+h-r, fill=outline, width=bw)

                def _draw(*_):
                    cv.delete("all")
                    if variable.get():
                        _rrect(fill=ACCENT, outline=ACCENT_DARK, bw=1)
                        cv.create_line(4, 9, 7, 13, fill=WHITE, width=2, capstyle="round", joinstyle="round")
                        cv.create_line(7, 13, 14, 5, fill=WHITE, width=2, capstyle="round", joinstyle="round")
                    else:
                        _rrect(fill=bg_color, outline="#94A3B8", bw=2)

                def _toggle(e):
                    variable.set(1 - variable.get())
                    _draw()
                    if command:
                        command()

                cv.bind("<Button-1>", _toggle)
                variable.trace_add("write", _draw)
                _draw()
                return cv

            if state["disabled_edc"] and state["disabled_ar"]:
                _cell1 = tk.Frame(scrollable_frame, bg=bg_row)
                _cell1.grid(row=r_main, column=1, sticky="nsew", pady=4)
                _cell1.grid_rowconfigure(0, weight=1)
                _cell1.grid_columnconfigure(0, weight=1)
                tk.Label(_cell1, text="—", bg=bg_row, fg=MUTED, font=(FONT_FAMILY, 11, "bold")).grid(row=0, column=0)
            else:
                _make_cb(scrollable_frame, var_item, bg_row, command=_on_jurnal_toggle).grid(row=r_main, column=1, pady=4)

            tk.Label(scrollable_frame, text=str(item['tanggal']), bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r_main, column=2, sticky="w", padx=14, ipady=6)
            tk.Label(scrollable_frame, text=str(item['group']), bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r_main, column=3, sticky="w", padx=14, ipady=6)
            amt_merch = float(item.get('merchant_amount') or 0)
            amt_odoo = float(item.get('odoo_amount') or 0)
            tk.Label(scrollable_frame, text=f"Rp {amt_merch:,.0f}", bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r_main, column=4, sticky="e", padx=14, ipady=6)
            tk.Label(scrollable_frame, text=f"Rp {amt_odoo:,.0f}", bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r_main, column=5, sticky="e", padx=14, ipady=6)
            mut_amt = float(item.get("mutation_amount", 0))
            tk.Label(scrollable_frame, text=f"Rp {mut_amt:,.0f}", bg=bg_row, fg=TEXT, font=(FONT_FAMILY, 10, "bold")).grid(row=r_main, column=6, sticky="e", padx=14, ipady=6)
            sel_color = WARN if sel != 0 else TEXT
            tk.Label(scrollable_frame, text=f"Rp {sel:,.0f}", bg=bg_row, fg=sel_color, font=(FONT_FAMILY, 10, "bold")).grid(row=r_main, column=7, sticky="e", padx=14, ipady=6)

            edc_badges = []
            ar_badges = []
            is_reconciled = str(item.get("reconciled", "")).strip().lower() == "yes"
            status_valid = item.get("status_valid", True)

            if not is_reconciled:
                edc_badges.append(("Unreconciled", False))
            elif not status_valid:
                edc_badges.append(("Difference", False))

            if not is_reconciled:
                ar_badges.append(("Unreconciled", False))
            elif not status_valid:
                ar_badges.append(("Difference", False))
            elif not item.get("mutation_matched", False):
                if not item.get("mutation_found", False):
                    ar_badges.append(("No Mutation", False))
                else:
                    ar_badges.append(("Mut Difference", False))

            j_status = item.get("journal_status")
            if j_status:
                j_status_str = str(j_status).strip()
                if j_status_str not in ["", "None", "Not Yet"]:
                    parts = [p.strip() for p in j_status_str.split("|")]
                    for p in parts:
                        if "Posted" in p:
                            label = "Posted (Diff)" if "Difference" in p else "Posted"
                            is_good = "Difference" not in p
                        elif "Draft" in p:
                            label = "Draft"
                            is_good = True
                        else:
                            label = p
                            is_good = False
                        if "(Both" in p:
                            if not edc_badges: edc_badges.append((label, is_good))
                            if not ar_badges: ar_badges.append((label, is_good))
                        elif "(EDC" in p:
                            if not edc_badges: edc_badges.append((label, is_good))
                        elif "(AR" in p:
                            if not ar_badges: ar_badges.append((label, is_good))

            edc_num_val = str(item.get("edc_number") or "").strip()
            ar_num_val = str(item.get("ar_number") or "").strip()

            is_edc_draft_no_diff = any(b[0] == "Draft" for b in edc_badges) and bool(edc_num_val and edc_num_val not in ["-", "None", ""])
            is_ar_draft_no_diff = any(b[0] == "Draft" for b in ar_badges) and bool(ar_num_val and ar_num_val not in ["-", "None", ""])

            if is_edc_draft_no_diff:
                _cell8 = tk.Frame(scrollable_frame, bg=bg_row)
                _cell8.grid(row=r_main, column=8, sticky="nsew", pady=4)
                _cell8.grid_rowconfigure(0, weight=1)
                _cell8.grid_columnconfigure(0, weight=1)
                ctk.CTkButton(
                    _cell8, text="🚀 Post", width=58, height=22,
                    fg_color=ACCENT, hover_color=ACCENT_DARK,
                    text_color=WHITE, font=(FONT_FAMILY, 9, "bold"),
                    corner_radius=4,
                    command=lambda num=edc_num_val, jt="EDC", dt=item['tanggal'], gp=item['group']: _post_single_draft(num, jt, dt, gp)
                ).grid(row=0, column=0)
            elif state["disabled_edc"]:
                _cell8 = tk.Frame(scrollable_frame, bg=bg_row)
                _cell8.grid(row=r_main, column=8, sticky="nsew", pady=4)
                _cell8.grid_rowconfigure(0, weight=1)
                _cell8.grid_columnconfigure(0, weight=1)
                tk.Label(_cell8, text="—", bg=bg_row, fg=MUTED, font=(FONT_FAMILY, 11, "bold")).grid(row=0, column=0)
            else:
                _make_cb(scrollable_frame, var_edc, bg_row, command=_on_sub_toggle).grid(row=r_main, column=8, pady=4)

            def _render_status_label(parent_w, col, badges, bg_row_col):
                if not badges:
                    return
                txt = "\n".join(t for t, _ in badges)
                good = all(g for _, g in badges)
                tk.Label(
                    parent_w, text=txt, bg=bg_row_col,
                    fg=SUCCESS if good else WARN,
                    font=(FONT_FAMILY, 10, "bold"),
                    justify="center", anchor="center"
                ).grid(row=r_main, column=col, sticky="nsew", padx=6, ipady=6)

            _render_status_label(scrollable_frame, 9, edc_badges, bg_row)

            if is_ar_draft_no_diff:
                _cell10 = tk.Frame(scrollable_frame, bg=bg_row)
                _cell10.grid(row=r_main, column=10, sticky="nsew", pady=4)
                _cell10.grid_rowconfigure(0, weight=1)
                _cell10.grid_columnconfigure(0, weight=1)
                ctk.CTkButton(
                    _cell10, text="🚀 Post", width=58, height=22,
                    fg_color=ACCENT, hover_color=ACCENT_DARK,
                    text_color=WHITE, font=(FONT_FAMILY, 9, "bold"),
                    corner_radius=4,
                    command=lambda num=ar_num_val, jt="AR", dt=item['tanggal'], gp=item['group']: _post_single_draft(num, jt, dt, gp)
                ).grid(row=0, column=0)
            elif state["disabled_ar"]:
                _cell10 = tk.Frame(scrollable_frame, bg=bg_row)
                _cell10.grid(row=r_main, column=10, sticky="nsew", pady=4)
                _cell10.grid_rowconfigure(0, weight=1)
                _cell10.grid_columnconfigure(0, weight=1)
                tk.Label(_cell10, text="—", bg=bg_row, fg=MUTED, font=(FONT_FAMILY, 11, "bold")).grid(row=0, column=0)
            else:
                _make_cb(scrollable_frame, var_ar, bg_row, command=_on_sub_toggle).grid(row=r_main, column=10, pady=4)

            _render_status_label(scrollable_frame, 11, ar_badges, bg_row)
            tk.Frame(scrollable_frame, bg=BORDER, height=1).grid(row=r_sep, column=0, columnspan=len(headers)+1, sticky="ew")

        total_pages = max(1, (len(journal_state) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
        lbl_page.configure(text=f"Page {page_idx + 1} of {total_pages}")
        btn_prev.configure(state="normal" if page_idx > 0 else "disabled")
        btn_next.configure(state="normal" if page_idx < total_pages - 1 else "disabled")
        canvas.yview_moveto(0)

    def _prev_page():
        if current_page[0] > 0:
            current_page[0] -= 1
            render_page(current_page[0])

    def _next_page():
        total_pages = (len(journal_state) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
        if current_page[0] < total_pages - 1:
            current_page[0] += 1
            render_page(current_page[0])

    btn_prev.configure(command=_prev_page)
    btn_next.configure(command=_next_page)

    render_page(0)

    def _get_selected_config():
        selected = []
        for v in journal_state:
            is_edc = bool(v["var_edc"].get()) and not v["disabled_edc"]
            is_ar = bool(v["var_ar"].get()) and not v["disabled_ar"]
            is_sel = bool(v["var_item"].get())
            if (is_sel and (is_edc or is_ar)) or is_edc or is_ar:
                selected.append({
                    "row": v["item"]["row"],
                    "edc": is_edc,
                    "ar": is_ar
                })
        return selected

    def _export(mode="edc"):
        selected = _get_selected_config()
        if not selected:
            messagebox.showwarning("Warning", "Select at least 1 transaction")
            return

        if mode == "ar":
            ar_items = [item for item in selected if item.get("ar", False)]
            if not ar_items:
                messagebox.showinfo(
                    "No AR Entries",
                    "None of the selected transactions have AR checked.\n\n"
                    "To export an AR journal, tick the AR checkbox on at least one transaction first."
                )
                return
        elif mode == "edc":
            edc_items = [item for item in selected if item.get("edc", False)]
            if not edc_items:
                messagebox.showinfo(
                    "No EDC Entries",
                    "None of the selected transactions have EDC checked.\n\n"
                    "To export an EDC journal, tick the EDC checkbox on at least one transaction first."
                )
                return

        config_path = base_dir / ".journal_config.json"
        if sys.platform == "win32" and config_path.exists():
            try:
                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                subprocess.run(["attrib", "-H", str(config_path)], check=False, capture_output=True, creationflags=flags)
            except Exception:
                pass
        config_path.write_text(json.dumps(selected))
        try:
            out_path = generate_journal_import(latest_file, config_path, mode=mode, is_preview=True)
            if out_path:
                messagebox.showinfo("Success", f"{mode.upper()} Journal exported:\n{out_path.name}")
                _open_path(str(out_path))
            else:
                messagebox.showwarning(
                    "Nothing Exported",
                    f"No {mode.upper()} journal rows were generated.\n\n"
                    "This can happen if all matching transactions are within the rounding tolerance "
                    "or if the reconciliation file has no data in the Daily Summary sheet."
                )
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting {mode.upper()}: {e}")

    def _process():
        selected = _get_selected_config()
        if not selected:
            messagebox.showwarning("Warning", "Select at least 1 transaction")
            return

        config_path = base_dir / ".journal_config.json"
        if sys.platform == "win32" and config_path.exists():
            try:
                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                subprocess.run(["attrib", "-H", str(config_path)], check=False, capture_output=True, creationflags=flags)
            except Exception:
                pass
        config_path.write_text(json.dumps(selected))

        recon_file = Path(latest_file)
        if not recon_file.exists():
            messagebox.showerror("Error", f"Reconciliation file not found at {recon_file.name}. Did you rename it while this window was open?")
            return

        try:
            out_path = generate_journal_import(recon_file, config_path, mode="both", is_preview=True)
            if not out_path:
                messagebox.showerror("Error", "Failed to generate combined preview file.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Error generating combined preview: {e}")
            return

        ar_count = sum(1 for item in selected if item.get("ar", False))
        edc_count = sum(1 for item in selected if item.get("edc", False))

        def _show_custom_confirm():
            dlg = ctk.CTkToplevel(top)
            dlg.withdraw()
            dlg.title("Confirm Upload")
            dlg.configure(fg_color=PANEL)
            dlg.transient(top)

            x = max(0, top.winfo_x() + (top.winfo_width() - 520) // 2)
            y = max(0, top.winfo_y() + (top.winfo_height() - 290) // 2)
            dlg.geometry(f"520x290+{x}+{y}")

            ctk.CTkLabel(dlg, text="Confirm Upload to Odoo", font=(FONT_FAMILY, 15, "bold"), text_color=TEXT).pack(pady=(24, 8))
            ctk.CTkLabel(
                dlg, text=f"Ready to import to Odoo.\n\nAR Journals: {ar_count}   |   EDC Journals: {edc_count}\n\nTo make manual edits before importing, click 'Edit Excel'.",
                justify="center", text_color=MUTED, font=(FONT_FAMILY, 10, "bold")
            ).pack(pady=(0, 20), padx=20)

            result = {"confirm": False}

            def on_upload():
                result["confirm"] = True
                dlg.destroy()

            def on_cancel():
                dlg.destroy()

            btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
            btn_frame.pack(fill="x", padx=20, pady=10)

            ctk.CTkButton(
                btn_frame, text="Cancel", height=36,
                fg_color=PANEL, hover_color=PREVIEW_BG,
                border_color=BORDER_DARK, border_width=1,
                text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
                corner_radius=6, command=on_cancel
            ).pack(side="left", padx=4, expand=True, fill="x")

            ctk.CTkButton(
                btn_frame, text="Edit Excel", height=36,
                fg_color=PANEL, hover_color=PREVIEW_BG,
                border_color=BORDER_DARK, border_width=1,
                text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
                corner_radius=6, command=lambda: _open_path(str(out_path))
            ).pack(side="left", padx=4, expand=True, fill="x")

            ctk.CTkButton(
                btn_frame, text="Upload to Odoo", height=36,
                fg_color=SUCCESS, hover_color=SUCCESS_DARK,
                text_color=WHITE, font=(FONT_FAMILY, 10, "bold"),
                corner_radius=6, command=on_upload
            ).pack(side="left", padx=4, expand=True, fill="x")

            dlg.update_idletasks()
            dlg.deiconify()
            dlg.grab_set()
            top.wait_window(dlg)
            return result["confirm"]

        confirm = _show_custom_confirm()
        if not confirm:
            return

        top.destroy()

        if set_status_fn:
            set_status_fn("Uploading Edited Journal to Odoo...", WARN)

        def run_script():
            try:
                env = os.environ.copy()
                env.pop("TCL_LIBRARY", None)
                env.pop("TK_LIBRARY", None)
                env["PYTHONIOENCODING"] = "utf-8"
                env["PYTHONUTF8"] = "1"

                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if IS_WINDOWS else 0

                if getattr(sys, "frozen", False):
                    cmd = [
                        sys.executable, "--run-journal-creator",
                        "--file", str(recon_file),
                        "--import-file", str(out_path),
                        "--config", str(config_path)
                    ]
                else:
                    cmd = [
                        venv_python, "odoo_journal_creator.py",
                        "--file", str(recon_file),
                        "--import-file", str(out_path),
                        "--config", str(config_path)
                    ]

                proc = subprocess.Popen(
                    cmd, cwd=str(base_dir),
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, encoding="utf-8",
                    env=env, creationflags=flags
                )

                if proc.stdout:
                    for line in iter(proc.stdout.readline, ''):
                        if log_write_fn:
                            parent.after(0, log_write_fn, line, "")
                proc.wait()

                if proc.returncode == 0:
                    def _auto_recon_followup():
                        if on_done_fn:
                            on_done_fn(0, None)
                        if log_write_fn:
                            log_write_fn("\n🔄 Automatically re-running Reconciliation to sync newly created draft journals...\n", "head")
                        if on_run_fn:
                            on_run_fn()
                    parent.after(0, _auto_recon_followup)
                else:
                    if on_done_fn:
                        parent.after(0, on_done_fn, proc.returncode, None)

            except Exception as e:
                if log_write_fn:
                    parent.after(0, log_write_fn, f"ERROR: {str(e)}\n", "err")
                if on_done_fn:
                    parent.after(0, on_done_fn, 1, None)

        threading.Thread(target=run_script, daemon=True).start()

    footer_right = ctk.CTkFrame(footer_inner, fg_color="transparent")
    footer_right.pack(side="right")

    def _select_all():
        for v in journal_state:
            if not (v["disabled_edc"] and v["disabled_ar"]):
                v["var_item"].set(True)
            if not v["disabled_edc"]:
                v["var_edc"].set(True)
            if not v["disabled_ar"]:
                v["var_ar"].set(True)
        render_page(current_page[0])

    def _deselect_all():
        for v in journal_state:
            v["var_item"].set(False)
            if not v["disabled_edc"]:
                v["var_edc"].set(False)
            if not v["disabled_ar"]:
                v["var_ar"].set(False)
        render_page(current_page[0])

    tools_frame = ctk.CTkFrame(pagination_frame, fg_color="transparent")
    tools_frame.pack(side="left", padx=(20, 0))

    ctk.CTkButton(
        tools_frame, text="Select All", height=32, width=80,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=_select_all
    ).pack(side="left", padx=(0, 6))

    ctk.CTkButton(
        tools_frame, text="Deselect All", height=32, width=85,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=_deselect_all
    ).pack(side="left")

    def _post_single_draft(move_number: str, jrn_type: str, date_str: str, group_str: str):
        clean_num = str(move_number).strip()
        draft_info = get_draft_journal_details(clean_num, jrn_type=jrn_type, date_val=date_str, group_val=group_str)
        target_id = draft_info["id"] if draft_info else clean_num
        disp_ref = draft_info["name"] if draft_info else (clean_num if clean_num not in ["/", "-"] else f"Draft {jrn_type} Entry")
        disp_jrn = draft_info["journal"] if draft_info else f"{jrn_type} Settlement Journal"
        disp_date = draft_info["date"] if draft_info else date_str
        disp_amt = draft_info["amount_total"] if draft_info else 0.0
        draft_lines = draft_info.get("lines", []) if draft_info else []

        def _show_custom_post_confirm():
            dlg = ctk.CTkToplevel(top)
            dlg.withdraw()
            dlg.title("Confirm Post Journal")
            dlg.minsize(500, 460)
            dlg.configure(fg_color=PANEL)
            dlg.transient(top)

            x = max(0, top.winfo_x() + (top.winfo_width() - 540) // 2)
            y = max(0, top.winfo_y() + (top.winfo_height() - 510) // 2)
            dlg.geometry(f"540x510+{x}+{y}")

            ctk.CTkLabel(dlg, text="Confirm Post to Odoo", font=(FONT_FAMILY, 16, "bold"), text_color=TEXT).pack(pady=(20, 2))
            ctk.CTkLabel(dlg, text="Review draft settlement journal before posting to Odoo.", font=(FONT_FAMILY, 10), text_color=MUTED).pack(pady=(0, 14))

            info_card = ctk.CTkFrame(dlg, fg_color=BG, corner_radius=8, border_color=BORDER, border_width=1)
            info_card.pack(fill="x", padx=24, pady=(0, 12))

            overview_rows = [
                ("Reference", disp_ref, ACCENT),
                ("Target Journal", disp_jrn, TEXT),
                ("Date", disp_date, TEXT),
                ("Total Amount", f"Rp {disp_amt:,.0f}" if disp_amt > 0 else "-", TEXT),
                ("State Transition", "🟡 Draft   ➔   🟢 Posted", SUCCESS),
            ]
            for r_i, (lbl_txt, val_txt, val_clr) in enumerate(overview_rows):
                ctk.CTkLabel(info_card, text=lbl_txt, font=(FONT_FAMILY, 10, "bold"), text_color=MUTED, anchor="w").grid(row=r_i, column=0, sticky="w", padx=14, pady=3)
                ctk.CTkLabel(info_card, text=val_txt, font=(FONT_FAMILY, 10, "bold"), text_color=val_clr, anchor="e").grid(row=r_i, column=1, sticky="e", padx=14, pady=3)
            info_card.grid_columnconfigure(1, weight=1)

            if draft_lines:
                lines_frame = ctk.CTkFrame(dlg, fg_color=BG, corner_radius=8, border_color=BORDER, border_width=1)
                lines_frame.pack(fill="both", expand=True, padx=24, pady=(0, 16))

                hdr_f = tk.Frame(lines_frame, bg=BG)
                hdr_f.pack(fill="x", padx=12, pady=(8, 4))
                tk.Label(hdr_f, text="Account", bg=BG, fg=MUTED, font=(FONT_FAMILY, 9, "bold")).pack(side="left")
                tk.Label(hdr_f, text="Credit", bg=BG, fg=MUTED, font=(FONT_FAMILY, 9, "bold"), width=14, anchor="e").pack(side="right")
                tk.Label(hdr_f, text="Debit", bg=BG, fg=MUTED, font=(FONT_FAMILY, 9, "bold"), width=14, anchor="e").pack(side="right")

                tk.Frame(lines_frame, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(0, 4))

                for l in draft_lines:
                    row_f = tk.Frame(lines_frame, bg=BG)
                    row_f.pack(fill="x", padx=12, pady=2)
                    acc_lbl = tk.Label(row_f, text=l["account"], bg=BG, fg=TEXT, font=(FONT_FAMILY, 9, "bold"), anchor="w")
                    acc_lbl.pack(side="left", fill="x", expand=True)

                    cr_str = f"Rp {l['credit']:,.0f}" if l['credit'] > 0 else "-"
                    dr_str = f"Rp {l['debit']:,.0f}" if l['debit'] > 0 else "-"
                    tk.Label(row_f, text=cr_str, bg=BG, fg=TEXT if l['credit'] > 0 else MUTED, font=(FONT_FAMILY, 9, "bold"), width=14, anchor="e").pack(side="right")
                    tk.Label(row_f, text=dr_str, bg=BG, fg=TEXT if l['debit'] > 0 else MUTED, font=(FONT_FAMILY, 9, "bold"), width=14, anchor="e").pack(side="right")

            result = {"confirm": False}

            def on_confirm():
                result["confirm"] = True
                dlg.destroy()

            def on_cancel():
                dlg.destroy()

            btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
            btn_frame.pack(fill="x", side="bottom", padx=24, pady=(0, 20))

            ctk.CTkButton(
                btn_frame, text="Cancel", height=38, width=120,
                fg_color=PANEL, hover_color=PREVIEW_BG,
                border_color=BORDER_DARK, border_width=1,
                text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
                corner_radius=6, command=on_cancel
            ).pack(side="left", padx=(0, 10))

            ctk.CTkButton(
                btn_frame, text="🚀 Post to Odoo", height=38,
                fg_color=ACCENT, hover_color=ACCENT_DARK,
                text_color=WHITE, font=(FONT_FAMILY, 10, "bold"),
                corner_radius=6, command=on_confirm
            ).pack(side="left", expand=True, fill="x")

            dlg.update_idletasks()
            dlg.deiconify()
            dlg.grab_set()
            top.wait_window(dlg)
            return result["confirm"]

        if not _show_custom_post_confirm():
            return

        def _run_single_post():
            try:
                res = post_draft_settlement_journals([target_id])
                if res.get("success"):
                    if log_write_fn:
                        log_write_fn(f"\n✅ Successfully posted {jrn_type} journal ({disp_ref}) in Odoo!\n", "ok")
                    if set_status_fn:
                        set_status_fn(f"Posted {jrn_type} journal {disp_ref}", SUCCESS)

                    try:
                        check_journals(latest_file, skip_download=True)
                    except Exception:
                        pass

                    def _done_ui():
                        _refresh_modal()
                        if update_dashboard_fn:
                            update_dashboard_fn(skip_drill=True)
                    parent.after(0, _done_ui)
                else:
                    err = res.get("error", "Unknown error")
                    if log_write_fn:
                        log_write_fn(f"\n❌ Failed to post {disp_ref}: {err}\n", "err")
                    if set_status_fn:
                        set_status_fn(f"Post failed: {err}", ERROR)
            except Exception as ex:
                if log_write_fn:
                    log_write_fn(f"\n❌ Error during post: {ex}\n", "err")

        threading.Thread(target=_run_single_post, daemon=True).start()

    ctk.CTkButton(
        footer_right, text="Cancel", height=36, width=80,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=top.destroy
    ).pack(side="left", padx=(0, 8))

    ctk.CTkButton(
        footer_right, text="Export EDC", height=36, width=95,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=lambda: _export("edc")
    ).pack(side="left", padx=(0, 6))

    ctk.CTkButton(
        footer_right, text="Export AR", height=36, width=90,
        fg_color=PANEL, hover_color=PREVIEW_BG,
        border_color=BORDER_DARK, border_width=1,
        text_color=TEXT, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=lambda: _export("ar")
    ).pack(side="left", padx=(0, 12))

    ctk.CTkButton(
        footer_right, text="Submit", height=36, width=100,
        fg_color=SUCCESS, hover_color=SUCCESS_DARK,
        text_color=WHITE, font=(FONT_FAMILY, 10, "bold"),
        corner_radius=6, command=_process
    ).pack(side="left")

    top.update_idletasks()
    top.deiconify()
    top.grab_set()
    return top
