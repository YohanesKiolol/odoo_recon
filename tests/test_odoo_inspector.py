import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import unittest
from unittest.mock import patch, MagicMock
import odoo_inspector


class TestOdooInspector(unittest.TestCase):

    def test_normalize_date_to_iso(self):
        self.assertEqual(odoo_inspector._normalize_date_to_iso("15/07/2026"), "2026-07-15")
        self.assertEqual(odoo_inspector._normalize_date_to_iso("2026-07-15"), "2026-07-15")
        self.assertEqual(odoo_inspector._normalize_date_to_iso("15-07-2026"), "2026-07-15")
        self.assertEqual(odoo_inspector._normalize_date_to_iso(""), "")

    def test_inspect_bank_only_invalid_input(self):
        res = odoo_inspector.inspect_bank_only(amount=0, date_str="")
        self.assertFalse(res["success"])
        self.assertEqual(res["invoices"], [])

    @patch("odoo_inspector._execute_kw")
    def test_inspect_bank_only_draft_invoice(self, mock_kw):
        # Mock Odoo returning a draft invoice
        mock_kw.return_value = [{
            'id': 102,
            'name': 'INV/2026/0002',
            'invoice_date': '2026-08-15',
            'amount_total': 500000.0,
            'payment_state': 'not_paid',
            'state': 'draft',
            'partner_id': [1, 'Test Partner'],
        }]

        res = odoo_inspector.inspect_bank_only(amount=500000.0, date_str="2026-08-15")
        self.assertTrue(res["success"])
        self.assertEqual(res["found_count"], 1)
        inv = res["invoices"][0]
        self.assertEqual(inv["status_code"], "DRAFT_INVOICE")
        self.assertIn("Draft", inv["badge"])

    @patch("odoo_inspector._execute_kw")
    def test_inspect_bank_only_unpaid_open(self, mock_kw):
        # Mock Odoo returning open unpaid invoice
        mock_kw.return_value = [{
            'id': 103,
            'name': 'INV/2026/0003',
            'invoice_date': '2026-08-15',
            'amount_total': 500000.0,
            'payment_state': 'not_paid',
            'state': 'posted',
            'partner_id': [1, 'Test Partner'],
        }]

        res = odoo_inspector.inspect_bank_only(amount=500000.0, date_str="2026-08-15")
        self.assertTrue(res["success"])
        inv = res["invoices"][0]
        self.assertEqual(inv["status_code"], "AVAILABLE_OPEN")
        self.assertIn("Unpaid", inv["badge"])

    @patch("odoo_inspector._execute_kw")
    def test_inspect_bank_only_filters_out_paid(self, mock_kw):
        # Mock Odoo returning paid invoice -> should be filtered out
        mock_kw.return_value = [{
            'id': 104,
            'name': 'INV/2026/0004',
            'invoice_date': '2026-08-15',
            'amount_total': 500000.0,
            'payment_state': 'paid',
            'state': 'posted',
            'partner_id': [1, 'Test Partner'],
        }]

        res = odoo_inspector.inspect_bank_only(amount=500000.0, date_str="2026-08-15")
        self.assertTrue(res["success"])
        self.assertEqual(res["found_count"], 0)
        self.assertEqual(res["invoices"], [])

    @patch("odoo_inspector._execute_kw")
    def test_inspect_odoo_only_linked_invoice_number(self, mock_kw):
        # First call returns payment with reconciled_invoice_ids, second call returns invoice
        mock_kw.side_effect = [
            [{
                'id': 201,
                'name': 'PCSH2/2026/02092',
                'date': '2026-08-16',
                'amount': 119400.0,
                'state': 'posted',
                'partner_id': [124, 'WALK IN CUSTOMER'],
                'journal_id': [24, 'Petty Cash Seminyak'],
                'ref': 'INV/2026/19047',
                'reconciled_invoice_ids': [331082],
                'move_id': [331083, 'PCSH2/2026/02092 (INV/2026/19047)']
            }],
            [{
                'id': 331082,
                'name': 'INV/2026/19047',
                'invoice_date': '2026-08-16',
                'amount_total': 119400.0,
                'state': 'posted',
                'payment_state': 'paid',
                'partner_id': [124, 'WALK IN CUSTOMER']
            }]
        ]

        res = odoo_inspector.inspect_odoo_only(odoo_number="PCSH2/2026/02092", amount=119400.0)
        self.assertTrue(res["success"])
        self.assertTrue(res["found"])
        self.assertEqual(len(res["linked_invoices"]), 1)
        self.assertEqual(res["linked_invoices"][0]["invoice_number"], "INV/2026/19047")

    @patch("odoo_inspector._execute_kw")
    def test_inspect_unreconciled_other_payment(self, mock_kw):
        # 1st call returns payment with ref, 2nd call returns invoice with widget having other payment
        mock_kw.side_effect = [
            [{
                'id': 301,
                'name': 'PCSH1/2026/001',
                'date': '2026-08-15',
                'amount': 300000.0,
                'state': 'posted',
                'partner_id': [3, 'Partner X'],
                'journal_id': [5, 'Cash'],
                'ref': 'INV/2026/0401',
                'move_id': [302, 'PCSH1/2026/001 (INV/2026/0401)']
            }],
            [{
                'id': 401,
                'name': 'INV/2026/0401',
                'invoice_date': '2026-08-15',
                'amount_total': 300000.0,
                'payment_state': 'paid',
                'state': 'posted',
                'partner_id': [3, 'Partner X'],
                'invoice_payments_widget': {
                    'content': [{
                        'ref': 'PBNK1/2026/0500',
                        'date': '2026-08-15',
                        'amount': 300000.0,
                        'journal_name': 'BCA EDC'
                    }]
                }
            }]
        ]

        res = odoo_inspector.inspect_unreconciled(odoo_number="PCSH1/2026/001", amount=300000.0, date_str="2026-08-15")
        self.assertTrue(res["success"])
        self.assertEqual(res["invoices_found"], 1)
        self.assertTrue(res["has_other_payment"])
        inv = res["invoices"][0]
        self.assertEqual(inv["badge"], "🔴 Linked to Another Payment")
        self.assertEqual(inv["other_payments"][0]["ref"], "PBNK1/2026/0500")


if __name__ == "__main__":
    unittest.main()
