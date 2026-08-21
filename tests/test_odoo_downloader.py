import unittest
import os
import sys
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

from odoo_downloader import download_via_xmlrpc
from readers.odoo_reader import read_odoo
from config import BANK_ACCOUNTS, ODO_EXCEL_PATH, ODO_JOURNAL_EXCEL_PATH

class TestOdooDownloaderXMLRPC(unittest.TestCase):
    def test_xmlrpc_download_and_reader_compatibility(self):
        # Download July 6, 2026 data via direct XML-RPC (06/07/2026)
        success = download_via_xmlrpc(
            date_from="06/07/2026",
            date_to="06/07/2026",
            banks="BCA,Mandiri,BRI",
            status="Posted",
            mode="both"
        )
        self.assertTrue(success, "XML-RPC download should return True")
        self.assertTrue(ODO_EXCEL_PATH.exists(), f"{ODO_EXCEL_PATH} should exist")
        self.assertTrue(ODO_JOURNAL_EXCEL_PATH.exists(), f"{ODO_JOURNAL_EXCEL_PATH} should exist")

        # Verify readers/odoo_reader.py can read the generated file without error
        group_map = {}
        for bank_name, bank_conf in BANK_ACCOUNTS.items():
            for grp in bank_conf.get("odoo_groups", []):
                group_map[grp] = bank_name

        parsed_date, bank_txns = read_odoo(
            excel_path=ODO_EXCEL_PATH,
            amount_col="Amount Signed",
            group_map=group_map,
            number_col="Number",
            reference_col="Reference",
            include_others=True
        )

        self.assertEqual(parsed_date.isoformat(), "2026-07-06")
        total_txns = sum(len(txns) for txns in bank_txns.values())
        self.assertGreater(total_txns, 0, "Should have parsed transactions")

if __name__ == "__main__":
    unittest.main()
