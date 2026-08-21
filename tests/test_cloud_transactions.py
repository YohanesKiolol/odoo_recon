import unittest
import cloud_sync


class TestCloudTransactions(unittest.TestCase):
    def test_merchant_hash(self):
        h1 = cloud_sync.generate_recon_hash("eyerizz", "BCA", "2026-08-30", "123456", 150000.0, "Credit")
        h2 = cloud_sync.generate_recon_hash("eyerizz", "BCA", "2026-08-30", "123456", 150000.0, "Credit")
        self.assertEqual(h1, h2)
        self.assertEqual(len(h1), 64)

    def test_mutation_hash(self):
        h1 = cloud_sync.generate_mutation_hash("eyerizz", "BCA", "7571188817", "2026-08-30", 500000.0, "CR", "TRSF E-BANKING")
        h2 = cloud_sync.generate_mutation_hash("eyerizz", "BCA", "7571188817", "2026-08-30", 500000.0, "CR", "TRSF E-BANKING")
        self.assertEqual(h1, h2)
        self.assertEqual(len(h1), 64)


if __name__ == "__main__":
    unittest.main()
