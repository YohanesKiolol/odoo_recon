import unittest
import cloud_sync


class TestCloudSync(unittest.TestCase):
    def test_recon_hash_determinism(self):
        # Same transaction must produce identical hash
        h1 = cloud_sync.generate_recon_hash("eyerizz", "BCA", "2026-08-30", "123456", 150000.0, "bank_only")
        h2 = cloud_sync.generate_recon_hash("eyerizz", "BCA", "2026-08-30", "123456", 150000.0, "bank_only")
        self.assertEqual(h1, h2)
        self.assertEqual(len(h1), 64)

    def test_recon_hash_distinctness(self):
        # Different amount or different bank must produce different hash
        h1 = cloud_sync.generate_recon_hash("eyerizz", "BCA", "2026-08-30", "123456", 150000.0, "bank_only")
        h2 = cloud_sync.generate_recon_hash("eyerizz", "BCA", "2026-08-30", "123456", 200000.0, "bank_only")
        h3 = cloud_sync.generate_recon_hash("company2", "BCA", "2026-08-30", "123456", 150000.0, "bank_only")
        self.assertNotEqual(h1, h2)
        self.assertNotEqual(h1, h3)

    def test_device_id(self):
        dev = cloud_sync.get_device_id()
        self.assertTrue(bool(dev))


if __name__ == "__main__":
    unittest.main()
