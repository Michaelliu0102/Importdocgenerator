import tempfile
import textwrap
import unittest
from pathlib import Path

from config_loader import ConfigLoader


class ConfigLoaderInputNormalizationTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.config_path = Path(self.temp_dir.name) / "config.yaml"
        self.config_path.write_text(
            textwrap.dedent(
                """
                suppliers:
                  demo_supplier:
                    name: Demo Supplier
                product_categories:
                  demo_product:
                    name: Demo Product
                    category: Demo Category
                    sub_category: Demo Subcategory
                """
            ).strip(),
            encoding="utf-8",
        )
        self.loader = ConfigLoader(str(self.config_path))

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_match_supplier_by_name_treats_none_as_no_match(self):
        self.assertIsNone(self.loader.match_supplier_by_name(None))


if __name__ == "__main__":
    unittest.main()
