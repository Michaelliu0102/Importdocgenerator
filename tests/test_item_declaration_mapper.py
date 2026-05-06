import unittest

from item_declaration_mapper import _match_excel_row


class MatchExcelRowTest(unittest.TestCase):
    def test_matches_explicit_item_rows_even_if_appended_after_special_rows(self):
        rows = [
            {"row_no": 2, "item": "Watch Band", "product_name": "表带"},
            {"row_no": 21, "item": "5012/5015", "product_name": "面料"},
            {"row_no": 22, "item": "NAPPA/VERONA", "product_name": "牛皮"},
            {"row_no": 23, "item": "FA3018/FA3055", "product_name": "面料"},
            {"row_no": 26, "item": "MF", "product_name": "超细纤维革"},
            {"row_no": 27, "item": "Buckle Bands", "product_name": "表带"},
            {"row_no": 28, "item": "Keychain", "product_name": "钥匙扣"},
        ]

        self.assertEqual(
            "Buckle Bands",
            _match_excel_row({"description": "Buckle Bands"}, rows)["item"],
        )
        self.assertEqual(
            "Keychain",
            _match_excel_row({"description": "Keychain"}, rows)["item"],
        )

    def test_preserves_special_prefix_matching_without_fixed_row_numbers(self):
        rows = [
            {"row_no": 27, "item": "Buckle Bands", "product_name": "表带"},
            {"row_no": 28, "item": "Keychain", "product_name": "钥匙扣"},
            {"row_no": 29, "item": "5012/5015", "product_name": "面料"},
            {"row_no": 30, "item": "NAPPA/VERONA", "product_name": "牛皮"},
            {"row_no": 31, "item": "FA3018/FA3055", "product_name": "面料"},
            {"row_no": 32, "item": "MF", "product_name": "超细纤维革"},
        ]

        self.assertEqual(
            "5012/5015",
            _match_excel_row({"item_code": "50123333"}, rows)["item"],
        )
        self.assertEqual(
            "NAPPA/VERONA",
            _match_excel_row({"description": "VERONA 5002"}, rows)["item"],
        )
        self.assertEqual(
            "FA3018/FA3055",
            _match_excel_row({"description": "FA3018 BLACK"}, rows)["item"],
        )
        self.assertEqual(
            "MF",
            _match_excel_row({"description": "MF 1387 1.2mm"}, rows)["item"],
        )


if __name__ == "__main__":
    unittest.main()
