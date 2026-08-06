import unittest

from eur_invoice_converter import apply_eur_conversion
from eur_invoice_pdf_overlay import _build_replacement_map
from pdf_parser import InvoiceParser


class CamariCustItemParsingTest(unittest.TestCase):
    def test_prefers_cny_terms_currency_over_eur_bank_account_text(self):
        parser = InvoiceParser("__dummy__.pdf")
        parser._fmt = "camari_cust"
        parser.raw_text = """
Due Date
Currency
CNY
Sales Order
￥377.00
Subtotal
Total
33090050201000003421 (EUR)
"""

        self.assertEqual("CNY", parser._extract_currency())

    def test_parses_usd_item_lines(self):
        parser = InvoiceParser("__dummy__.pdf")
        parser.raw_text = """
 #   Item                                            Quantity    Units        Unit Price       Tax Rate    Amount

 1  Phone Case                                661       PCS        $12.00       0%         $7,932.00
     n.w 21.8kg
 2   Buckle Bands                              50        PCS        $12.00       0%         $600.00
     n.w 0.7kg
 3   Spectacle Case                             20        PCS        $13.50       0%         $270.00
     n.w 2.1kg
 4   Keychain                                  100       PCS        $4.00        0%         $400.00
     n.w 0.6kg
 5   Keychain                                  300       PCS        $1.50        0%         $450.00
    Hardware for Keychain
     n.w 5.7kg

                                                                                     Subtotal           $9,652.00
"""

        items = parser._parse_camari_cust_items()

        self.assertEqual(5, len(items))
        self.assertEqual(
            {
                "line_no": "1",
                "item_code": "",
                "item_code_prefix": "",
                "description": "Phone Case",
                "quantity": "661",
                "unit": "PCS",
                "unit_price": "12.00",
                "amount": "7932.00",
            },
            items[0],
        )
        self.assertEqual("Keychain", items[4]["description"])
        self.assertEqual("450.00", items[4]["amount"])

    def test_extracts_transport_cost_and_fullwidth_cny_total(self):
        parser = InvoiceParser("__dummy__.pdf")
        parser._fmt = "camari_cust"
        parser.raw_text = """
 #   Item                   Quantity Units Unit Price Tax Rate Amount
 1 NAPPA 3251               32.5 SQM ￥280.80 0% ￥9,126.00
 4 Transport Cost           1 ￥1,992.12 0% ￥1,992.12
 Subtotal ￥21,451.56
 Tax Total ￥0.00
 Total ￥21,451.56
 Bank Information:
"""

        items = parser._parse_camari_cust_items()

        self.assertEqual(1, len(items))
        self.assertEqual("1992.12", parser._extract_transport_cost())
        self.assertEqual("21451.56", parser._compute_total(items))

    def test_eur_conversion_includes_transport_and_invoice_total(self):
        original = {
            "currency": "CNY",
            "items": [
                {
                    "description": "NAPPA 3251",
                    "quantity": "32.5",
                    "unit": "SQM",
                    "unit_price": "280.80",
                    "amount": "9126.00",
                }
            ],
            "customs_items": [],
            "transport_cost": "1992.12",
            "total_amount": "21451.56",
        }

        converted = apply_eur_conversion(original, 7.8)
        replacements = _build_replacement_map(original, converted)

        self.assertEqual("255.40", converted["transport_cost"])
        self.assertEqual("2750.20", converted["total_amount"])
        self.assertEqual("€255,40", replacements["￥1,992.12"])
        self.assertEqual("€2 750,20", replacements["￥21,451.56"])


if __name__ == "__main__":
    unittest.main()
