import unittest

from pdf_parser import InvoiceParser


class CamariCustItemParsingTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
