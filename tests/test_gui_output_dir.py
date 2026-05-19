import unittest
from pathlib import Path

from output_paths import default_output_dir_for_base_dir


class DefaultOutputDirTest(unittest.TestCase):
    def test_uses_project_output_dir_during_local_development(self):
        base_dir = Path("/Users/michael/customs_doc_generator")

        output_dir = default_output_dir_for_base_dir(
            base_dir,
            app_name="ClearanceOS",
            home_dir=Path("/Users/tester"),
            is_frozen=False,
        )

        self.assertEqual(base_dir / "output", output_dir)

    def test_uses_documents_dir_for_translocated_macos_app(self):
        base_dir = Path(
            "/private/var/folders/wn/gmxph2g94jj3yz5lhzt823c00000gn/T/"
            "AppTranslocation/88EE7108-EB90-4AE3-AC47-DCEC7812AEAD/d/"
            "ClearanceOS.app/Contents/Frameworks"
        )

        output_dir = default_output_dir_for_base_dir(
            base_dir,
            app_name="ClearanceOS",
            home_dir=Path("/Users/tester"),
            is_frozen=True,
        )

        self.assertEqual(
            Path("/Users/tester/Documents/ClearanceOS/output"),
            output_dir,
        )


if __name__ == "__main__":
    unittest.main()
