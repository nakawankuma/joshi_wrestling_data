import tempfile
import unittest
from pathlib import Path

import pandas as pd

from xlsx_to_html_converter_all_in_one import (
    PROMOTION_NAMES_END,
    PROMOTION_NAMES_START,
    XlsxToHtmlConverter,
)


class XlsxToHtmlConverterTest(unittest.TestCase):
    def setUp(self):
        self.converter = XlsxToHtmlConverter.__new__(XlsxToHtmlConverter)
        self.converter.config = {
            "encoding": "utf-8",
            "js_variable_name": "wrestlerData",
        }

    def make_dataframe(self, rows):
        return pd.DataFrame(rows)

    def test_marker_replacement_allows_terminator_inside_json_string(self):
        content = "// START\nconst data = [\"name ]; value\"];\n// END"
        replaced = self.converter.replace_generated_block(
            content,
            "// START",
            "// END",
            'const data = ["updated ]; value"];',
            "test data",
        )

        parsed = self.converter.extract_generated_json(
            replaced,
            "// START",
            "// END",
            "data",
            "test data",
        )
        self.assertEqual(parsed, ["updated ]; value"])

    def test_missing_marker_is_rejected(self):
        with self.assertRaisesRegex(Exception, "生成マーカー"):
            self.converter.replace_generated_block(
                "const data = [];",
                "// START",
                "// END",
                "const data = [];",
                "test data",
            )

    def test_blank_and_duplicate_promotion_names_are_rejected(self):
        blank = self.make_dataframe([
            [None, "デビュー年", "A", None],
            [None, 2026, "選手A", "選手B"],
        ])
        duplicate = self.make_dataframe([
            [None, "デビュー年", "A", "A"],
            [None, 2026, "選手A", "選手B"],
        ])

        with self.assertRaisesRegex(Exception, "団体名が空"):
            self.converter.extract_promotion_headers(blank)
        with self.assertRaisesRegex(Exception, "団体名が重複"):
            self.converter.extract_promotion_headers(duplicate)

    def test_invalid_source_rows_are_rejected(self):
        missing_year = self.make_dataframe([
            [None, "デビュー年", "A"],
            [None, None, "選手A"],
        ])
        duplicate_year = self.make_dataframe([
            [None, "デビュー年", "A"],
            [None, 2026, "選手A"],
            [None, 2026, "選手B"],
        ])
        non_integer_year = self.make_dataframe([
            [None, "デビュー年", "A"],
            [None, 2026.5, "選手A"],
        ])
        duplicate_wrestler = self.make_dataframe([
            [None, "デビュー年", "A"],
            [None, 2025, "選手A"],
            [None, 2026, "選手A"],
        ])

        with self.assertRaisesRegex(Exception, "年度なし"):
            self.converter.validate_source_data(
                missing_year,
                ["A"],
                {"A": ["選手A"]},
            )
        with self.assertRaisesRegex(Exception, "年度ラベルが重複"):
            self.converter.validate_source_data(
                duplicate_year,
                ["A"],
                {"A": ["選手A", "選手B"]},
            )
        with self.assertRaisesRegex(Exception, "整数ではありません"):
            self.converter.validate_source_data(
                non_integer_year,
                ["A"],
                {"A": ["選手A"]},
            )
        with self.assertRaisesRegex(Exception, "選手名が重複"):
            self.converter.validate_source_data(
                duplicate_wrestler,
                ["A"],
                {"A": ["選手A", "選手A"]},
            )

    def test_promotion_header_is_escaped_without_inline_javascript(self):
        source = (
            '<table><tr><th class="sortable desc">デビュー年</th>'
            '<th class="promotion-header">Old</th></tr></table>\n'
            f'{PROMOTION_NAMES_START}\n'
            'const promotionNames = ["Old"];\n'
            f'{PROMOTION_NAMES_END}\n'
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            html_file = Path(temp_dir) / "index.html"
            html_file.write_text(source, encoding="utf-8")
            self.converter.update_html_headers(
                ["O'Reilly <Test>"],
                str(html_file),
            )
            result = html_file.read_text(encoding="utf-8")

        self.assertIn("O&#x27;Reilly &lt;Test&gt;", result)
        self.assertIn('data-promotion-index="1"', result)
        self.assertNotIn("onclick=", result)
        parsed = self.converter.extract_generated_json(
            result,
            PROMOTION_NAMES_START,
            PROMOTION_NAMES_END,
            "promotionNames",
            "promotion names",
        )
        self.assertEqual(parsed, ["O'Reilly <Test>"])


if __name__ == "__main__":
    unittest.main()
