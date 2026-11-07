import unittest

from wc_product_bot import build_payload, slugify


class FakeConfig:
    publish_status = "draft"
    copy_fields = ["type", "regular_price", "manage_stock"]
    slug_template = "{brand}-{model}-{board_no}"
    set_focus_keyword = True
    focus_keyword_meta_key = "rank_math_focus_keyword"


class PayloadTests(unittest.TestCase):
    def setUp(self):
        self.cfg = FakeConfig()
        self.template = {
            "type": "simple",
            "regular_price": "12.50",
            "manage_stock": True,
            "categories": [{"id": 7, "name": "BIOS"}],
            "images": [{"id": 21, "src": "https://example.com/image.jpg"}],
        }
        self.data = {
            "name": "Fallback title",
            "brand": "Dell",
            "model": "Latitude 5400",
            "board_no": "LA-G891P",
            "capacity": "16MB",
        }

    def test_slugify_normalizes_spaces_and_punctuation(self):
        self.assertEqual(slugify(" Dell  Latitude! 5400 "), "dell-latitude-5400")

    def test_build_payload_uses_templates_and_copies_selected_fields(self):
        payload = build_payload(
            self.cfg,
            self.template,
            self.data,
            "BIOS {brand} {model}",
            "Description for {board_no}",
            "Short {capacity}",
        )

        self.assertEqual(payload["name"], "BIOS Dell Latitude 5400")
        self.assertEqual(payload["description"], "Description for LA-G891P")
        self.assertEqual(payload["short_description"], "Short 16MB")
        self.assertEqual(payload["slug"], "dell-latitude-5400-la-g891p")
        self.assertEqual(payload["status"], "draft")
        self.assertEqual(payload["regular_price"], "12.50")
        self.assertEqual(payload["categories"], [{"id": 7}])
        self.assertEqual(payload["images"], [{"id": 21}])
        self.assertEqual(
            payload["meta_data"],
            [{"key": "rank_math_focus_keyword", "value": "BIOS Dell Latitude 5400"}],
        )

    def test_missing_optional_title_template_uses_name_column(self):
        payload = build_payload(
            self.cfg, self.template, self.data, None, "Description", None
        )
        self.assertEqual(payload["name"], "Fallback title")
        self.assertNotIn("short_description", payload)


if __name__ == "__main__":
    unittest.main()
