import importlib.util
import os
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]


def load_module(name: str, relative_path: str):
    spec = importlib.util.spec_from_file_location(name, REPO_ROOT / relative_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


briefing = load_module("outlook_briefing_data", "scripts/outlook_briefing_data.py")
helper = load_module("outlook_helper", "scripts/outlook_helper.py")


class OutlookHelperTests(unittest.TestCase):
    def test_windows_style_env_vars_expand_on_any_platform(self):
        os.environ["MAIL_AUTOMATION_TEST_ROOT"] = str(REPO_ROOT)
        expanded = helper.expand_path("%MAIL_AUTOMATION_TEST_ROOT%/scripts")
        self.assertEqual(expanded, REPO_ROOT / "scripts")

    def test_zero_max_items_returns_empty_without_auth(self):
        self.assertEqual(helper.fetch_recent_messages_http(hours=24, max_items=0), [])

    def test_dedupe_keeps_distinct_messages_without_entry_id(self):
        messages = [
            {"sender": "A", "topic": "Topic", "received": "2026-04-11 09:00", "preview": "first"},
            {"sender": "A", "topic": "Topic", "received": "2026-04-11 09:00", "preview": "second"},
        ]
        self.assertEqual(len(briefing.dedupe(messages)), 2)


if __name__ == "__main__":
    unittest.main()
