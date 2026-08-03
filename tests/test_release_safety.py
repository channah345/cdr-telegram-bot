import ast
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class ReleaseSafetyTests(unittest.TestCase):
    def test_sensitive_sample_workbook_is_not_packaged(self):
        self.assertFalse((ROOT / "jobs.csv").exists())

    def test_polling_preserves_queued_updates(self):
        source = (ROOT / "bot.py").read_text(encoding="utf-8")
        self.assertNotIn("drop_pending_updates=True", source)
        self.assertNotIn("PicklePersistence", source)
        self.assertNotIn("cdn.jsdelivr.net/npm/signature_pad", source)

    def test_python_sources_parse(self):
        for path in ROOT.rglob("*.py"):
            ast.parse(path.read_text(encoding="utf-8"), filename=str(path))


if __name__ == "__main__":
    unittest.main()
