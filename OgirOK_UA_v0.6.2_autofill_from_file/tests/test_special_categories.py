import os
import tempfile
import unittest
from pathlib import Path

from docx import Document

# Import from project
from apps.parser_orders.src.core.docx_parser import parse_docx_to_structured


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = PROJECT_ROOT / "templates" / "N_fixed.xlsx"


def _make_docx(paragraphs):
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    doc = Document()
    for p in paragraphs:
        doc.add_paragraph(p)
    doc.save(path)
    return path


class TestSpecialCategories(unittest.TestCase):
    def test_annual_leave_category(self):
        docx = _make_docx([
            "З щорічної основної відпустки:",
            "Солдат ІВАНЕНКО Іван Іванович прибув у розпорядження."
        ])
        data, _ = parse_docx_to_structured(docx, str(TEMPLATE), learning_dir=str(PROJECT_ROOT / "config" / "learning"),
                                          category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "event_categories_rules_v1.json"),
                                          user_category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "user_category_rules.json"))
        os.unlink(docx)

        records = data.get("records") or []
        self.assertTrue(records, "No records parsed")
        self.assertEqual(records[0].get("category"), "З щорічної основної відпустки")

    def test_detach_med_category(self):
        docx = _make_docx([
            'Нижчепойменованих військовослужбовців, у зв\'язку з службовою необхідністю, прикомандирувати до медичної роти військової частини А0989:',
            "Солдат ПЕТРЕНКО Петро Петрович прикомандирувати."
        ])
        data, _ = parse_docx_to_structured(docx, str(TEMPLATE), learning_dir=str(PROJECT_ROOT / "config" / "learning"),
                                          category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "event_categories_rules_v1.json"),
                                          user_category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "user_category_rules.json"))
        os.unlink(docx)
        records = data.get("records") or []
        self.assertTrue(records, "No records parsed")
        self.assertEqual(records[0].get("category"), "Прикомандировано до медичної роти (в/ч А0989)")

    def test_trip_category(self):
        docx = _make_docx([
            "Вважати такими, що вибули:",
            "У відрядження:",
            "Солдат СИДОРЕНКО Сидір Сидорович вибув у відрядження."
        ])
        data, _ = parse_docx_to_structured(docx, str(TEMPLATE), learning_dir=str(PROJECT_ROOT / "config" / "learning"),
                                          category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "event_categories_rules_v1.json"),
                                          user_category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "user_category_rules.json"))
        os.unlink(docx)
        records = data.get("records") or []
        self.assertTrue(records, "No records parsed")
        self.assertEqual(records[0].get("category"), "У відрядження")

    def test_ignore_kmu_phrase(self):
        docx = _make_docx([
            "Постанова Кабінету Міністрів України №123 від 01.01.2026",
            "ІВАНЕНКО Іван Іванович (ця строка не повинна парситись як людина)",
        ])
        data, _ = parse_docx_to_structured(docx, str(TEMPLATE), learning_dir=str(PROJECT_ROOT / "config" / "learning"),
                                          category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "event_categories_rules_v1.json"),
                                          user_category_rules_path=str(PROJECT_ROOT / "config" / "categories" / "user_category_rules.json"))
        os.unlink(docx)
        people = data.get("people_index") or {}
        self.assertFalse(people, "People were parsed from ignored KМУ block")

if __name__ == "__main__":
    unittest.main()
