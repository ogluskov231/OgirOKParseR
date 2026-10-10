import os
import tempfile
import unittest
from docx import Document

from apps.parser_orders.src.core.docx_parser import parse_docx_to_structured

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(PROJECT_ROOT, "templates", "N_fixed.xlsx")

def make_docx(lines):
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    return path

class TestJsonSchemaIntegration(unittest.TestCase):
    def test_schema_document_is_valid(self):
        path = make_docx([
            "9. Полковника Ткаченка Олексія Миколайовича призначити на посаду командира бригади.",
            "Підстава: витяг з наказу від 01.06.2025 № 190/ос."
        ])
        try:
            data, _ = parse_docx_to_structured(path, TEMPLATE)
            self.assertIn("structured_schema", data)
            self.assertTrue(data["schema_validation"]["valid"], data["schema_validation"]["errors"])
            self.assertEqual(data["structured_schema"]["items"][0]["action"]["type"], "appointment")
        finally:
            os.unlink(path)

if __name__ == "__main__":
    unittest.main()
