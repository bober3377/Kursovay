import unittest
from io import BytesIO

from docx import Document
from docx.oxml.ns import qn

from practice_template import build_practice_template


class PracticeTemplateTests(unittest.TestCase):
    def test_template_has_example_and_independent_assignment(self) -> None:
        document = Document(BytesIO(build_practice_template()))
        paragraphs = [paragraph.text for paragraph in document.paragraphs]
        combined_text = "\n".join(paragraphs)

        self.assertIn("Пример выполнения", combined_text)
        self.assertIn("ТВОЯ РАБОТА", combined_text)
        self.assertIn("заполни функцию каждого устройства", combined_text)
        self.assertEqual(len(document.tables), 2)
        self.assertEqual(len(document.tables[0].rows), 3)
        self.assertEqual(len(document.tables[1].rows), 5)

    def test_assignment_starts_on_a_separate_page(self) -> None:
        document = Document(BytesIO(build_practice_template()))
        page_breaks = document.element.body.findall(
            ".//" + qn("w:br")
        )

        self.assertTrue(
            any(
                page_break.get(qn("w:type")) == "page"
                for page_break in page_breaks
            )
        )


if __name__ == "__main__":
    unittest.main()
