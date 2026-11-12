import unittest
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from openpyxl import load_workbook

from logic import (
    calculate_grade,
    check_question_answer,
    hash_password,
    shuffle_question,
    validate_practice_file,
    validate_question,
    verify_password,
)
from practice_variants import (
    PRACTICE_NUMBERS,
    PRACTICES_WITH_VARIANTS,
    build_practice_csv,
    build_practice_file,
)


class PasswordTests(unittest.TestCase):
    def test_password_hash_verifies_only_matching_password(self) -> None:
        stored = hash_password("secure-password")
        self.assertTrue(verify_password("secure-password", stored))
        self.assertFalse(verify_password("wrong-password", stored))

    def test_invalid_hash_is_rejected(self) -> None:
        self.assertFalse(verify_password("password", "not-a-valid-hash"))


class GradeTests(unittest.TestCase):
    def test_grade_thresholds(self) -> None:
        self.assertEqual(calculate_grade(9, 10), (90, 5))
        self.assertEqual(calculate_grade(7, 10), (70, 4))
        self.assertEqual(calculate_grade(5, 10), (50, 3))
        self.assertEqual(calculate_grade(4, 10), (40, 2))

    def test_empty_test_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            calculate_grade(0, 0)


class QuestionTests(unittest.TestCase):
    def test_single_answer_is_preserved_when_options_are_shuffled(self) -> None:
        question = validate_question(
            "single", "Какой ответ верный?", ["неверно", "верно"], [1]
        )

        class ReverseRandom:
            @staticmethod
            def shuffle(values) -> None:
                values.reverse()

        shuffled = shuffle_question(question, ReverseRandom())
        self.assertEqual(shuffled["options"], ["верно", "неверно"])
        self.assertTrue(check_question_answer(shuffled, 0))
        self.assertFalse(check_question_answer(shuffled, 1))

    def test_multiple_answer_requires_exact_selection(self) -> None:
        question = validate_question(
            "multiple", "Выбери все верные", ["А", "Б", "В"], [0, 2]
        )
        self.assertTrue(check_question_answer(question, {0, 2}))
        self.assertFalse(check_question_answer(question, {0}))
        self.assertFalse(check_question_answer(question, {0, 1, 2}))

    def test_multiple_answers_are_remapped_after_option_shuffle(self) -> None:
        question = validate_question(
            "multiple", "Выбери все верные", ["А", "Б", "В"], [0, 1]
        )

        class ReverseRandom:
            @staticmethod
            def shuffle(values) -> None:
                values.reverse()

        shuffled = shuffle_question(question, ReverseRandom())
        self.assertEqual(shuffled["options"], ["В", "Б", "А"])
        self.assertEqual(shuffled["correct_answers"], [1, 2])
        self.assertTrue(check_question_answer(shuffled, {1, 2}))
        self.assertFalse(check_question_answer(shuffled, {0, 2}))

    def test_text_answer_ignores_case_and_outer_whitespace(self) -> None:
        question = validate_question("text", "Ответ", [], ["  Информация  "])
        self.assertTrue(check_question_answer(question, "  ИНФОРМАЦИЯ "))
        self.assertFalse(check_question_answer(question, "информационные данные"))

    def test_single_question_requires_exactly_one_correct_answer(self) -> None:
        with self.assertRaisesRegex(ValueError, "ровно один"):
            validate_question("single", "Вопрос", ["А", "Б"], [0, 1])


class PracticeFileTests(unittest.TestCase):
    def test_supported_file_is_accepted(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "work.xlsx"
            path.write_bytes(b"placeholder")
            self.assertEqual(validate_practice_file(path), path)

    def test_unsupported_file_is_rejected(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "work.pdf"
            path.write_bytes(b"placeholder")
            with self.assertRaisesRegex(ValueError, "Word или Excel"):
                validate_practice_file(path)


class PracticeVariantTests(unittest.TestCase):
    def test_all_practice_templates_generate_valid_office_files(self) -> None:
        for practice_number in PRACTICE_NUMBERS:
            with self.subTest(practice_number=practice_number):
                extension, content = build_practice_file(practice_number, 7)
                self.assertGreater(len(content), 1000)
                if extension == ".docx":
                    document = Document(BytesIO(content))
                    text = "\n".join(
                        paragraph.text for paragraph in document.paragraphs
                    )
                    self.assertIn("ВАРИАНТ 07", text)
                else:
                    workbook = load_workbook(
                        BytesIO(content), read_only=True
                    )
                    self.assertIn("Инструкция", workbook.sheetnames)
                    self.assertIn(
                        "Данные" if practice_number == 8 else "Мой вариант",
                        workbook.sheetnames,
                    )
                    workbook.close()

    def test_user_manual_template_is_shared_and_uses_student_project(self) -> None:
        self.assertNotIn(2, PRACTICES_WITH_VARIANTS)
        extension, content = build_practice_file(2, None)
        self.assertEqual(extension, ".docx")
        document = Document(BytesIO(content))
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("ПРАКТИЧЕСКОЕ ЗАНЯТИЕ №2", text)
        self.assertNotIn("ВАРИАНТ", text)
        self.assertIn("практической работе №1", text)

    def test_practice_one_document_changes_with_selected_variant(self) -> None:
        _, variant_one = build_practice_file(1, 1)
        _, variant_two = build_practice_file(1, 2)

        text_one = "\n".join(
            paragraph.text for paragraph in Document(BytesIO(variant_one)).paragraphs
        )
        text_two = "\n".join(
            paragraph.text for paragraph in Document(BytesIO(variant_two)).paragraphs
        )

        self.assertNotEqual(variant_one, variant_two)
        self.assertIn("RoomPlan", text_one)
        self.assertIn("HelpDesk Lite", text_two)

    def test_all_thirty_variants_have_distinct_assignment_data(self) -> None:
        for practice_number in PRACTICES_WITH_VARIANTS:
            signatures = []
            for variant in range(1, 31):
                extension, content = build_practice_file(practice_number, variant)
                if extension == ".docx":
                    document = Document(BytesIO(content))
                    if practice_number == 1:
                        signature = tuple(
                            paragraph.text
                            for paragraph in document.paragraphs
                            if paragraph.text.startswith(
                                (
                                    "Сценарий:",
                                    "Разработайте техническое задание",
                                    "Обязательное бизнес-правило:",
                                    "Критерий приёмки:",
                                )
                            )
                        )
                    else:
                        signature = tuple(
                            cell.text
                            for row in document.tables[1].rows[1:]
                            for cell in row.cells
                        )
                else:
                    workbook = load_workbook(BytesIO(content), data_only=False)
                    if practice_number == 5:
                        sheet = workbook["Мой вариант"]
                        signature = tuple(
                            cell
                            for row in sheet.iter_rows(
                                min_row=3, max_row=5, values_only=True
                            )
                            for cell in row
                        )
                    elif practice_number == 6:
                        sheet = workbook["Сырые данные"]
                        signature = tuple(
                            cell
                            for row in sheet.iter_rows(min_row=3, values_only=True)
                            for cell in row[1:]
                        )
                    elif practice_number == 7:
                        sheet = workbook["Мой вариант"]
                        signature = tuple(
                            cell
                            for row in sheet.iter_rows(
                                min_row=4, max_row=10, max_col=7, values_only=True
                            )
                            for cell in row
                        )
                    elif practice_number == 8:
                        sheet = workbook["Данные"]
                        signature = tuple(
                            cell
                            for row in sheet.iter_rows(min_row=3, values_only=True)
                            for cell in row[1:]
                        )
                    else:
                        sheet = workbook["Мой вариант"]
                        signature = tuple(
                            cell
                            for row in sheet.iter_rows(
                                min_row=3, max_row=8, values_only=True
                            )
                            for cell in row
                        )
                    workbook.close()
                signatures.append(signature)

            with self.subTest(practice_number=practice_number):
                self.assertEqual(len(set(signatures)), 30)

    def test_practices_that_need_distinct_data_require_a_variant(self) -> None:
        for practice_number in PRACTICES_WITH_VARIANTS:
            with self.subTest(practice_number=practice_number):
                with self.assertRaisesRegex(ValueError, "необходимо указать"):
                    build_practice_file(practice_number, None)

    def test_budget_variant_has_correct_calculation_formulas(self) -> None:
        _, content = build_practice_file(5, 1)
        workbook = load_workbook(BytesIO(content), data_only=False)
        sheet = workbook["Мой вариант"]
        self.assertEqual(sheet["E3"].value, "=C3*D3")
        self.assertEqual(sheet["E6"].value, "=SUM(E3:E5)")
        self.assertEqual(sheet["E7"].value, "=E6*B7")
        self.assertEqual(sheet["E8"].value, "=E6+E7")
        workbook.close()

    def test_variant_csv_contains_variant_specific_log_ids(self) -> None:
        content = build_practice_csv(7).decode("utf-8-sig")
        self.assertIn("LOG-07-", content)
        self.assertIn("ID;Временная метка;Уровень;Модуль;Сообщение", content)

    def test_variant_range_is_limited_to_thirty(self) -> None:
        for invalid_variant in (0, 31):
            with self.subTest(variant=invalid_variant):
                with self.assertRaisesRegex(ValueError, "от 1 до 30"):
                    build_practice_file(1, invalid_variant)


if __name__ == "__main__":
    unittest.main()
