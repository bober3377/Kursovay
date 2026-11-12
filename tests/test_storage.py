import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from main import TextbookApp


class SharedWorkStorageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary_directory.cleanup)
        self.root = Path(self.temporary_directory.name)
        self.student_id = 42
        self.student_folder = self.root / str(self.student_id)
        self.student_folder.mkdir()
        app = TextbookApp.__new__(TextbookApp)
        app.db = SimpleNamespace(
            work_root=lambda: self.root,
            work_dir=lambda student_id: self.root / str(student_id),
        )
        self.app = app

    def test_resolves_file_stored_relative_to_shared_root(self) -> None:
        work_file = self.student_folder / "submission.docx"
        work_file.write_bytes(b"work")

        result = self.app.resolve_submission_path(
            {
                "student_id": self.student_id,
                "relative_path": "42/submission.docx",
            }
        )

        self.assertEqual(result, work_file.resolve())

    def test_resolves_legacy_student_work_path_after_copy_to_share(self) -> None:
        work_file = self.student_folder / "submission.docx"
        work_file.write_bytes(b"work")

        result = self.app.resolve_submission_path(
            {
                "student_id": self.student_id,
                "relative_path": "student_work/42/submission.docx",
            }
        )

        self.assertEqual(result, work_file.resolve())

    def test_rejects_submission_path_outside_student_folder(self) -> None:
        other_file = self.root / "43" / "submission.docx"
        other_file.parent.mkdir()
        other_file.write_bytes(b"other student's work")

        with self.assertRaises(FileNotFoundError):
            self.app.resolve_submission_path(
                {
                    "student_id": self.student_id,
                    "relative_path": "../43/submission.docx",
                }
            )


if __name__ == "__main__":
    unittest.main()
