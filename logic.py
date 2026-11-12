import hashlib
import hmac
import random
import secrets
import unicodedata
from pathlib import Path
from typing import Any


PASSWORD_ITERATIONS = 310_000
ALLOWED_PRACTICE_EXTENSIONS = {".doc", ".docx", ".xls", ".xlsx"}
MAX_UPLOAD_SIZE = 20 * 1024 * 1024


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt, PASSWORD_ITERATIONS
    )
    return f"{salt.hex()}${digest.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt_hex, digest_hex = stored_hash.split("$", maxsplit=1)
        salt = bytes.fromhex(salt_hex)
        expected_digest = bytes.fromhex(digest_hex)
    except (ValueError, TypeError):
        return False

    actual_digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt, PASSWORD_ITERATIONS
    )
    return hmac.compare_digest(actual_digest, expected_digest)


def calculate_grade(correct_count: int, question_count: int) -> tuple[int, int]:
    if question_count <= 0:
        raise ValueError("В тесте должен быть хотя бы один вопрос.")
    if not 0 <= correct_count <= question_count:
        raise ValueError("Количество правильных ответов вне допустимого диапазона.")

    percentage = round(correct_count * 100 / question_count)
    if percentage >= 85:
        grade = 5
    elif percentage >= 70:
        grade = 4
    elif percentage >= 50:
        grade = 3
    else:
        grade = 2
    return percentage, grade


def validate_question(
    question_type: str,
    prompt: str,
    options: list[str],
    correct_answers: list[int] | list[str],
) -> dict[str, Any]:
    question_type = question_type.strip().lower()
    prompt = prompt.strip()
    options = [option.strip() for option in options if option.strip()]
    if not prompt:
        raise ValueError("Введите текст вопроса.")

    if question_type in {"single", "multiple"}:
        if len(options) < 2:
            raise ValueError("Добавьте не менее двух вариантов ответа.")
        if len({option.casefold() for option in options}) != len(options):
            raise ValueError("Варианты ответа не должны повторяться.")
        if not all(isinstance(answer, int) for answer in correct_answers):
            raise ValueError("Отметьте правильный вариант ответа.")
        correct = sorted(set(correct_answers))
        if not correct or any(index < 0 or index >= len(options) for index in correct):
            raise ValueError("Отметьте правильный вариант ответа.")
        if question_type == "single" and len(correct) != 1:
            raise ValueError("Для вопроса с одним ответом отметьте ровно один вариант.")
        return {
            "question_type": question_type,
            "prompt": prompt,
            "options": options,
            "correct_answers": correct,
        }

    if question_type == "text":
        answers = [
            unicodedata.normalize("NFKC", answer).strip().casefold()
            for answer in correct_answers
            if isinstance(answer, str) and answer.strip()
        ]
        if not answers:
            raise ValueError("Укажите хотя бы один правильный текстовый ответ.")
        return {
            "question_type": "text",
            "prompt": prompt,
            "options": [],
            "correct_answers": list(dict.fromkeys(answers)),
        }

    raise ValueError("Выберите тип вопроса.")


def shuffle_question(
    question: dict[str, Any], rng: random.Random | Any | None = None
) -> dict[str, Any]:
    prepared = dict(question)
    if prepared["question_type"] == "text":
        return prepared

    randomizer = rng or random
    correct = set(
        prepared.get("correct_answers")
        or [prepared["correct_index"]]
    )
    indexed_options = list(enumerate(prepared["options"]))
    randomizer.shuffle(indexed_options)
    prepared["options"] = [option for _, option in indexed_options]
    prepared["correct_answers"] = [
        display_index
        for display_index, (original_index, _option) in enumerate(indexed_options)
        if original_index in correct
    ]
    prepared["correct_index"] = prepared["correct_answers"][0]
    return prepared


def check_question_answer(question: dict[str, Any], answer: Any) -> bool:
    correct_answers = question.get("correct_answers")
    if correct_answers is None:
        correct_answers = [question["correct_index"]]

    question_type = question["question_type"]
    if question_type == "single":
        return answer in correct_answers
    if question_type == "multiple":
        try:
            return set(answer) == set(correct_answers)
        except TypeError:
            return False
    if question_type == "text" and isinstance(answer, str):
        normalized = unicodedata.normalize("NFKC", answer).strip().casefold()
        accepted = {
            unicodedata.normalize("NFKC", item).strip().casefold()
            for item in correct_answers
        }
        return normalized in accepted
    return False


def validate_practice_file(file_path: str | Path) -> Path:
    path = Path(file_path)
    if path.suffix.lower() not in ALLOWED_PRACTICE_EXTENSIONS:
        raise ValueError("Можно загрузить файл Word или Excel.")
    if not path.is_file():
        raise ValueError("Выбранный файл не найден.")
    if path.stat().st_size > MAX_UPLOAD_SIZE:
        raise ValueError("Размер файла не должен превышать 20 МБ.")
    return path
