import configparser
import json
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from logic import hash_password


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    RESOURCE_DIR = APP_DIR
CONFIG_PATH = APP_DIR / "config.ini"
SCHEMA_PATH = RESOURCE_DIR / "schema.sql"

LESSONS = [
    (
        "office-docs",
        "Word и техническая документация",
        """В IT-проекте документы фиксируют требования, решения и правила работы с программой. Техническое задание (ТЗ) описывает цель продукта, его функции, ограничения и критерии приёмки. Для оформления технической документации применяют стандарты ЕСПД, в том числе ГОСТ 19 и ГОСТ 34.

Стили заголовков помогают поддерживать единое оформление и автоматически собирать оглавление. Перекрёстная ссылка ведёт к другому разделу или рисунку. Рецензирование и Track Changes позволяют видеть правки и обсуждать их, не теряя исходный текст.

Термины: ЕСПД — система стандартов программной документации; ТЗ — требования к разрабатываемой системе; стиль — набор параметров оформления; Track Changes — запись исправлений; Markdown — простой текст с разметкой, который можно экспортировать в DOCX или PDF.

На практике: создайте ТЗ с заголовками, оглавлением и нумерацией; для руководства добавьте схему алгоритма и подписанные снимки экрана.""",
    ),
    (
        "office-spreadsheets",
        "Excel для анализа и планирования",
        """Табличный процессор помогает считать бюджет, анализировать журналы и планировать работу IT-команды. Формулы используют адреса ячеек: относительная ссылка меняется при копировании, а абсолютная (например, $B$2) остаётся закреплённой. Функции SUM, IF и AVERAGE выполняют типовые расчёты.

CSV — текстовый формат обмена таблицами. Перед анализом данные очищают от лишних пробелов и дублей, а условное форматирование выделяет важные значения. Диаграмма Ганта показывает сроки этапов; сводная таблица группирует и подсчитывает записи. Надстройка «Поиск решения» подбирает значения с учётом заданных ограничений.

Термины: ссылка — адрес ячейки в формуле; CSV — строки и столбцы, разделённые символами; сводная таблица — инструмент группировки данных; диаграмма Ганта — календарный план; ограничение — условие, которое должно выполняться при оптимизации.

На практике: рассчитайте стоимость проекта, очистите журнал ошибок, постройте план работ, проанализируйте баг-трекер или распределите задачи между специалистами.""",
    ),
    (
        "office-databases",
        "Базы данных и SQL",
        """Реляционная база данных хранит сведения в связанных таблицах. Например, в системе учёта ПО можно отдельно хранить программы, лицензии и пользователей. Таблица описывает сущность, столбцы — её атрибуты, строки — отдельные записи.

Первичный ключ однозначно определяет запись, внешний ключ связывает её с записью другой таблицы. Нормализация уменьшает дублирование и связанные с ним ошибки. SQL используют для выборки и изменения данных: SELECT читает записи, INSERT добавляет, UPDATE изменяет, DELETE удаляет.

Термины: сущность — объект учёта; атрибут — его свойство; первичный ключ — уникальный идентификатор; внешний ключ — ссылка на другую таблицу; связь «один-ко-многим» — одной записи соответствуют несколько записей другой таблицы; СУБД — программа для создания и управления базами.

На практике: спроектируйте таблицы учёта лицензий, задайте ключи и связи, затем проверьте данные запросами.""",
    ),
    (
        "office-collaboration",
        "Презентации и командная работа",
        """ИТ-презентация должна быстро объяснять продукт и его ценность. В коротком питче показывают проблему, решение, аудиторию, экономику и демонстрацию продукта. Схемы, макеты интерфейса и графики должны помогать понять мысль, а не перегружать слайды.

Scrum делит работу на короткие спринты, Kanban показывает задачи на доске по их статусу. Бэклог — упорядоченный список будущей работы; задачу полезно описывать, назначать ответственного и срок. В совместных облачных документах важно выдавать только необходимые права и защищать учётную запись двухфакторной аутентификацией.

Термины: Product Owner отвечает за ценность продукта и приоритеты; Scrum Master помогает команде соблюдать процесс; QA проверяет качество; эпик — крупная задача, которую делят на меньшие; 2FA — подтверждение входа двумя способами.

На практике: разложите эпик на задачи Канбан-доски, укажите сроки и исполнителей или подготовьте короткую презентацию продукта.""",
    ),
]

QUESTIONS = [
    (
        "office-docs",
        "Для чего в документе используют стили заголовков?",
        ["Чтобы единообразно оформлять разделы и собирать оглавление", "Чтобы удалить текст", "Чтобы заменить документ таблицей"],
        0,
        "single",
        [0],
    ),
    (
        "office-docs",
        "Выберите инструменты, помогающие подготовить технический документ.",
        ["Стили заголовков", "Автоматическое оглавление", "Track Changes", "Сортировка строк в таблице"],
        0,
        "multiple",
        [0, 1, 2],
    ),
    (
        "office-docs",
        "Как называется документ, который фиксирует требования к разрабатываемой системе?",
        [],
        0,
        "text",
        ["ТЗ", "техническое задание"],
    ),
    (
        "office-spreadsheets",
        "Какая ссылка не меняется при копировании формулы?",
        ["Абсолютная", "Относительная", "Текстовая"],
        0,
        "single",
        [0],
    ),
    (
        "office-spreadsheets",
        "Какие задачи решают инструменты Excel?",
        ["Группировка записей сводной таблицей", "Выделение значений условным форматированием", "Оптимизация с ограничениями через «Поиск решения»", "Автоматическая настройка прав пользователей базы"],
        0,
        "multiple",
        [0, 1, 2],
    ),
    (
        "office-spreadsheets",
        "Как называется календарный план проекта в виде полос по срокам?",
        [],
        0,
        "text",
        ["диаграмма Ганта", "гант"],
    ),
    (
        "office-databases",
        "Что делает первичный ключ?",
        ["Однозначно определяет запись таблицы", "Раскрашивает ячейки", "Создаёт резервную копию файла"],
        0,
        "single",
        [0],
    ),
    (
        "office-databases",
        "Выберите верные утверждения о реляционной базе данных.",
        ["Данные хранятся в таблицах", "Внешний ключ связывает таблицы", "Нормализация помогает уменьшить дублирование", "Каждая запись обязательно хранится в отдельном файле"],
        0,
        "multiple",
        [0, 1, 2],
    ),
    (
        "office-databases",
        "Как называется язык запросов для работы с реляционными базами данных?",
        [],
        0,
        "text",
        ["SQL"],
    ),
    (
        "office-collaboration",
        "Что показывает Kanban-доска?",
        ["Состояние и движение задач по этапам", "Только список сотрудников", "Скорость сети"],
        0,
        "single",
        [0],
    ),
    (
        "office-collaboration",
        "Какие практики помогают команде совместно работать над проектом?",
        ["Назначать ответственных и сроки", "Разделять работу на небольшие задачи", "Выдавать доступ только по необходимости", "Публиковать пароль в общей доске"],
        0,
        "multiple",
        [0, 1, 2],
    ),
    (
        "office-collaboration",
        "Как называется упорядоченный список будущих задач продукта?",
        [],
        0,
        "text",
        ["бэклог", "backlog"],
    ),
]

PRACTICE_TASKS = [
    (
        "practice-1",
        "Практическое занятие №1 — Техническое задание",
        "По индивидуальному варианту подготовьте ТЗ на разработку ПО: оформите структуру, заголовки и автоматическое оглавление.",
        ["docx"],
    ),
    (
        "practice-2",
        "Практическое занятие №2 — Руководство пользователя",
        "Создайте руководство для программного продукта из своего ТЗ по практической работе №1: добавьте пошаговую инструкцию, схему алгоритма и аннотированные снимки экрана. Отдельный вариант не требуется.",
        ["docx"],
    ),
    (
        "practice-3",
        "Практическое занятие №3 — Сравнение редакций договора",
        "Сравните редакции договора на разработку ПО, подготовьте протокол разногласий и сохраните документ Word.",
        ["docx"],
    ),
    (
        "practice-5",
        "Практическое занятие №5 — Бюджет проекта",
        "Рассчитайте стоимость нормо-часа и бюджет проекта, используя абсолютные и относительные ссылки Excel.",
        ["xlsx"],
    ),
    (
        "practice-6",
        "Практическое занятие №6 — Очистка лога ошибок",
        "Импортируйте и очистите сырой журнал ошибок, разделите данные по столбцам и выделите критические сбои.",
        ["xlsx"],
    ),
    (
        "practice-7",
        "Практическое занятие №7 — Диаграмма Ганта",
        "Подготовьте интерактивный план этапов жизненного цикла разработки ПО в виде диаграммы Ганта.",
        ["xlsx"],
    ),
    (
        "practice-8",
        "Практическое занятие №8 — Анализ баг-трекера",
        "Создайте сводную таблицу по выгрузке баг-трекера и проанализируйте закрытые задачи и скорость исправления.",
        ["xlsx"],
    ),
    (
        "practice-9",
        "Практическое занятие №9 — Распределение задач",
        "Распределите задачи между Junior, Middle и Senior с помощью надстройки Поиск решения.",
        ["xlsx"],
    ),
]


class Database:
    def __init__(self) -> None:
        self.config = configparser.ConfigParser()
        if not CONFIG_PATH.exists():
            self.config["postgresql"] = {
                "host": "localhost",
                "port": "5432",
                "dbname": "electronic_textbook",
                "user": "postgres",
                "password": "1234",
            }
            self.config["paths"] = {"student_work_dir": "student_work"}
        self.config.read(CONFIG_PATH, encoding="utf-8")
        changed = False
        if "admin" not in self.config:
            self.config["admin"] = {"username": "admin", "password": "admin1"}
            changed = True
        else:
            for key, value in (("username", "admin"), ("password", "admin1")):
                if key not in self.config["admin"]:
                    self.config["admin"][key] = value
                    changed = True
            if not self.config["admin"]["password"].strip():
                self.config["admin"]["password"] = "admin1"
                changed = True
        if "paths" not in self.config:
            self.config["paths"] = {"student_work_dir": "student_work"}
            changed = True
        if changed or not CONFIG_PATH.exists():
            with CONFIG_PATH.open("w", encoding="utf-8") as config_file:
                self.config.write(config_file)

    @contextmanager
    def connect(self) -> Iterator:
        try:
            import psycopg
            from psycopg.rows import dict_row
        except ImportError as exc:
            raise RuntimeError(
                "Не установлен драйвер PostgreSQL. Выполните: "
                "python -m pip install -r requirements.txt"
            ) from exc

        database_settings = dict(self.config["postgresql"])
        database_settings["port"] = int(database_settings["port"])
        with psycopg.connect(**database_settings, row_factory=dict_row) as connection:
            yield connection

    def initialize(self) -> None:
        schema = SCHEMA_PATH.read_text(encoding="utf-8")
        with self.connect() as connection:
            connection.execute(schema)
            for slug, title, body in LESSONS:
                connection.execute(
                    """
                    INSERT INTO lessons (slug, title, body)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (slug) DO UPDATE
                    SET title = EXCLUDED.title,
                        body = EXCLUDED.body
                    """,
                    (slug, title, body),
                )
            for lesson_slug, prompt, options, correct_index, question_type, correct_answers in QUESTIONS:
                connection.execute(
                    """
                    INSERT INTO questions
                        (lesson_slug, prompt, options, correct_index, question_type, correct_answers)
                    VALUES (%s, %s, %s::jsonb, %s, %s, %s::jsonb)
                    ON CONFLICT (lesson_slug, prompt) DO NOTHING
                    """,
                    (
                        lesson_slug,
                        prompt,
                        json.dumps(options, ensure_ascii=False),
                        correct_index,
                        question_type,
                        json.dumps(correct_answers, ensure_ascii=False),
                    ),
                )
            for slug, title, instructions, extensions in PRACTICE_TASKS:
                connection.execute(
                    """
                    INSERT INTO practice_tasks (slug, title, instructions, extensions)
                    VALUES (%s, %s, %s, %s)
                    ON CONFLICT (slug) DO UPDATE
                    SET title = EXCLUDED.title,
                        instructions = EXCLUDED.instructions,
                        extensions = EXCLUDED.extensions
                    """,
                    (slug, title, instructions, extensions),
                )
            admin_exists = connection.execute(
                "SELECT EXISTS (SELECT 1 FROM students WHERE role = 'admin') AS found"
            ).fetchone()["found"]
            if not admin_exists:
                admin_username = self.config.get("admin", "username").strip()
                admin_password = self.config.get("admin", "password")
                if not admin_username or len(admin_password) < 4:
                    raise RuntimeError(
                        "Для первого запуска задайте логин администратора и пароль "
                        "(не менее 4 символов) в разделе [admin] файла config.ini."
                    )
                connection.execute(
                    """
                    INSERT INTO students
                        (full_name, student_group, username, password_hash, role)
                    VALUES ('Администратор', 'Администраторы', %s, %s, 'admin')
                    ON CONFLICT (username) DO NOTHING
                    """,
                    (admin_username, hash_password(admin_password)),
                )
                admin_created = connection.execute(
                    "SELECT EXISTS (SELECT 1 FROM students WHERE username = %s AND role = 'admin') AS found",
                    (admin_username,),
                ).fetchone()["found"]
                if not admin_created:
                    raise RuntimeError(
                        "Логин администратора уже занят обычным студентом. "
                        "Измените admin.username в config.ini."
                    )

    def authenticate(self, username: str) -> dict | None:
        with self.connect() as connection:
            return connection.execute(
                "SELECT id, full_name, student_group, username, password_hash, role "
                "FROM students WHERE username = %s",
                (username,),
            ).fetchone()

    def register(self, full_name: str, student_group: str, username: str, password_hash: str) -> dict:
        with self.connect() as connection:
            return connection.execute(
                """
                INSERT INTO students (full_name, student_group, username, password_hash)
                VALUES (%s, %s, %s, %s)
                RETURNING id, full_name, student_group, username, role
                """,
                (full_name, student_group, username, password_hash),
            ).fetchone()

    def get_lessons(self) -> list[dict]:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT slug, title, body FROM lessons
                WHERE slug NOT IN ('information', 'computer', 'software')
                ORDER BY id
                """
            ).fetchall()

    def get_questions(self, lesson_slug: str) -> list[dict]:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT id, lesson_slug, prompt, options, correct_index,
                       question_type, correct_answers
                FROM questions WHERE lesson_slug = %s ORDER BY id
                """,
                (lesson_slug,),
            ).fetchall()

    def save_attempt(self, student_id: int, lesson_slug: str, correct: int, total: int, percentage: int, grade: int) -> None:
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO test_attempts
                    (student_id, lesson_slug, correct_count, question_count, percentage, grade)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (student_id, lesson_slug, correct, total, percentage, grade),
            )

    def get_practice_tasks(self) -> list[dict]:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT slug, title, instructions, extensions FROM practice_tasks
                WHERE slug LIKE 'practice-%'
                ORDER BY CAST(SUBSTRING(slug FROM 10) AS INTEGER)
                """
            ).fetchall()

    def save_submission(
        self,
        student_id: int,
        task_slug: str,
        file_name: str,
        file_data: bytes,
        variant_number: int | None,
    ) -> int:
        with self.connect() as connection:
            result = connection.execute(
                """
                INSERT INTO practice_submissions
                    (student_id, task_slug, relative_path, file_name, file_data,
                     status, variant_number)
                VALUES (%s, %s, NULL, %s, %s, 'criteria_pending', %s)
                RETURNING id
                """,
                (student_id, task_slug, file_name, file_data, variant_number),
            ).fetchone()
            return result["id"]

    def get_or_assign_practice_variant(self, student_id: int, task_slug: str) -> int:
        with self.connect() as connection:
            existing = connection.execute(
                """
                SELECT variant_number FROM practice_variant_assignments
                WHERE student_id = %s AND task_slug = %s
                """,
                (student_id, task_slug),
            ).fetchone()
            if existing:
                return existing["variant_number"]

            for variant_number in range(1, 31):
                assigned = connection.execute(
                    """
                    INSERT INTO practice_variant_assignments
                        (student_id, task_slug, variant_number)
                    VALUES (%s, %s, %s)
                    ON CONFLICT DO NOTHING
                    RETURNING variant_number
                    """,
                    (student_id, task_slug, variant_number),
                ).fetchone()
                if assigned:
                    return assigned["variant_number"]

                existing = connection.execute(
                    """
                    SELECT variant_number FROM practice_variant_assignments
                    WHERE student_id = %s AND task_slug = %s
                    """,
                    (student_id, task_slug),
                ).fetchone()
                if existing:
                    return existing["variant_number"]
            raise RuntimeError(
                "Все 30 вариантов этого задания уже распределены между студентами."
            )

    def get_practice_variant_options(self, student_id: int, task_slug: str) -> dict:
        with self.connect() as connection:
            current = connection.execute(
                """
                SELECT variant_number FROM practice_variant_assignments
                WHERE student_id = %s AND task_slug = %s
                """,
                (student_id, task_slug),
            ).fetchone()
        return {
            "current": current["variant_number"] if current else None,
            "available": list(range(1, 31)),
        }

    def set_practice_variant(
        self, student_id: int, task_slug: str, variant_number: int
    ) -> int:
        if not 1 <= variant_number <= 30:
            raise ValueError("Номер варианта должен быть от 1 до 30.")
        try:
            with self.connect() as connection:
                result = connection.execute(
                    """
                    INSERT INTO practice_variant_assignments
                        (student_id, task_slug, variant_number)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (student_id, task_slug) DO UPDATE
                    SET variant_number = EXCLUDED.variant_number
                    RETURNING variant_number
                    """,
                    (student_id, task_slug, variant_number),
                ).fetchone()
                return result["variant_number"]
        except Exception as exc:
            import psycopg

            if isinstance(exc, psycopg.errors.UniqueViolation):
                raise ValueError(
                    "Этот вариант уже назначен другому студенту. Выберите другой номер."
                ) from exc
            raise

    def create_question(self, lesson_slug: str, question: dict) -> int:
        correct = question["correct_answers"]
        correct_index = correct[0] if question["question_type"] != "text" else 0
        with self.connect() as connection:
            result = connection.execute(
                """
                INSERT INTO questions
                    (lesson_slug, prompt, options, correct_index, question_type, correct_answers)
                VALUES (%s, %s, %s::jsonb, %s, %s, %s::jsonb)
                RETURNING id
                """,
                (
                    lesson_slug,
                    question["prompt"],
                    json.dumps(question["options"], ensure_ascii=False),
                    correct_index,
                    question["question_type"],
                    json.dumps(correct, ensure_ascii=False),
                ),
            ).fetchone()
            return result["id"]

    def update_question(self, question_id: int, question: dict) -> None:
        correct = question["correct_answers"]
        correct_index = correct[0] if question["question_type"] != "text" else 0
        with self.connect() as connection:
            connection.execute(
                """
                UPDATE questions
                SET prompt = %s, options = %s::jsonb, correct_index = %s,
                    question_type = %s, correct_answers = %s::jsonb
                WHERE id = %s
                """,
                (
                    question["prompt"],
                    json.dumps(question["options"], ensure_ascii=False),
                    correct_index,
                    question["question_type"],
                    json.dumps(correct, ensure_ascii=False),
                    question_id,
                ),
            )

    def delete_question(self, question_id: int) -> None:
        with self.connect() as connection:
            connection.execute("DELETE FROM questions WHERE id = %s", (question_id,))

    def get_questions_for_teacher(self, lesson_slug: str) -> list[dict]:
        return self.get_questions(lesson_slug)

    def get_results(self, student_id: int) -> list[dict]:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT 'Тест' AS result_type, l.title AS title,
                       a.grade, a.percentage, a.created_at
                FROM test_attempts a
                JOIN lessons l ON l.slug = a.lesson_slug
                WHERE a.student_id = %s
                UNION ALL
                SELECT 'Практика' AS result_type, t.title AS title,
                       s.grade, NULL AS percentage, s.created_at
                FROM practice_submissions s
                JOIN practice_tasks t ON t.slug = s.task_slug
                WHERE s.student_id = %s
                ORDER BY created_at DESC
                """,
                (student_id, student_id),
            ).fetchall()

    def get_practice_submissions(self) -> list[dict]:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT s.id, s.student_id, s.relative_path, s.file_name, s.status,
                       s.grade, s.created_at,
                       s.variant_number, u.full_name, u.username, u.student_group,
                       t.title AS task_title
                FROM practice_submissions s
                JOIN students u ON u.id = s.student_id
                JOIN practice_tasks t ON t.slug = s.task_slug
                ORDER BY s.created_at DESC
                """
            ).fetchall()

    def get_submission_file(self, submission_id: int) -> dict | None:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT file_name, file_data
                FROM practice_submissions
                WHERE id = %s
                """,
                (submission_id,),
            ).fetchone()

    def get_test_results_for_teacher(self) -> list[dict]:
        with self.connect() as connection:
            return connection.execute(
                """
                SELECT a.id, u.full_name, u.username, u.student_group, l.title AS lesson_title,
                       a.correct_count, a.question_count, a.percentage, a.grade,
                       a.created_at
                FROM test_attempts a
                JOIN students u ON u.id = a.student_id
                JOIN lessons l ON l.slug = a.lesson_slug
                ORDER BY a.created_at DESC, a.id DESC
                """
            ).fetchall()

    def grade_submission(self, submission_id: int, grade: int) -> None:
        with self.connect() as connection:
            result = connection.execute(
                """
                UPDATE practice_submissions
                SET grade = %s, status = 'graded'
                WHERE id = %s
                """,
                (grade, submission_id),
            )
            if result.rowcount != 1:
                raise ValueError("Работа не найдена или уже удалена.")

    def work_root(self) -> Path:
        configured = self.config.get("paths", "student_work_dir", fallback="student_work")
        base = Path(configured)
        if not base.is_absolute():
            base = APP_DIR / base
        return base

    def work_dir(self, student_id: int) -> Path:
        return self.work_root() / str(student_id)
