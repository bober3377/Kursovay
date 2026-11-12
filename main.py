import os
import random
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText
from uuid import uuid4

from database import Database
from logic import (
    MAX_UPLOAD_SIZE,
    calculate_grade,
    check_question_answer,
    hash_password,
    shuffle_question,
    validate_practice_file,
    validate_question,
    verify_password,
)
from practice_variants import (
    PRACTICES_WITH_VARIANTS,
    build_practice_csv,
    build_practice_file,
)


class TextbookApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Электронный учебник по информационным технологиям")
        self.root.geometry("1160x760")
        self.root.minsize(900, 620)
        self.root.configure(bg="#f1f5f9")
        self.db = Database()
        self.student: dict | None = None

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background="#f1f5f9")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("TLabel", background="#f1f5f9", foreground="#243449")
        style.configure("Card.TLabel", background="#ffffff", foreground="#243449")
        style.configure("Title.TLabel", background="#ffffff", foreground="#17365d",
                        font=("Segoe UI", 22, "bold"))
        style.configure("Section.TLabel", foreground="#17365d", font=("Segoe UI", 14, "bold"))
        style.configure("Header.TFrame", background="#17365d")
        style.configure("Sidebar.TFrame", background="#10243b")
        style.configure("SidebarBrand.TLabel", background="#10243b", foreground="#ffffff",
                        font=("Segoe UI", 14, "bold"))
        style.configure("SidebarCaption.TLabel", background="#10243b", foreground="#9eb3ca",
                        font=("Segoe UI", 9))
        style.configure("SidebarUser.TLabel", background="#10243b", foreground="#dbeafe")
        style.configure("HeaderTitle.TLabel", background="#17365d", foreground="#ffffff",
                        font=("Segoe UI", 16, "bold"))
        style.configure("HeaderMeta.TLabel", background="#17365d", foreground="#dbeafe")
        style.configure("TButton", padding=(12, 8), background="#e2eaf4",
                        foreground="#17365d", borderwidth=0)
        style.map("TButton", background=[("active", "#cbd9e9")])
        style.configure("Primary.TButton", padding=(12, 9), background="#087e8b",
                        foreground="#ffffff", borderwidth=0, font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", "#066b76")])
        style.configure("Nav.TButton", padding=(14, 12), background="#10243b",
                        foreground="#d4e0ed", borderwidth=0, anchor="w")
        style.map("Nav.TButton", background=[("active", "#1b3856")],
                  foreground=[("active", "#ffffff")])
        style.configure("NavActive.TButton", padding=(14, 12), background="#087e8b",
                        foreground="#ffffff", borderwidth=0, anchor="w",
                        font=("Segoe UI", 10, "bold"))
        style.map("NavActive.TButton", background=[("active", "#066b76")])
        style.configure("TEntry", padding=9, fieldbackground="#ffffff")
        style.configure("TNotebook", background="#f1f5f9", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 11), background="#e2eaf4",
                        foreground="#42556d", borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", "#ffffff")],
                  foreground=[("selected", "#087e8b")])
        style.configure("TLabelframe", background="#ffffff", bordercolor="#dce5ef")
        style.configure("TLabelframe.Label", background="#ffffff", foreground="#17365d",
                        font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", rowheight=34, background="#ffffff",
                        fieldbackground="#ffffff", foreground="#243449")
        style.configure("Treeview.Heading", background="#e7eef6", foreground="#17365d",
                        font=("Segoe UI", 10, "bold"), padding=(8, 9))

        try:
            self.db.initialize()
        except Exception as exc:
            messagebox.showerror(
                "Не удалось запустить приложение",
                f"{exc}\n\nПроверьте PostgreSQL и настройки в config.ini. "
                "Инструкция находится в README.md.",
                parent=root,
            )
            root.destroy()
            return

        self.show_login()

    def clear(self) -> None:
        for widget in self.root.winfo_children():
            widget.destroy()

    def show_login(self) -> None:
        self.clear()
        card = ttk.Frame(self.root, style="Card.TFrame", padding=(38, 34))
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="ЭЛЕКТРОННЫЙ УЧЕБНИК", style="Title.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        ttk.Label(card, text="Информационные технологии · учись и проверяй знания",
                  style="Card.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 26)
        )
        ttk.Label(card, text="Логин", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=7)
        username = ttk.Entry(card, width=38)
        username.grid(row=2, column=1, pady=7, padx=(18, 0))
        ttk.Label(card, text="Пароль", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=7)
        password = ttk.Entry(card, width=38, show="*")
        password.grid(row=3, column=1, pady=7, padx=(18, 0))
        ttk.Button(
            card, text="Войти", style="Primary.TButton",
            command=lambda: self.login(username.get(), password.get())
        ).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(22, 9))
        ttk.Button(card, text="Зарегистрироваться", command=self.show_registration).grid(
            row=5, column=0, columnspan=2, sticky="ew"
        )
        password.bind("<Return>", lambda _event: self.login(username.get(), password.get()))
        username.focus_set()

    def show_registration(self) -> None:
        self.clear()
        card = ttk.Frame(self.root, style="Card.TFrame", padding=(38, 32))
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Регистрация студента", style="Title.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 18)
        )
        fields = [
            ("ФИО", "name"),
            ("Группа", "group"),
            ("Логин", "username"),
            ("Пароль (не менее 8 символов)", "password"),
        ]
        entries: dict[str, ttk.Entry] = {}
        for row, (label, key) in enumerate(fields, start=1):
            ttk.Label(card, text=label, style="Card.TLabel").grid(row=row, column=0, sticky="w", pady=7)
            entry = ttk.Entry(card, width=38, show="*" if key == "password" else "")
            entry.grid(row=row, column=1, pady=7, padx=(18, 0))
            entries[key] = entry

        ttk.Button(
            card,
            text="Создать аккаунт",
            style="Primary.TButton",
            command=lambda: self.register(
                entries["name"].get(),
                entries["group"].get(),
                entries["username"].get(),
                entries["password"].get(),
            ),
        ).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(16, 8))
        ttk.Button(card, text="Назад ко входу", command=self.show_login).grid(
            row=6, column=0, columnspan=2, sticky="ew"
        )

    def register(self, full_name: str, student_group: str, username: str, password: str) -> None:
        full_name, student_group, username = (
            full_name.strip(), student_group.strip(), username.strip()
        )
        if not all((full_name, student_group, username, password)):
            messagebox.showwarning("Проверьте данные", "Заполните все поля.", parent=self.root)
            return
        if len(username) < 3 or any(character.isspace() for character in username):
            messagebox.showwarning(
                "Проверьте данные", "Логин должен быть не короче 3 символов и без пробелов.",
                parent=self.root,
            )
            return
        if len(password) < 8:
            messagebox.showwarning("Проверьте данные", "Пароль должен содержать не менее 8 символов.", parent=self.root)
            return

        try:
            self.student = self.db.register(
                full_name, student_group, username, hash_password(password)
            )
        except Exception as exc:
            if getattr(exc, "sqlstate", None) == "23505":
                messagebox.showerror("Регистрация", "Такой логин уже занят.", parent=self.root)
            else:
                messagebox.showerror("Ошибка регистрации", str(exc), parent=self.root)
            return
        self.show_dashboard()

    def login(self, username: str, password: str) -> None:
        try:
            student = self.db.authenticate(username.strip())
        except Exception as exc:
            messagebox.showerror("Ошибка входа", str(exc), parent=self.root)
            return
        if not student or not verify_password(password, student["password_hash"]):
            messagebox.showerror("Ошибка входа", "Неверный логин или пароль.", parent=self.root)
            return
        self.student = student
        if student["role"] == "admin":
            self.show_admin_dashboard()
            return
        self.show_dashboard()

    def show_dashboard(self) -> None:
        if self.student is None:
            self.show_login()
            return
        self.clear()
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(26, 16))
        header.pack(fill="x")
        identity = ttk.Frame(header, style="Header.TFrame")
        identity.pack(side="left")
        ttk.Label(identity, text="ЭЛЕКТРОННЫЙ УЧЕБНИК", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(identity, text="Информационные технологии", style="HeaderMeta.TLabel").pack(
            anchor="w", pady=(2, 0)
        )
        ttk.Button(header, text="Выйти", command=self.logout).pack(side="right")
        ttk.Label(
            header,
            text=f"{self.student['full_name']}  ·  {self.student['student_group']}",
            style="HeaderMeta.TLabel",
        ).pack(side="right", padx=(0, 18))

        shell = ttk.Frame(self.root)
        shell.pack(fill="both", expand=True)
        sidebar = ttk.Frame(shell, style="Sidebar.TFrame", width=230, padding=(14, 22))
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        ttk.Label(sidebar, text="МОЁ ОБУЧЕНИЕ", style="SidebarCaption.TLabel").pack(
            anchor="w", padx=12, pady=(4, 16)
        )
        self.nav_buttons: dict[str, ttk.Button] = {}
        navigation = [
            ("materials", "01   Учебные материалы"),
            ("test", "02   Проверка знаний"),
            ("practice", "03   Практические работы"),
            ("results", "04   Мои результаты"),
        ]
        for page, label in navigation:
            button = ttk.Button(
                sidebar,
                text=label,
                style="Nav.TButton",
                command=lambda selected=page: self.switch_student_page(selected),
            )
            button.pack(fill="x", pady=3)
            self.nav_buttons[page] = button

        ttk.Frame(sidebar, style="Sidebar.TFrame").pack(fill="both", expand=True)
        ttk.Separator(sidebar).pack(fill="x", padx=8, pady=(14, 12))
        ttk.Label(sidebar, text="ЛИЧНЫЙ КАБИНЕТ", style="SidebarCaption.TLabel").pack(
            anchor="w", padx=12
        )
        ttk.Label(
            sidebar,
            text=f"{self.student['full_name']}\nГруппа {self.student['student_group']}",
            style="SidebarUser.TLabel",
            justify="left",
            wraplength=190,
        ).pack(anchor="w", padx=12, pady=(7, 8))

        self.content_area = ttk.Frame(shell, padding=(26, 24))
        self.content_area.pack(side="left", fill="both", expand=True)
        self.switch_student_page("materials")

    def switch_student_page(self, page: str) -> None:
        if not hasattr(self, "content_area"):
            return
        for widget in self.content_area.winfo_children():
            widget.destroy()
        for key, button in self.nav_buttons.items():
            button.configure(style="NavActive.TButton" if key == page else "Nav.TButton")
        builders = {
            "materials": self.build_lessons_tab,
            "test": self.build_test_tab,
            "practice": self.build_practice_tab,
            "results": self.build_results_tab,
        }
        builders[page]()

    def show_admin_dashboard(self) -> None:
        if self.student is None:
            self.show_login()
            return
        self.clear()
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(24, 16))
        header.pack(fill="x")
        ttk.Label(header, text="Панель преподавателя", style="HeaderTitle.TLabel").pack(side="left")
        ttk.Label(
            header, text=f"Администратор: {self.student['username']}",
            style="HeaderMeta.TLabel",
        ).pack(side="right", padx=(12, 0))
        ttk.Button(header, text="Выйти", command=self.logout).pack(side="right")

        self.admin_tabs = ttk.Notebook(self.root)
        self.admin_tabs.pack(fill="both", expand=True, padx=22, pady=20)

        practice_page = ttk.Frame(self.admin_tabs, padding=18)
        self.admin_tabs.add(practice_page, text="Практические работы")
        toolbar = ttk.Frame(practice_page)
        toolbar.pack(fill="x", pady=(0, 14))
        ttk.Label(toolbar, text="Работы студентов", style="Section.TLabel").pack(side="left")
        ttk.Button(
            toolbar, text="Обновить список", command=self.refresh_submissions
        ).pack(side="right")
        self.submission_search, self.submission_group, self.submission_group_selector = (
            self.build_admin_student_filters(practice_page, self.refresh_submissions)
        )

        card = ttk.Frame(practice_page, style="Card.TFrame", padding=12)
        card.pack(fill="both", expand=True)
        columns = ("student", "group", "task", "variant", "file", "grade", "status", "created")
        self.submissions_table = ttk.Treeview(card, columns=columns, show="headings")
        headings = {
            "student": ("Студент", 190),
            "group": ("Группа", 100),
            "task": ("Задание", 210),
            "variant": ("Вариант", 75),
            "file": ("Файл", 220),
            "grade": ("Оценка", 85),
            "status": ("Статус", 120),
            "created": ("Дата отправки", 145),
        }
        for column, (title, width) in headings.items():
            self.submissions_table.heading(column, text=title)
            self.submissions_table.column(column, width=width, anchor="w")
        self.submissions_table.pack(fill="both", expand=True)
        self.submission_paths: dict[str, dict] = {}

        actions = ttk.Frame(practice_page)
        actions.pack(fill="x", pady=(14, 0))
        ttk.Button(
            actions, text="Открыть работу", command=self.open_selected_submission
        ).pack(side="left")
        ttk.Button(
            actions, text="Выставить / изменить оценку", style="Primary.TButton",
            command=self.grade_selected_submission,
        ).pack(side="left", padx=10)
        ttk.Label(
            practice_page,
            text="Выберите работу в списке. Оценка за практику выставляется вручную преподавателем.",
        ).pack(anchor="w", pady=(12, 0))

        tests_page = ttk.Frame(self.admin_tabs, padding=18)
        self.admin_tabs.add(tests_page, text="Результаты тестов")
        test_toolbar = ttk.Frame(tests_page)
        test_toolbar.pack(fill="x", pady=(0, 14))
        ttk.Label(test_toolbar, text="Оценки студентов за тесты", style="Section.TLabel").pack(
            side="left"
        )
        ttk.Button(
            test_toolbar, text="Обновить список", command=self.refresh_test_results
        ).pack(side="right")
        self.test_search, self.test_group, self.test_group_selector = (
            self.build_admin_student_filters(tests_page, self.refresh_test_results)
        )
        ttk.Label(
            tests_page,
            text="Здесь отображаются результаты каждой попытки: правильные ответы, процент и оценка.",
            foreground="#687b90",
        ).pack(anchor="w", pady=(0, 12))
        test_card = ttk.Frame(tests_page, style="Card.TFrame", padding=12)
        test_card.pack(fill="both", expand=True)
        test_columns = ("student", "group", "lesson", "correct", "percentage", "grade", "created")
        self.test_results_table = ttk.Treeview(
            test_card, columns=test_columns, show="headings"
        )
        test_headings = {
            "student": ("Студент", 210),
            "group": ("Группа", 100),
            "lesson": ("Раздел", 250),
            "correct": ("Верно", 95),
            "percentage": ("Результат", 105),
            "grade": ("Оценка", 85),
            "created": ("Дата", 155),
        }
        for column, (title, width) in test_headings.items():
            self.test_results_table.heading(column, text=title)
            self.test_results_table.column(column, width=width, anchor="w")
        self.test_results_table.pack(fill="both", expand=True)

        constructor_page = ttk.Frame(self.admin_tabs, padding=18)
        self.admin_tabs.add(constructor_page, text="Конструктор тестов")
        self.build_test_constructor(constructor_page)

        self.refresh_submissions()
        self.refresh_test_results()

    def build_admin_student_filters(
        self, parent: ttk.Frame, refresh_callback
    ) -> tuple[tk.StringVar, tk.StringVar, ttk.Combobox]:
        filters = ttk.Frame(parent)
        filters.pack(fill="x", pady=(0, 12))
        ttk.Label(filters, text="Поиск студента:").pack(side="left")
        search = tk.StringVar()
        search_entry = ttk.Entry(filters, textvariable=search, width=30)
        search_entry.pack(side="left", padx=(8, 18))
        ttk.Label(filters, text="Группа:").pack(side="left")
        group = tk.StringVar(value="Все группы")
        selector = ttk.Combobox(
            filters,
            textvariable=group,
            values=["Все группы"],
            state="readonly",
            width=24,
        )
        selector.pack(side="left", padx=(8, 12))
        search_entry.bind("<KeyRelease>", lambda _event: refresh_callback())
        selector.bind("<<ComboboxSelected>>", lambda _event: refresh_callback())

        def reset_filters() -> None:
            search.set("")
            group.set("Все группы")
            refresh_callback()

        ttk.Button(filters, text="Сбросить", command=reset_filters).pack(side="left")
        return search, group, selector

    @staticmethod
    def matches_admin_student(
        student: dict, search_text: str, selected_group: str
    ) -> bool:
        normalized_search = search_text.strip().casefold()
        matches_search = not normalized_search or any(
            normalized_search in str(student.get(field, "")).casefold()
            for field in ("full_name", "username")
        )
        matches_group = (
            selected_group == "Все группы"
            or student["student_group"] == selected_group
        )
        return matches_search and matches_group

    @staticmethod
    def update_admin_group_options(
        selector: ttk.Combobox, group_variable: tk.StringVar, records: list[dict]
    ) -> None:
        groups = sorted({record["student_group"] for record in records})
        values = ["Все группы", *groups]
        selector.configure(values=values)
        if group_variable.get() not in values:
            group_variable.set("Все группы")

    def refresh_submissions(self) -> None:
        if not hasattr(self, "submissions_table"):
            return
        try:
            self.all_submissions = self.db.get_practice_submissions()
        except Exception as exc:
            messagebox.showerror("Ошибка загрузки работ", str(exc), parent=self.root)
            return
        self.update_admin_group_options(
            self.submission_group_selector,
            self.submission_group,
            self.all_submissions,
        )
        self.submissions_table.delete(*self.submissions_table.get_children())
        self.submission_paths.clear()
        for submission in self.all_submissions:
            if not self.matches_admin_student(
                submission, self.submission_search.get(), self.submission_group.get()
            ):
                continue
            item_id = str(submission["id"])
            self.submission_paths[item_id] = submission
            grade = submission["grade"] if submission["grade"] is not None else "—"
            status = "Оценена" if submission["grade"] is not None else "Не оценена"
            created = submission["created_at"].astimezone().strftime("%d.%m.%Y %H:%M")
            self.submissions_table.insert(
                "", "end", iid=item_id,
                values=(
                    submission["full_name"], submission["student_group"],
                    submission["task_title"],
                    submission["variant_number"] or "—",
                    submission["file_name"]
                    or Path(submission["relative_path"] or "Файл").name,
                    grade, status, created,
                ),
            )

    def refresh_test_results(self) -> None:
        if (
            not hasattr(self, "test_results_table")
            or not self.test_results_table.winfo_exists()
        ):
            return
        try:
            self.all_test_results = self.db.get_test_results_for_teacher()
        except Exception as exc:
            messagebox.showerror("Ошибка загрузки оценок за тесты", str(exc), parent=self.root)
            return
        self.update_admin_group_options(
            self.test_group_selector,
            self.test_group,
            self.all_test_results,
        )
        self.test_results_table.delete(*self.test_results_table.get_children())
        for result in self.all_test_results:
            if not self.matches_admin_student(
                result, self.test_search.get(), self.test_group.get()
            ):
                continue
            created = result["created_at"].astimezone().strftime("%d.%m.%Y %H:%M")
            self.test_results_table.insert(
                "",
                "end",
                values=(
                    result["full_name"],
                    result["student_group"],
                    result["lesson_title"],
                    f"{result['correct_count']} из {result['question_count']}",
                    f"{result['percentage']}%",
                    result["grade"],
                    created,
                ),
            )

    def build_test_constructor(self, page: ttk.Frame) -> None:
        toolbar = ttk.Frame(page)
        toolbar.pack(fill="x", pady=(0, 14))
        ttk.Label(toolbar, text="Вопросы по разделам", style="Section.TLabel").pack(side="left")
        ttk.Label(toolbar, text="Раздел курса", padding=(18, 0, 8, 0)).pack(side="left")
        self.constructor_lesson = tk.StringVar()
        self.constructor_lesson_selector = ttk.Combobox(
            toolbar,
            textvariable=self.constructor_lesson,
            values=[lesson["title"] for lesson in self.db.get_lessons()],
            state="readonly",
            width=44,
        )
        self.constructor_lesson_selector.pack(side="left")
        self.constructor_lesson_selector.bind(
            "<<ComboboxSelected>>", lambda _event: self.refresh_question_list()
        )

        table_card = ttk.Frame(page, style="Card.TFrame", padding=12)
        table_card.pack(fill="both", expand=True)
        self.question_table = ttk.Treeview(
            table_card, columns=("type", "prompt", "answer"), show="headings"
        )
        for column, title, width in (
            ("type", "Тип", 140),
            ("prompt", "Текст вопроса", 530),
            ("answer", "Правильный ответ", 260),
        ):
            self.question_table.heading(column, text=title)
            self.question_table.column(column, width=width, anchor="w")
        self.question_table.pack(fill="both", expand=True)
        actions = ttk.Frame(page)
        actions.pack(fill="x", pady=(14, 0))
        ttk.Button(
            actions,
            text="Добавить вопрос",
            style="Primary.TButton",
            command=lambda: self.open_question_editor(),
        ).pack(side="left")
        ttk.Button(
            actions, text="Изменить", command=self.edit_selected_question
        ).pack(side="left", padx=8)
        ttk.Button(
            actions, text="Удалить", command=self.delete_selected_question
        ).pack(side="left")
        ttk.Button(
            actions, text="Обновить", command=self.refresh_question_list
        ).pack(side="right")
        self.question_records: dict[str, dict] = {}
        lessons = self.db.get_lessons()
        if lessons:
            self.constructor_lesson_selector.current(0)
            self.refresh_question_list()

    def selected_constructor_lesson(self) -> dict | None:
        index = self.constructor_lesson_selector.current()
        lessons = self.db.get_lessons()
        if index < 0 or index >= len(lessons):
            return None
        return lessons[index]

    def refresh_question_list(self) -> None:
        if not hasattr(self, "question_table") or not self.question_table.winfo_exists():
            return
        lesson = self.selected_constructor_lesson()
        if lesson is None:
            self.question_table.delete(*self.question_table.get_children())
            self.question_records = {}
            return
        try:
            questions = self.db.get_questions_for_teacher(lesson["slug"])
        except Exception as exc:
            messagebox.showerror("Не удалось загрузить вопросы", str(exc), parent=self.root)
            return
        self.question_table.delete(*self.question_table.get_children())
        self.question_records = {}
        type_names = {"single": "Один ответ", "multiple": "Несколько ответов", "text": "Ввод текста"}
        for question in questions:
            item_id = str(question["id"])
            self.question_records[item_id] = question
            correct_answers = question.get("correct_answers") or [question["correct_index"]]
            if question["question_type"] == "text":
                answer_text = "; ".join(correct_answers)
            else:
                answer_text = "; ".join(
                    question["options"][index]
                    for index in correct_answers
                    if 0 <= index < len(question["options"])
                )
            self.question_table.insert(
                "", "end", iid=item_id,
                values=(
                    type_names.get(question["question_type"], question["question_type"]),
                    question["prompt"],
                    answer_text,
                ),
            )

    def edit_selected_question(self) -> None:
        selection = self.question_table.selection()
        if not selection:
            messagebox.showinfo("Конструктор тестов", "Сначала выберите вопрос.", parent=self.root)
            return
        self.open_question_editor(self.question_records[selection[0]])

    def delete_selected_question(self) -> None:
        selection = self.question_table.selection()
        if not selection:
            messagebox.showinfo("Конструктор тестов", "Сначала выберите вопрос.", parent=self.root)
            return
        question = self.question_records[selection[0]]
        if not messagebox.askyesno(
            "Удалить вопрос",
            f"Удалить вопрос «{question['prompt']}»?",
            parent=self.root,
        ):
            return
        try:
            self.db.delete_question(question["id"])
        except Exception as exc:
            messagebox.showerror("Не удалось удалить вопрос", str(exc), parent=self.root)
            return
        self.refresh_question_list()

    def open_question_editor(self, question: dict | None = None) -> None:
        lesson = self.selected_constructor_lesson()
        if lesson is None:
            messagebox.showwarning(
                "Конструктор тестов", "Сначала добавьте раздел курса.", parent=self.root
            )
            return
        window = tk.Toplevel(self.root)
        window.title("Редактировать вопрос" if question else "Новый вопрос")
        window.geometry("720x640")
        window.minsize(620, 520)
        window.transient(self.root)
        window.grab_set()

        frame = ttk.Frame(window, padding=22)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="Редактирование вопроса" if question else "Новый вопрос",
            style="Section.TLabel",
        ).pack(anchor="w", pady=(0, 14))
        ttk.Label(frame, text="Формулировка вопроса").pack(anchor="w")
        prompt = ScrolledText(frame, height=4, wrap="word", font=("Segoe UI", 10))
        prompt.pack(fill="x", pady=(5, 14))
        if question:
            prompt.insert("1.0", question["prompt"])

        ttk.Label(frame, text="Тип вопроса").pack(anchor="w")
        type_names = {
            "Один правильный ответ": "single",
            "Несколько правильных ответов": "multiple",
            "Свободный текстовый ответ": "text",
        }
        name_by_type = {value: key for key, value in type_names.items()}
        selected_type = tk.StringVar(
            value=name_by_type.get(question["question_type"], "Один правильный ответ")
            if question
            else "Один правильный ответ"
        )
        type_selector = ttk.Combobox(
            frame,
            textvariable=selected_type,
            values=list(type_names),
            state="readonly",
        )
        type_selector.pack(fill="x", pady=(5, 12))

        dynamic = ttk.Frame(frame)
        dynamic.pack(fill="both", expand=True)
        dynamic_widgets: dict[str, tk.Widget] = {}

        def render_type(_event=None) -> None:
            for child in dynamic.winfo_children():
                child.destroy()
            dynamic_widgets.clear()
            question_type = type_names[selected_type.get()]
            if question_type in {"single", "multiple"}:
                ttk.Label(
                    dynamic,
                    text="Варианты ответа — по одному на строку. Выдели правильный вариант(ы) в списке.",
                    wraplength=660,
                ).pack(anchor="w", pady=(0, 6))
                ttk.Label(dynamic, text="Варианты ответов").pack(anchor="w")
                options_text = ScrolledText(
                    dynamic, height=7, wrap="word", font=("Segoe UI", 10)
                )
                options_text.pack(fill="x", pady=(4, 8))
                correct_label = (
                    "Выбери ровно один правильный ответ:"
                    if question_type == "single"
                    else "Выбери все правильные ответы:"
                )
                ttk.Label(dynamic, text=correct_label).pack(anchor="w")
                correct_list = tk.Listbox(
                    dynamic,
                    height=6,
                    exportselection=False,
                    selectmode="browse" if question_type == "single" else "extended",
                )
                correct_list.pack(fill="x", pady=(4, 0))

                def refresh_options(_event=None) -> None:
                    old_selection = set(correct_list.curselection())
                    values = [
                        line.strip()
                        for line in options_text.get("1.0", "end").splitlines()
                        if line.strip()
                    ]
                    correct_list.delete(0, "end")
                    for value in values:
                        correct_list.insert("end", value)
                    for index in old_selection:
                        if index < len(values):
                            correct_list.selection_set(index)

                options_text.bind("<KeyRelease>", refresh_options)
                if question and question["question_type"] in {"single", "multiple"}:
                    options_text.insert("1.0", "\n".join(question["options"]))
                    refresh_options()
                    for index in question.get("correct_answers", [question["correct_index"]]):
                        if index < correct_list.size():
                            correct_list.selection_set(index)
                dynamic_widgets.update(
                    options=options_text,
                    correct=correct_list,
                )
            else:
                ttk.Label(
                    dynamic,
                    text="Укажи правильные ответы по одному на строку. Ввод у студента "
                    "не зависит от регистра и пробелов по краям.",
                    wraplength=660,
                ).pack(anchor="w", pady=(0, 6))
                ttk.Label(dynamic, text="Допустимые ответы").pack(anchor="w")
                answers_text = ScrolledText(
                    dynamic, height=9, wrap="word", font=("Segoe UI", 10)
                )
                answers_text.pack(fill="both", expand=True, pady=(4, 0))
                if question and question["question_type"] == "text":
                    answers_text.insert(
                        "1.0", "\n".join(question.get("correct_answers", []))
                    )
                dynamic_widgets["answers"] = answers_text

        type_selector.bind("<<ComboboxSelected>>", render_type)
        render_type()

        def save() -> None:
            question_type = type_names[selected_type.get()]
            try:
                if question_type == "text":
                    answer_lines = [
                        value.strip()
                        for value in dynamic_widgets["answers"].get("1.0", "end").splitlines()
                        if value.strip()
                    ]
                    question_data = validate_question(
                        question_type,
                        prompt.get("1.0", "end"),
                        [],
                        answer_lines,
                    )
                else:
                    options = dynamic_widgets["options"].get("1.0", "end").splitlines()
                    correct = list(dynamic_widgets["correct"].curselection())
                    question_data = validate_question(
                        question_type,
                        prompt.get("1.0", "end"),
                        options,
                        correct,
                    )
                if question:
                    self.db.update_question(question["id"], question_data)
                else:
                    self.db.create_question(lesson["slug"], question_data)
            except Exception as exc:
                messagebox.showerror(
                    "Не удалось сохранить вопрос", str(exc), parent=window
                )
                return
            window.destroy()
            self.refresh_question_list()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="Отмена", command=window.destroy).pack(side="right")
        ttk.Button(
            buttons, text="Сохранить вопрос", style="Primary.TButton", command=save
        ).pack(side="right", padx=(0, 8))

    def selected_submission(self) -> dict | None:
        selection = self.submissions_table.selection()
        if not selection:
            messagebox.showinfo("Работы студентов", "Сначала выберите работу в списке.", parent=self.root)
            return None
        return self.submission_paths.get(selection[0])

    def open_selected_submission(self) -> None:
        submission = self.selected_submission()
        if submission is None:
            return
        try:
            stored_file = self.db.get_submission_file(submission["id"])
            if stored_file and stored_file["file_data"] is not None:
                file_name = Path(stored_file["file_name"]).name
                suffix = Path(file_name).suffix
                temporary_folder = Path(tempfile.gettempdir()) / "electronic_textbook"
                temporary_folder.mkdir(parents=True, exist_ok=True)
                path = temporary_folder / (
                    f"submission_{submission['id']}_{uuid4().hex}{suffix}"
                )
                path.write_bytes(bytes(stored_file["file_data"]))
            else:
                path = self.resolve_submission_path(submission)
            os.startfile(str(path))
        except Exception as exc:
            messagebox.showerror("Не удалось открыть работу", str(exc), parent=self.root)

    def resolve_submission_path(self, submission: dict) -> Path:
        student_id = submission["student_id"]
        student_folder = self.db.work_dir(student_id).resolve()
        stored_path = Path(submission["relative_path"])
        if stored_path.is_absolute():
            candidates = [stored_path]
        else:
            candidates = [self.db.work_root() / stored_path]
            if (
                len(stored_path.parts) >= 3
                and stored_path.parts[0].casefold() == "student_work"
                and stored_path.parts[1] == str(student_id)
            ):
                candidates.append(
                    self.db.work_root().joinpath(*stored_path.parts[1:])
                )
            candidates.append(Path(__file__).resolve().parent / stored_path)

        for candidate in candidates:
            resolved = candidate.resolve()
            try:
                resolved.relative_to(student_folder)
            except ValueError:
                continue
            if resolved.is_file():
                return resolved
        raise FileNotFoundError(
            "Файл работы не найден в общей папке студента. "
            "Проверьте сетевую папку и доступ к ней."
        )

    def grade_selected_submission(self) -> None:
        submission = self.selected_submission()
        if submission is None:
            return
        grade = simpledialog.askinteger(
            "Оценка за практику",
            f"Введите оценку для работы студента {submission['full_name']} (от 2 до 5):",
            parent=self.root,
            minvalue=2,
            maxvalue=5,
        )
        if grade is None:
            return
        try:
            self.db.grade_submission(submission["id"], grade)
        except Exception as exc:
            messagebox.showerror("Не удалось сохранить оценку", str(exc), parent=self.root)
            return
        self.refresh_submissions()
        messagebox.showinfo("Оценка сохранена", f"Работе выставлена оценка {grade}.", parent=self.root)

    def build_lessons_tab(self) -> None:
        try:
            lessons = self.db.get_lessons()
        except Exception as exc:
            messagebox.showerror("Ошибка загрузки материалов", str(exc), parent=self.root)
            return
        self.lessons = lessons
        ttk.Label(self.content_area, text="Учебные материалы", style="Section.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            self.content_area,
            text="Выбери раздел слева и изучай тему в удобном темпе.",
            foreground="#687b90",
        ).pack(anchor="w", pady=(4, 18))

        workspace = ttk.Frame(self.content_area)
        workspace.pack(fill="both", expand=True)
        catalog = ttk.Frame(workspace, style="Card.TFrame", padding=16)
        catalog.pack(side="left", fill="y", padx=(0, 18))
        ttk.Label(catalog, text="РАЗДЕЛЫ КУРСА", style="Card.TLabel",
                  font=("Segoe UI", 9, "bold"), foreground="#687b90").pack(anchor="w")
        ttk.Label(
            catalog, text=f"{len(lessons)} учебных тем", style="Card.TLabel",
            foreground="#8795a6",
        ).pack(anchor="w", pady=(3, 12))
        catalog_list = ttk.Frame(catalog, style="Card.TFrame")
        catalog_list.pack(fill="both", expand=True)
        lesson_scroll = ttk.Scrollbar(catalog_list, orient="vertical")
        lesson_scroll.pack(side="right", fill="y")
        self.lesson_list = tk.Listbox(
            catalog_list,
            width=33,
            height=18,
            exportselection=False,
            activestyle="none",
            relief="flat",
            bd=0,
            bg="#f1f5f9",
            fg="#34465d",
            selectbackground="#d5efed",
            selectforeground="#075b65",
            highlightthickness=0,
            font=("Segoe UI", 10),
            selectmode="browse",
            yscrollcommand=lesson_scroll.set,
        )
        self.lesson_list.pack(side="left", fill="both", expand=True)
        lesson_scroll.configure(command=self.lesson_list.yview)
        for lesson in lessons:
            self.lesson_list.insert("end", lesson["title"])
        self.lesson_list.bind("<<ListboxSelect>>", self.show_lesson)

        reader = ttk.Frame(workspace, style="Card.TFrame", padding=(24, 22))
        reader.pack(side="left", fill="both", expand=True)
        ttk.Label(
            reader, text="УЧЕБНЫЙ МАТЕРИАЛ", style="Card.TLabel",
            foreground="#087e8b", font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", pady=(0, 8))
        self.lesson_title = ttk.Label(
            reader, text="Выбери учебный раздел", style="Card.TLabel",
            foreground="#17365d", font=("Segoe UI", 18, "bold"), wraplength=650,
        )
        self.lesson_title.pack(anchor="w", pady=(0, 18))
        ttk.Separator(reader).pack(fill="x", pady=(0, 16))
        text_frame = ttk.Frame(reader, style="Card.TFrame")
        text_frame.pack(fill="both", expand=True)
        text_scroll = ttk.Scrollbar(text_frame, orient="vertical")
        text_scroll.pack(side="right", fill="y")
        self.lesson_body = tk.Text(
            text_frame,
            wrap="word",
            font=("Segoe UI", 12),
            padx=4,
            pady=4,
            bd=0,
            relief="flat",
            bg="#ffffff",
            fg="#334155",
            insertbackground="#334155",
            selectbackground="#d5efed",
            highlightthickness=0,
            spacing1=3,
            spacing3=11,
            yscrollcommand=text_scroll.set,
        )
        self.lesson_body.pack(side="left", fill="both", expand=True)
        text_scroll.configure(command=self.lesson_body.yview)
        self.lesson_body.configure(state="disabled")
        if lessons:
            self.lesson_list.selection_set(0)
            self.show_lesson()
        else:
            self.lesson_body.configure(state="normal")
            self.lesson_body.insert("1.0", "Учебные материалы пока не добавлены.")
            self.lesson_body.configure(state="disabled")

    def show_lesson(self, _event=None) -> None:
        selection = self.lesson_list.curselection()
        if not selection:
            return
        lesson = self.lessons[selection[0]]
        self.lesson_title.configure(text=lesson["title"])
        self.lesson_body.configure(state="normal")
        self.lesson_body.delete("1.0", "end")
        self.lesson_body.insert("1.0", lesson["body"])
        self.lesson_body.configure(state="disabled")

    def build_test_tab(self) -> None:
        ttk.Label(self.content_area, text="Проверка знаний", style="Section.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            self.content_area,
            text="Ответь на вопросы по выбранному разделу и сразу узнай результат.",
            foreground="#687b90",
        ).pack(anchor="w", pady=(4, 18))
        selector = ttk.Frame(self.content_area, style="Card.TFrame", padding=14)
        selector.pack(fill="x", pady=(0, 14))
        ttk.Label(selector, text="Раздел курса", style="Card.TLabel").pack(side="left", padx=(0, 12))
        self.test_lesson = tk.StringVar()
        self.test_selector = ttk.Combobox(
            selector,
            textvariable=self.test_lesson,
            values=[lesson["title"] for lesson in self.lessons],
            state="readonly",
            width=48,
        )
        self.test_selector.pack(side="left")
        self.test_selector.current(0 if self.lessons else -1)
        ttk.Button(
            selector, text="Загрузить тест", style="Primary.TButton", command=self.load_test
        ).pack(side="left", padx=12)
        self.test_content = ttk.Frame(self.content_area)
        self.test_content.pack(fill="both", expand=True)
        self.question_vars: list[tuple[object, dict]] = []
        self.load_test()

    def load_test(self) -> None:
        for widget in self.test_content.winfo_children():
            widget.destroy()
        self.question_vars.clear()
        lesson_index = self.test_selector.current()
        if lesson_index < 0:
            ttk.Label(self.test_content, text="Тесты пока не добавлены.").pack(anchor="w")
            return
        lesson = self.lessons[lesson_index]
        try:
            questions = self.db.get_questions(lesson["slug"])
        except Exception as exc:
            messagebox.showerror("Ошибка загрузки теста", str(exc), parent=self.root)
            return
        if not questions:
            ttk.Label(
                self.test_content,
                text="Для этого раздела пока нет вопросов.",
                style="Card.TLabel",
            ).pack(anchor="w", padx=16, pady=16)
            return
        random.shuffle(questions)
        questions = [shuffle_question(question) for question in questions]

        canvas = tk.Canvas(
            self.test_content, background="#f1f5f9", highlightthickness=0, borderwidth=0
        )
        question_scroll = ttk.Scrollbar(
            self.test_content, orient="vertical", command=canvas.yview
        )
        question_list = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=question_list, anchor="nw")

        def update_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def resize_question_list(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        question_list.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", resize_question_list)
        canvas.configure(yscrollcommand=question_scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        question_scroll.pack(side="right", fill="y")

        for index, question in enumerate(questions, start=1):
            group = ttk.LabelFrame(
                question_list, text=f"Вопрос {index}   ·   {question['prompt']}", padding=14
            )
            group.pack(fill="x", anchor="n", pady=(0, 10), padx=(0, 8))
            if question["question_type"] == "single":
                answer = tk.IntVar(value=-1)
                for option_index, option in enumerate(question["options"]):
                    ttk.Radiobutton(
                        group,
                        text=option,
                        variable=answer,
                        value=option_index,
                    ).pack(anchor="w", pady=3)
            elif question["question_type"] == "multiple":
                answer = [
                    tk.BooleanVar(value=False) for _option in question["options"]
                ]
                for option_index, option in enumerate(question["options"]):
                    ttk.Checkbutton(
                        group,
                        text=option,
                        variable=answer[option_index],
                    ).pack(anchor="w", pady=3)
            else:
                answer = ttk.Entry(group, width=70)
                answer.pack(fill="x", pady=(6, 2))
            self.question_vars.append((answer, question))
        ttk.Button(
            question_list,
            text="Завершить тест",
            style="Primary.TButton",
            command=self.submit_test,
        ).pack(anchor="e", pady=(4, 14), padx=(0, 8))

    def submit_test(self) -> None:
        if not self.question_vars:
            return
        if any(
            (
                answer.get() < 0
                if question["question_type"] == "single"
                else not any(variable.get() for variable in answer)
                if question["question_type"] == "multiple"
                else not answer.get().strip()
            )
            for answer, question in self.question_vars
        ):
            messagebox.showwarning("Тест", "Ответьте на все вопросы.", parent=self.root)
            return
        correct = sum(
            check_question_answer(
                question,
                (
                    answer.get()
                    if question["question_type"] == "single"
                    else {
                        index
                        for index, variable in enumerate(answer)
                        if variable.get()
                    }
                    if question["question_type"] == "multiple"
                    else answer.get()
                ),
            )
            for answer, question in self.question_vars
        )
        total = len(self.question_vars)
        percentage, grade = calculate_grade(correct, total)
        lesson = self.lessons[self.test_selector.current()]
        try:
            self.db.save_attempt(
                self.student["id"], lesson["slug"], correct, total, percentage, grade
            )
        except Exception as exc:
            messagebox.showerror("Не удалось сохранить результат", str(exc), parent=self.root)
            return
        messagebox.showinfo(
            "Результат теста",
            f"Правильных ответов: {correct} из {total}\n"
            f"Результат: {percentage}%\nОценка: {grade}",
            parent=self.root,
        )
        self.refresh_results()

    def build_practice_tab(self) -> None:
        ttk.Label(self.content_area, text="Практические работы", style="Section.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            self.content_area,
            text="Отправь выполненное задание в формате Word или Excel.",
            foreground="#687b90",
        ).pack(anchor="w", pady=(4, 18))
        try:
            tasks = self.db.get_practice_tasks()
        except Exception as exc:
            messagebox.showerror("Ошибка загрузки заданий", str(exc), parent=self.root)
            return
        if not tasks:
            ttk.Label(
                self.content_area, text="Практические задания пока не добавлены."
            ).pack(anchor="w")
            return
        list_frame = ttk.Frame(self.content_area)
        list_frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(
            list_frame, background="#f1f5f9", highlightthickness=0, borderwidth=0
        )
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        cards = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=cards, anchor="nw")
        cards.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(window_id, width=event.width),
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        for task in tasks:
            practice_number = int(task["slug"].removeprefix("practice-"))
            card = ttk.LabelFrame(cards, text=task["title"], padding=(18, 15))
            card.pack(fill="x", pady=(0, 12))
            ttk.Label(
                card,
                text=task["instructions"],
                wraplength=760,
                justify="left",
                style="Card.TLabel",
            ).pack(anchor="w", pady=(0, 12))
            extensions = ", ".join(f".{item}" for item in task["extensions"])
            actions = ttk.Frame(card, style="Card.TFrame")
            actions.pack(anchor="w")
            if practice_number in PRACTICES_WITH_VARIANTS:
                assigned_variant = None
                try:
                    variant_options = self.db.get_practice_variant_options(
                        self.student["id"], task["slug"]
                    )
                    assigned_variant = variant_options["current"]
                    variant_label = (
                        f"Мой вариант: №{assigned_variant:02d} · изменить"
                        if assigned_variant is not None
                        else "Выбрать вариант"
                    )
                except Exception as exc:
                    messagebox.showerror(
                        "Не удалось загрузить вариант", str(exc), parent=self.root
                    )
                    variant_label = "Выбрать вариант"
                variant_button = ttk.Button(
                    actions,
                    text=variant_label,
                )
                download_button = ttk.Button(
                    actions,
                    text=(
                        f"Скачать вариант №{assigned_variant:02d}"
                        if assigned_variant is not None
                        else "Скачать мой вариант"
                    ),
                    style="Primary.TButton",
                    command=lambda current_task=task, number=practice_number:
                        self.download_assigned_practice(current_task, number),
                )
                variant_button.configure(
                    command=lambda current_task=task, number=practice_number,
                    button=variant_button, download=download_button:
                        self.choose_practice_variant(
                            current_task, number, button, download
                        )
                )
                variant_button.pack(side="left", padx=(0, 8))
                download_button.pack(side="left", padx=(0, 8))
            else:
                ttk.Button(
                    actions,
                    text="Скачать шаблон",
                    command=lambda current_task=task, number=practice_number:
                        self.download_assigned_practice(current_task, number),
                ).pack(side="left", padx=(0, 8))
            if practice_number == 6:
                ttk.Button(
                    actions,
                    text="Скачать исходный CSV",
                    command=lambda current_task=task:
                        self.download_practice_csv(current_task),
                ).pack(side="left", padx=(0, 8))
            ttk.Button(
                actions,
                text=f"Загрузить свою работу  ·  {extensions}",
                style="Primary.TButton",
                command=lambda current_task=task, number=practice_number:
                    self.upload_practice(current_task, number),
            ).pack(side="left")
        note = ttk.Frame(self.content_area, style="Card.TFrame", padding=16)
        note.pack(fill="x", pady=(4, 0))
        ttk.Label(
            note,
            text="После отправки работа сохранится в твоей папке. Оценку за практику выставит преподаватель.",
            wraplength=780,
            style="Card.TLabel",
        ).pack(anchor="w")

    def choose_practice_variant(
        self,
        task: dict,
        practice_number: int,
        button: ttk.Button,
        download_button: ttk.Button,
    ) -> None:
        try:
            options = self.db.get_practice_variant_options(
                self.student["id"], task["slug"]
            )
        except Exception as exc:
            messagebox.showerror("Не удалось загрузить варианты", str(exc), parent=self.root)
            return
        available = options["available"]
        if not available:
            messagebox.showerror(
                "Нет свободных вариантов",
                "Все 30 вариантов этой практической работы уже заняты.",
                parent=self.root,
            )
            return

        window = tk.Toplevel(self.root)
        window.title(f"Вариант практической работы №{practice_number}")
        window.transient(self.root)
        window.resizable(False, False)
        window.grab_set()
        frame = ttk.Frame(window, padding=22)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text=f"Выбери вариант для практической работы №{practice_number}",
            style="Section.TLabel",
        ).pack(anchor="w", pady=(0, 12))
        ttk.Label(
            frame,
            text="Можно выбрать любой вариант. После сохранения скачай файл заново.",
            wraplength=420,
        ).pack(anchor="w", pady=(0, 10))
        selected = tk.StringVar()
        selector = ttk.Combobox(
            frame,
            textvariable=selected,
            values=[f"{variant:02d}" for variant in available],
            state="readonly",
            width=12,
        )
        selector.pack(anchor="w")
        initial = options["current"] or available[0]
        selector.current(available.index(initial))

        def save_variant() -> None:
            try:
                variant = self.db.set_practice_variant(
                    self.student["id"], task["slug"], int(selected.get())
                )
            except Exception as exc:
                messagebox.showerror(
                    "Не удалось изменить вариант", str(exc), parent=window
                )
                return
            button.configure(text=f"Мой вариант: №{variant:02d} · изменить")
            download_button.configure(text=f"Скачать вариант №{variant:02d}")
            window.destroy()
            messagebox.showinfo(
                "Вариант сохранён",
                f"Для практической работы №{practice_number} выбран вариант {variant:02d}.",
                parent=self.root,
            )

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(18, 0))
        ttk.Button(actions, text="Отмена", command=window.destroy).pack(side="right")
        ttk.Button(
            actions,
            text="Сохранить",
            style="Primary.TButton",
            command=save_variant,
        ).pack(side="right", padx=(0, 8))

    def download_practice_csv(self, task: dict) -> None:
        try:
            variant = self.db.get_or_assign_practice_variant(
                self.student["id"], task["slug"]
            )
            content = build_practice_csv(variant)
        except Exception as exc:
            messagebox.showerror("Не удалось подготовить CSV", str(exc), parent=self.root)
            return
        destination = filedialog.asksaveasfilename(
            parent=self.root,
            title=f"Сохранить CSV для варианта {variant:02d}",
            initialfile=f"Практика_6_вариант_{variant:02d}.csv",
            defaultextension=".csv",
            filetypes=[("Файл CSV", "*.csv")],
        )
        if not destination:
            return
        if not self.save_generated_file(
            destination,
            content,
            "CSV",
            ".csv",
            [("Файл CSV", "*.csv")],
        ):
            return
        messagebox.showinfo(
            "CSV сохранён",
            "Импортируйте этот файл в Excel для выполнения практической работы №6.",
            parent=self.root,
        )

    def download_assigned_practice(self, task: dict, practice_number: int) -> None:
        try:
            variant = (
                self.db.get_or_assign_practice_variant(
                    self.student["id"], task["slug"]
                )
                if practice_number in PRACTICES_WITH_VARIANTS
                else None
            )
            extension, content = build_practice_file(practice_number, variant)
        except Exception as exc:
            messagebox.showerror("Не удалось подготовить вариант", str(exc), parent=self.root)
            return
        variant_suffix = f"_вариант_{variant:02d}" if variant is not None else ""
        destination = filedialog.asksaveasfilename(
            parent=self.root,
            title=(
                f"Сохранить вариант {variant:02d}"
                if variant is not None
                else "Сохранить шаблон практической работы"
            ),
            initialfile=f"Практика_{practice_number}{variant_suffix}{extension}",
            defaultextension=extension,
            filetypes=[
                ("Документ Word", "*.docx")
                if extension == ".docx"
                else ("Книга Excel", "*.xlsx")
            ],
        )
        if not destination:
            return
        filetypes = [
            ("Документ Word", "*.docx")
            if extension == ".docx"
            else ("Книга Excel", "*.xlsx")
        ]
        if not self.save_generated_file(
            destination,
            content,
            "файл практической работы",
            extension,
            filetypes,
        ):
            return
        messagebox.showinfo(
            "Файл сохранён",
            (
                f"Сохранён файл практической работы №{practice_number}, вариант {variant:02d}."
                if variant is not None
                else "Сохранён общий шаблон. Выполняй работу по теме своего ТЗ из практики №1."
            ),
            parent=self.root,
        )

    def save_generated_file(
        self,
        destination: str,
        content: bytes,
        file_description: str,
        extension: str,
        filetypes: list[tuple[str, str]],
    ) -> bool:
        destination_path = Path(destination)
        try:
            destination_path.write_bytes(content)
            return True
        except PermissionError as exc:
            retry = messagebox.askyesno(
                "Файл занят или нет доступа",
                f"Не удалось сохранить {file_description}:\n{destination_path}\n\n"
                "Возможно, файл открыт в Word или Excel. Закрой его или выбери "
                "другое место для сохранения копии.",
                parent=self.root,
            )
            if not retry:
                return False
            alternative = filedialog.asksaveasfilename(
                parent=self.root,
                title=f"Сохранить копию: {file_description}",
                initialdir=str(destination_path.parent),
                initialfile=f"{destination_path.stem}_копия{extension}",
                defaultextension=extension,
                filetypes=filetypes,
            )
            if not alternative:
                return False
            try:
                Path(alternative).write_bytes(content)
                return True
            except OSError as retry_error:
                messagebox.showerror(
                    "Не удалось сохранить копию",
                    str(retry_error),
                    parent=self.root,
                )
                return False
        except OSError as exc:
            messagebox.showerror(
                f"Не удалось сохранить {file_description}",
                str(exc),
                parent=self.root,
            )
            return False

    def upload_practice(self, task: dict, practice_number: int) -> None:
        extensions = [f".{extension}" for extension in task["extensions"]]
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Выберите практическую работу",
            filetypes=[("Документы Word и Excel", " ".join(f"*{ext}" for ext in extensions))],
        )
        if not selected:
            return
        try:
            variant = None
            if practice_number in PRACTICES_WITH_VARIANTS:
                variant = self.db.get_or_assign_practice_variant(
                    self.student["id"], task["slug"]
                )
                if not messagebox.askyesno(
                    "Подтвердить вариант",
                    f"Отправить файл как вариант {variant:02d} практической работы "
                    f"№{practice_number}?",
                    parent=self.root,
                ):
                    return
            source = validate_practice_file(selected)
            if source.suffix.lower() not in extensions:
                raise ValueError(f"Для этого задания разрешены: {', '.join(extensions)}.")
            file_data = source.read_bytes()
            if len(file_data) > MAX_UPLOAD_SIZE:
                raise ValueError("Размер файла не должен превышать 20 МБ.")
            self.db.save_submission(
                self.student["id"],
                task["slug"],
                source.name,
                file_data,
                variant,
            )
        except Exception as exc:
            messagebox.showerror("Не удалось отправить работу", str(exc), parent=self.root)
            return
        messagebox.showinfo(
            "Работа сохранена",
            "Файл сохранён в общей базе данных. Преподаватель сможет открыть его "
            "со своего компьютера.",
            parent=self.root,
        )
        self.refresh_results()

    def build_results_tab(self) -> None:
        ttk.Label(self.content_area, text="Мои результаты", style="Section.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            self.content_area,
            text="История тестов и практических заданий.",
            foreground="#687b90",
        ).pack(anchor="w", pady=(4, 18))
        table_card = ttk.Frame(self.content_area, style="Card.TFrame", padding=12)
        table_card.pack(fill="both", expand=True)
        columns = ("type", "title", "grade", "percentage", "created")
        self.results_table = ttk.Treeview(table_card, columns=columns, show="headings")
        headings = {
            "type": ("Вид работы", 130),
            "title": ("Раздел / задание", 330),
            "grade": ("Оценка", 120),
            "percentage": ("Результат теста", 150),
            "created": ("Дата", 180),
        }
        for column, (title, width) in headings.items():
            self.results_table.heading(column, text=title)
            self.results_table.column(column, width=width, anchor="w")
        self.results_table.pack(fill="both", expand=True)
        self.refresh_results()

    def refresh_results(self) -> None:
        if (
            not hasattr(self, "results_table")
            or not self.results_table.winfo_exists()
            or self.student is None
        ):
            return
        try:
            results = self.db.get_results(self.student["id"])
        except Exception as exc:
            messagebox.showerror("Не удалось загрузить результаты", str(exc), parent=self.root)
            return
        self.results_table.delete(*self.results_table.get_children())
        for result in results:
            grade = result["grade"] if result["grade"] is not None else "Ожидает оценки"
            percentage = f"{result['percentage']}%" if result["percentage"] is not None else "—"
            created = result["created_at"].astimezone().strftime("%d.%m.%Y %H:%M")
            self.results_table.insert(
                "", "end", values=(result["result_type"], result["title"], grade, percentage, created)
            )

    def logout(self) -> None:
        self.student = None
        self.show_login()


def main() -> None:
    root = tk.Tk()
    TextbookApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
