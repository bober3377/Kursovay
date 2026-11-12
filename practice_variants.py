from __future__ import annotations

import csv
from datetime import date, timedelta
from io import BytesIO, StringIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


PRACTICE_NUMBERS = (1, 2, 3, 5, 6, 7, 8, 9)
PRACTICES_WITH_VARIANTS = frozenset((1, 3, 5, 6, 7, 8, 9))

SCENARIOS = [
    ("сервис бронирования переговорных комнат", "сотрудники офиса", "бронь комнаты по времени"),
    ("система учёта заявок в сервисный центр", "операторы поддержки", "создание и закрытие заявки"),
    ("приложение записи к врачу", "пациенты и регистраторы", "поиск свободного времени"),
    ("каталог учебных курсов", "студенты и преподаватели", "запись на курс"),
    ("сервис доставки готовых обедов", "клиенты и курьеры", "отслеживание заказа"),
    ("система выдачи книг в библиотеке", "читатели и библиотекари", "поиск и выдача книги"),
    ("приложение контроля личных расходов", "пользователи", "учёт расходов по категориям"),
    ("сервис регистрации на мероприятие", "гости и организаторы", "регистрация участника"),
    ("система заявок на ремонт оборудования", "сотрудники организации", "назначение исполнителя"),
    ("онлайн-каталог компьютерной техники", "покупатели и менеджеры", "сравнение товаров"),
    ("система планирования занятий", "студенты и учебная часть", "просмотр расписания"),
    ("сервис аренды спортивного инвентаря", "клиенты и администратор", "оформление аренды"),
    ("система учёта задач команды разработки", "разработчики и руководитель", "назначение задачи"),
    ("приложение заказа такси", "пассажиры и водители", "расчёт стоимости поездки"),
    ("сервис поиска стажировок", "студенты и работодатели", "отклик на вакансию"),
    ("система учёта складских товаров", "кладовщики и менеджеры", "списание товара"),
    ("приложение записи в коворкинг", "посетители и администратор", "выбор рабочего места"),
    ("сервис сбора обратной связи", "клиенты и аналитики", "отправка отзыва"),
    ("система регистрации оборудования", "сотрудники и системные администраторы", "выдача устройства"),
    ("сервис поиска попутчиков", "пассажиры и водители", "создание совместной поездки"),
    ("приложение учёта домашних заданий", "ученики и преподаватели", "сдача задания"),
    ("система управления заявками на отпуск", "сотрудники и кадровая служба", "согласование отпуска"),
    ("сервис заказа пропусков", "посетители и служба безопасности", "создание временного пропуска"),
    ("система мониторинга серверов", "системные администраторы", "просмотр критического события"),
    ("приложение бронирования учебных аудиторий", "преподаватели и диспетчеры", "резервирование аудитории"),
    ("сервис учёта волонтёрских смен", "волонтёры и координаторы", "запись на смену"),
    ("система заказов типографии", "клиенты и сотрудники типографии", "расчёт тиража"),
    ("приложение учёта компьютерных комплектующих", "кладовщики и инженеры", "поиск совместимой детали"),
    ("сервис планирования командировок", "сотрудники и бухгалтерия", "согласование поездки"),
    ("система регистрации обращений жителей", "жители и городская служба", "отправка обращения"),
]

SCENARIO_REQUIREMENTS = [
    ("найти переговорную нужной вместимости и забронировать её на выбранный интервал",
     "нельзя подтверждать две брони одной комнаты на пересекающееся время",
     "в результатах остаются только свободные комнаты, подходящие по вместимости и оборудованию"),
    ("зарегистрировать обращение, назначить исполнителя и закрыть заявку с комментарием",
     "у каждой заявки должны быть приоритет, ответственный и срок реакции",
     "просроченные заявки видны отдельно, а закрытие без комментария запрещено"),
    ("выбрать врача, специализацию и доступное время для записи пациента",
     "один временной слот нельзя назначить двум пациентам",
     "после подтверждения слот исчезает из списка свободных и появляется в расписании врача"),
    ("подобрать учебный курс по уровню и записать студента в группу",
     "число записанных не может превышать вместимость группы",
     "при заполненной группе система предлагает лист ожидания, а не подтверждает запись"),
    ("оформить заказ обеда с адресом доставки и выбранным временным интервалом",
     "заказ нельзя принять, если адрес вне зоны доставки или кухня закрыта",
     "итоговая цена включает выбранные блюда и стоимость доставки"),
    ("найти книгу по автору или названию и оформить её выдачу читателю",
     "нельзя выдать экземпляр, который уже числится выданным",
     "карточка книги показывает наличие экземпляров и дату возврата"),
    ("добавить расход, выбрать категорию и сформировать отчёт за месяц",
     "сумма расхода должна быть положительной, а дата не может быть будущей",
     "итог отчёта совпадает с суммой операций и разбит по категориям"),
    ("создать мероприятие, открыть регистрацию и выгрузить список участников",
     "регистрация закрывается при достижении лимита мест",
     "в списке нет повторных регистраций одного участника на одно мероприятие"),
    ("создать заявку на ремонт, приложить описание неисправности и назначить мастера",
     "мастер не может получить новую заявку, если его смена закончилась",
     "статус заявки и назначенный исполнитель видны заявителю"),
    ("сравнить компьютерные товары по цене, характеристикам и наличию",
     "в сравнении отображаются только товары с заполненными обязательными характеристиками",
     "покупатель может отсортировать результат по цене и увидеть различия моделей"),
    ("составить расписание занятий для группы и проверить свободную аудиторию",
     "у преподавателя и аудитории не должно быть двух занятий одновременно",
     "сохранённое расписание показывает конфликт до публикации"),
    ("выбрать спортивный инвентарь, срок аренды и оформить возврат",
     "нельзя выдать единицу инвентаря, уже арендованную на выбранный срок",
     "стоимость аренды пересчитывается при изменении срока"),
    ("создать задачу разработки, назначить исполнителя и связать её со спринтом",
     "задача не попадает в спринт, если её оценка превышает оставшуюся ёмкость команды",
     "доска показывает исполнителя, приоритет, статус и оставшуюся оценку"),
    ("указать адреса поездки и выбрать подходящий тариф такси",
     "заказ не подтверждается без адреса подачи и способа оплаты",
     "стоимость рассчитывается по тарифу и расстоянию до подтверждения поездки"),
    ("найти стажировку по направлению и отправить работодателю отклик",
     "нельзя отправить второй отклик на одну вакансию с того же аккаунта",
     "кандидат видит статус отклика и дату последнего изменения"),
    ("принять товар на склад, указать ячейку хранения и оформить списание",
     "списание не может уменьшить остаток ниже нуля",
     "после операции журнал показывает количество до и после изменения"),
    ("выбрать рабочее место и забронировать его на часть рабочего дня",
     "одно место нельзя забронировать разными посетителями на пересекающееся время",
     "пользователь видит план коворкинга с доступными и занятыми местами"),
    ("создать отзыв с оценкой и передать его в очередь модерации",
     "публикация недоступна без оценки от 1 до 5 и текста отзыва",
     "аналитик может отфильтровать отзывы по оценке и дате"),
    ("зарегистрировать устройство, указать инвентарный номер и назначить владельца",
     "инвентарный номер должен быть уникальным",
     "по карточке видны владелец, дата выдачи и история перемещений"),
    ("создать поездку с маршрутом и найти попутчика с совпадающим участком пути",
     "пассажир не может бронировать больше свободных мест в автомобиле",
     "результат поиска показывает совпадающий маршрут, дату и число мест"),
    ("создать задание ученику, установить срок и проверить загруженный ответ",
     "после срока сдачи работа помечается как просроченная, но не удаляется",
     "преподаватель видит статус сдачи и оценку по каждому ученику"),
    ("подать заявление на отпуск с датами и передать его руководителю на согласование",
     "заявления одного сотрудника не могут пересекаться по датам",
     "после решения руководителя сотрудник получает статус и комментарий"),
    ("создать временный пропуск с датой действия и принимающим сотрудником",
     "срок действия пропуска не может превышать один рабочий день",
     "служба безопасности видит активные, истёкшие и отменённые пропуска"),
    ("зарегистрировать сервер и настроить порог уведомления о высокой нагрузке",
     "уведомление создаётся только при превышении порога в двух последовательных проверках",
     "панель показывает время события, сервер и значение метрики"),
    ("выбрать аудиторию для занятия с учётом вместимости и оборудования",
     "аудитория должна быть свободна весь интервал занятия",
     "диспетчер получает список подходящих помещений и причину исключения остальных"),
    ("создать волонтёрскую смену и записаться на неё с выбранной ролью",
     "число участников каждой роли не может превышать лимит смены",
     "координатор видит заполненность ролей и список подтверждённых волонтёров"),
    ("оформить заказ печати с тиражом, форматом бумаги и типом переплёта",
     "тираж должен быть положительным, а выбранный переплёт доступен для этого формата",
     "смета пересчитывается при изменении тиража и параметров заказа"),
    ("подобрать совместимую комплектующую по разъёму, форм-фактору и мощности",
     "несовместимые с выбранной моделью детали нельзя добавить в подборку",
     "результат содержит объяснение совместимости по каждому обязательному параметру"),
    ("создать командировку с маршрутом, сроками и статьями расходов",
     "дата окончания не может быть раньше даты начала, расходы требуют категории",
     "отчёт показывает плановые и фактические расходы по каждой статье"),
    ("зарегистрировать обращение жителя, указать адрес и передать его исполнителю",
     "обращение нельзя закрыть без результата проверки и ответа заявителю",
     "житель видит номер обращения, статус и срок ответа"),
]

SOFTWARE_NAMES = [
    "RoomPlan", "HelpDesk Lite", "МедЗапись", "CourseHub", "LunchTrack",
    "БиблиоПлюс", "МойБюджет", "EventPass", "РемонтПро", "TechChoice",
    "Расписание+", "SportRent", "SprintBoard", "ТаксиПульс", "StartCareer",
    "СкладПрофи", "WorkSpace", "Отзывы360", "DeviceLog", "ВместеПуть",
    "Домашка+", "ОтпускОК", "Пропуск24", "ServerWatch", "Аудитория",
    "Волонтёр", "PrintCalc", "PartMatch", "ПоездкаПро", "ГородНаСвязи",
]

DEVELOPERS = [
    ("Анна", "Junior"),
    ("Борис", "Middle"),
    ("Вера", "Senior"),
    ("Глеб", "Junior"),
    ("Дарья", "Middle"),
    ("Егор", "Senior"),
]

NAVY = "17365D"
TEAL = "087E8B"
PALE = "EAF2F8"
WHITE = "FFFFFF"
TEXT = "304153"
GREEN = "D9EAD3"
RED = "F4CCCC"
THIN_BLUE = Side(style="thin", color="CBD7E4")


def _scenario(variant: int) -> tuple[str, str, str]:
    if not 1 <= variant <= 30:
        raise ValueError("Номер варианта должен быть от 1 до 30.")
    return SCENARIOS[variant - 1]


def _scenario_requirements(variant: int) -> tuple[str, str, str]:
    _scenario(variant)
    return SCENARIO_REQUIREMENTS[variant - 1]


def _set_cell_shading(cell, color: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color)
    properties.append(shading)


def _style_docx(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(TEXT)
    normal.paragraph_format.space_after = Pt(5)
    for name, size in (("Title", 22), ("Heading 1", 15), ("Heading 2", 12)):
        style = document.styles[name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(NAVY)
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header_run = header.add_run("ЭЛЕКТРОННЫЙ УЧЕБНИК · ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ")
    header_run.font.size = Pt(8)
    header_run.font.color.rgb = RGBColor.from_string(TEAL)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Практическая работа · ").font.size = Pt(8)
    _page_field(footer)


def _page_field(paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    result = OxmlElement("w:t")
    result.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instruction, separate, result, end):
        run._r.append(element)


def _add_docx_table(document: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for index, value in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = value
        _set_cell_shading(cell, NAVY)
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = value
            if row_index % 2 == 0:
                _set_cell_shading(cells[index], PALE)


def _docx_intro(
    document: Document, number: int, variant: int | None, title: str
) -> None:
    eyebrow = document.add_paragraph()
    eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    variant_label = f" · ВАРИАНТ {variant:02d}" if variant is not None else ""
    run = eyebrow.add_run(f"ПРАКТИЧЕСКОЕ ЗАНЯТИЕ №{number}{variant_label}")
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(TEAL)
    heading = document.add_heading(title, 0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    student = document.add_paragraph("Студент: ______________________________    Группа: ______________")
    student.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _add_toc_field(paragraph) -> None:
    run = paragraph.add_run()
    for tag, text in (
        ("begin", None),
        ("instruction", ' TOC \\o "1-3" \\h \\z \\u '),
        ("separate", None),
        ("text", "В Microsoft Word нажмите правой кнопкой и выберите «Обновить поле»."),
        ("end", None),
    ):
        if tag == "instruction":
            element = OxmlElement("w:instrText")
            element.set(qn("xml:space"), "preserve")
            element.text = text
        elif tag == "text":
            element = OxmlElement("w:t")
            element.text = text
        else:
            element = OxmlElement("w:fldChar")
            element.set(qn("w:fldCharType"), tag)
        run._r.append(element)


def _add_answer_lines(document: Document, count: int = 2) -> None:
    for _ in range(count):
        p = document.add_paragraph("________________________________________________________________________")
        p.paragraph_format.space_after = Pt(5)


def _build_word_1(variant: int) -> bytes:
    product, users, feature = _scenario(variant)
    workflow, rule, acceptance = _scenario_requirements(variant)
    document = Document()
    _style_docx(document)
    _docx_intro(document, 1, variant, "Техническое задание на разработку ПО")
    document.add_heading(f"Условия вашего варианта · {SOFTWARE_NAMES[variant - 1]}", 1)
    document.add_paragraph(
        f"Сценарий: {workflow}. Правило: {rule}. "
        f"Проверка результата: {acceptance}."
    )
    document.add_heading("Часть 1. Общий образец выполнения", 1)
    document.add_paragraph(
        "Тема образца: информационная система учёта аудиторий для колледжа."
    )
    document.add_heading("1. Общие сведения", 2)
    document.add_paragraph(
        "Программа «Аудитория» предназначена для учёта учебных помещений "
        "и расписания занятий."
    )
    document.add_heading("2. Назначение и цели", 2)
    document.add_paragraph(
        "Цель разработки — дать диспетчеру возможность быстро находить "
        "свободную аудиторию и проверять её вместимость."
    )
    document.add_heading("3. Требования", 2)
    document.add_paragraph(
        "Система должна хранить номер помещения, вместимость, оборудование "
        "и временные интервалы занятий."
    )
    document.add_heading("4. Проверка результата", 2)
    document.add_paragraph(
        "Пользователь выбирает дату и время; программа показывает только "
        "свободные аудитории, удовлетворяющие запросу."
    )
    document.add_heading("Автоматическое оглавление", 2)
    _add_toc_field(document.add_paragraph())

    document.add_page_break()
    document.add_heading("Часть 2. Задание для самостоятельной работы", 1)
    document.add_paragraph(
        f"Разработайте техническое задание для продукта «{SOFTWARE_NAMES[variant - 1]}» "
        f"для пользователей «{users}». Основная функция: {feature}. "
        f"Рабочий сценарий: {workflow}."
    )
    document.add_paragraph(f"Обязательное бизнес-правило: {rule}.")
    document.add_paragraph(f"Критерий приёмки: {acceptance}.")
    document.add_paragraph(
        "Создайте документ по структуре примера. Добавьте многоуровневые "
        "заголовки, автоматическое оглавление, нумерацию страниц и элементов."
    )
    for heading, prompt in [
        ("1. Общие сведения", "Укажите название, заказчика и назначение продукта."),
        ("2. Основание и цели разработки", "Опишите проблему и ожидаемый результат."),
        ("3. Требования к системе", f"Опишите функции и правило: {rule}."),
        ("4. Требования к документации", "Укажите состав документации и правила оформления."),
        ("5. Контроль и приёмка", f"Проверьте критерий: {acceptance}."),
    ]:
        document.add_heading(heading, 2)
        document.add_paragraph(prompt)
        _add_answer_lines(document, 1)
    document.add_heading("Оглавление", 2)
    _add_toc_field(document.add_paragraph())
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def _build_word_2(variant: int | None) -> bytes:
    if variant is None:
        users = None
        feature = None
    else:
        _product, users, feature = _scenario(variant)
    document = Document()
    _style_docx(document)
    _docx_intro(document, 2, variant, "Руководство пользователя")
    document.add_heading("Часть 1. Образец выполнения", 1)
    document.add_paragraph("Пример: руководство к системе записи в учебную библиотеку.")
    document.add_heading("Вход в систему", 2)
    document.add_paragraph("1. Откройте страницу приложения.\n2. Введите логин и пароль.\n3. Нажмите «Войти».")
    document.add_heading("Поиск книги", 2)
    document.add_paragraph(
        "1. Откройте каталог.\n2. Введите название книги.\n3. Выберите книгу из результатов.\n"
        "4. Нажмите «Забронировать»."
    )
    _add_docx_table(
        document,
        ["Место для схемы алгоритма", "Аннотированный снимок экрана"],
        [["Вставьте блок-схему действия «Поиск книги».", "Вставьте экран каталога; подпишите поле поиска и кнопку."]],
    )

    document.add_page_break()
    document.add_heading("Часть 2. Задание для самостоятельной работы", 1)
    if variant is None:
        document.add_paragraph(
            "Подготовьте руководство для программного продукта из вашего ТЗ "
            "по практической работе №1. Используйте указанные там назначение, "
            "пользователей и основной сценарий."
        )
    else:
        document.add_paragraph(
            f"Подготовьте руководство пользователя для системы «{SOFTWARE_NAMES[variant - 1]}», "
            f"предназначенной для пользователей «{users}». Основная функция варианта: {feature}."
        )
    for heading, prompt in [
        ("1. Назначение программы", "Опишите назначение и целевых пользователей."),
        ("2. Начало работы", "Опишите запуск, вход и первоначальную настройку."),
        (
            "3. Основной сценарий",
            (
                f"Дайте пошаговую инструкцию по операции «{feature}»."
                if feature is not None
                else "Опишите пошаговый сценарий, указанный в вашем ТЗ."
            ),
        ),
        ("4. Возможные ошибки", "Опишите не менее двух ошибок и действия пользователя."),
    ]:
        document.add_heading(heading, 2)
        document.add_paragraph(prompt)
        _add_answer_lines(document, 1)
    document.add_heading("5. Схема алгоритма", 2)
    document.add_paragraph("Постройте блок-схему основного сценария и вставьте её сюда.")
    _add_answer_lines(document, 2)
    document.add_heading("6. Аннотированные снимки экрана", 2)
    document.add_paragraph(
        "Вставьте не менее двух снимков своей программы. Добавьте подписи, "
        "стрелки или рамки для важных элементов."
    )
    _add_answer_lines(document, 3)
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def _build_word_3(variant: int) -> bytes:
    product_area, users, feature = _scenario(variant)
    workflow, rule, acceptance = _scenario_requirements(variant)
    product_name = SOFTWARE_NAMES[variant - 1]
    amount = 165_000 + variant * 11_750
    old_days = 10 + variant % 11
    new_days = old_days + 5 + variant % 7
    warranty_days = 30 + variant * 3
    advance_percent = 20 + variant % 5 * 5
    document = Document()
    _style_docx(document)
    _docx_intro(document, 3, variant, "Сравнение редакций договора")
    document.add_heading("Часть 1. Образец выполнения", 1)
    document.add_paragraph("Фрагмент договора на разработку тестовой информационной системы.")
    _add_docx_table(
        document,
        ["Условие", "Редакция заказчика", "Редакция исполнителя", "Разногласие"],
        [[
            "Срок работ",
            "30 календарных дней",
            "40 календарных дней",
            "Стороны согласуют плановый срок 35 календарных дней.",
        ]],
    )
    document.add_heading("Пример протокола разногласий", 2)
    document.add_paragraph(
        "В протоколе укажите пункт, исходный текст, предлагаемую редакцию "
        "и итоговое решение сторон."
    )

    document.add_page_break()
    document.add_heading("Часть 2. Задание для самостоятельной работы", 1)
    document.add_paragraph(
        f"Подготовьте сравнение двух редакций договора на разработку продукта "
        f"«{product_name}» в области «{product_area}» для пользователей «{users}». "
        f"Предмет договора: реализовать функцию «{feature}» по сценарию: {workflow}."
    )
    document.add_paragraph(
        f"Согласуйте бизнес-правило «{rule}» и критерий приёмки «{acceptance}», "
        "затем внесите решения в протокол разногласий."
    )
    _add_docx_table(
        document,
        ["Пункт договора", "Редакция заказчика", "Редакция исполнителя"],
        [
            [
                "Цена работ",
                f"{amount} рублей, аванс {advance_percent}%",
                f"{amount + 27_500} рублей, оплата по этапам",
            ],
            [
                "Срок исправления критической ошибки",
                f"{old_days} рабочих дней после обращения",
                f"{new_days} рабочих дней после подтверждения ошибки",
            ],
            [
                "Гарантийный период",
                f"{warranty_days} дней",
                f"{warranty_days + 30} дней",
            ],
            [
                "Приёмка результата",
                f"Проверка функции «{feature}» по сценарию: {workflow}",
                f"Приёмка по критерию: {acceptance}",
            ],
        ],
    )
    document.add_heading("Протокол разногласий", 2)
    _add_docx_table(
        document,
        ["№ пункта", "Предлагаемая редакция", "Решение / обоснование"],
        [
            ["1. Цена", "\n\n", "\n\n"],
            ["2. Срок", "\n\n", "\n\n"],
            ["3. Гарантия", "\n\n", "\n\n"],
            ["4. Приёмка", "\n\n", "\n\n"],
        ],
    )
    document.add_paragraph(
        "Сохраните файл Word. Необходимо заполнить все четыре строки протокола "
        "и кратко обосновать каждое предложенное решение."
    )
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def _style_workbook(workbook: Workbook) -> None:
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"


def _style_sheet(sheet, widths: dict[str, int] | None = None) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A4"
    if widths:
        for column, width in widths.items():
            sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows():
        for cell in row:
            cell.font = Font(name="Calibri", size=10, color=TEXT)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=THIN_BLUE)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 30


def _add_xlsx_instructions(sheet, number: int, variant: int, title: str, instruction: str) -> None:
    sheet.title = "Инструкция"
    sheet.append([f"ПРАКТИЧЕСКОЕ ЗАНЯТИЕ №{number} · ВАРИАНТ {variant:02d}"])
    sheet.append([title])
    sheet.append(["Студент: __________________________", "Группа: __________________"])
    sheet.append([instruction])
    sheet.freeze_panes = "A4"
    sheet.column_dimensions["A"].width = 38
    sheet.column_dimensions["B"].width = 36
    sheet.row_dimensions[1].height = 30
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE, size=14)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font = Font(color=TEXT, size=11)
    sheet["A2"].font = Font(bold=True, color=NAVY, size=16)
    sheet["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.row_dimensions[4].height = max(90, min(240, 30 + len(instruction) * 0.32))


def _add_table_sheet(workbook: Workbook, name: str, title: str, headers: list[str], rows: list[list]) :
    sheet = workbook.create_sheet(name)
    sheet.append([title])
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A3"
    for cell in sheet[2]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE)
    for row in sheet.iter_rows(min_row=3):
        for cell in row:
            cell.border = Border(bottom=THIN_BLUE)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    return sheet


def _add_example_sheet(workbook: Workbook, number: int, title: str, headings: list[str], rows: list[list]) -> None:
    sheet = workbook.create_sheet("Пример")
    sheet.append([f"ЗАНЯТИЕ №{number} · ПРИМЕР"])
    sheet.append([title])
    sheet.append(headings)
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A4"
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE, size=14)
    for cell in sheet[3]:
        cell.fill = PatternFill("solid", fgColor=TEAL)
        cell.font = Font(bold=True, color=WHITE)
    for column in range(1, len(headings) + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 22
    for row in sheet.iter_rows(min_row=4):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    return sheet


def _build_xlsx_5(variant: int) -> bytes:
    product, _users, _feature = _scenario(variant)
    workflow, rule, acceptance = _scenario_requirements(variant)
    rate = 1300 + variant * 35
    overhead = 0.12 + (variant % 5) * 0.02
    roles = [
        ("Аналитик", f"Сценарий: {workflow}", 18 + variant % 8, rate + 450),
        ("Разработчик", f"Реализация правила: {rule}", 46 + variant % 12, rate),
        ("Тестировщик", f"Приёмка: {acceptance}", 24 + variant % 7, rate - 250),
    ]
    estimated_cost = sum(hours * hourly_rate for _, _, hours, hourly_rate in roles)
    budget_limit = round(estimated_cost * (1 + overhead) * 1.08 / 1000) * 1000
    wb = Workbook()
    _style_workbook(wb)
    _add_xlsx_instructions(
        wb.active, 5, variant, "Расчёт бюджета и стоимости нормо-часа",
        f"Проект: {product}. Рассчитайте стоимость каждого результата работ "
        "на листе «Мой вариант». "
        "Стоимость этапа рассчитайте как трудозатраты, умноженные на ставку. "
        f"Добавьте накладные расходы {overhead:.0%} и проверьте, что общий бюджет "
        f"не превышает лимит {budget_limit:,} руб.".replace(",", " "),
    )
    _add_example_sheet(
        wb, 5, "Пример расчёта проекта «Сайт кафе»",
        ["Роль", "Часы", "Цена часа", "Стоимость"],
        [
            ["Аналитик", 10, 1600, "=B4*C4"],
            ["Разработчик", 24, 1400, "=B5*C5"],
            ["Тестировщик", 12, 1200, "=B6*C6"],
            ["Итого", "=SUM(B4:B6)", "", "=SUM(D4:D6)"],
            ["Накладные расходы", 0.15, "", "=D7*B8"],
            ["Общий бюджет", "", "", "=D7+D8"],
        ],
    )
    sheet = _add_table_sheet(
        wb,
        "Мой вариант",
        f"Вариант {variant:02d}: {product}",
        [
            "Роль",
            "Результат работ",
            "Трудозатраты, ч",
            "Цена нормо-часа, руб.",
            "Стоимость, руб.",
        ],
        [[role, deliverable, hours, hourly_rate, None]
         for role, deliverable, hours, hourly_rate in roles],
    )
    total_row = 3 + len(roles)
    sheet.cell(total_row, 1, "Итого")
    sheet.cell(total_row, 3, f"=SUM(C3:C{total_row - 1})")
    sheet.cell(total_row, 5, f"=SUM(E3:E{total_row - 1})")
    sheet.cell(total_row + 1, 1, "Накладные расходы")
    sheet.cell(total_row + 1, 2, overhead)
    sheet.cell(total_row + 1, 5, f"=E{total_row}*B{total_row + 1}")
    sheet.cell(total_row + 2, 1, "ИТОГОВЫЙ БЮДЖЕТ")
    sheet.cell(total_row + 2, 5, f"=E{total_row}+E{total_row + 1}")
    for row in range(3, total_row):
        sheet.cell(row, 5, f"=C{row}*D{row}")
    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 58
    sheet.column_dimensions["C"].width = 20
    sheet.column_dimensions["D"].width = 26
    sheet.column_dimensions["E"].width = 24
    for row in range(3, total_row):
        sheet.row_dimensions[row].height = 38
    for cell in sheet[2]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE)
    for row in range(3, total_row + 3):
        sheet.cell(row, 4).number_format = '#,##0.00 "₽"'
        sheet.cell(row, 5).number_format = '#,##0.00 "₽"'
    sheet.cell(total_row + 1, 2).number_format = "0%"
    return _save_workbook(wb)


def _raw_logs(variant: int) -> list[list]:
    systems = ["API", "Авторизация", "База данных", "Отчёты", "Интерфейс"]
    severities = ["INFO", "WARNING", "ERROR", "CRITICAL"]
    messages = [
        "тайм-аут запроса",
        "неверный пароль",
        "ошибка соединения",
        "сбой формирования отчёта",
        "недостаточно места",
    ]
    row_count = 18 + (variant - 1) % 7
    duplicate_index = 4 + (variant - 1) % 8
    malformed_date_index = 1 + variant % row_count
    rows = []
    for index in range(1, row_count + 1):
        event = (index + variant) % 4
        identifier = index - 1 if index == duplicate_index else index
        level = severities[(index + variant) % len(severities)]
        if index % 5 == 0:
            level = " CRITICAL "
        if index == malformed_date_index:
            timestamp = f"2026/03/{1 + (index + variant) % 27:02d}"
        else:
            timestamp = f"2026-03-{1 + (index + variant) % 27:02d} {8 + index % 10:02d}:{(index * 7) % 60:02d}"
        log_id = f"LOG-{variant:02d}-{identifier:03d}"
        if (index + variant) % 3 == 0:
            log_id = f" {log_id} "
        rows.append([
            log_id,
            timestamp,
            f" {level} ",
            systems[(index + variant) % len(systems)],
            f" {messages[(index + event) % len(messages)]} ",
        ])
    return rows


def _build_xlsx_6(variant: int) -> bytes:
    row_count = 18 + (variant - 1) % 7
    duplicate_index = 4 + (variant - 1) % 8
    malformed_date_index = 1 + variant % row_count
    wb = Workbook()
    _style_workbook(wb)
    _add_xlsx_instructions(
        wb.active, 6, variant, "Очистка и анализ журнала ошибок",
        f"На листе «Сырые данные» размещён импортированный журнал из {row_count} записей. "
        "Скопируйте данные на лист «Мой вариант», разделите поля по столбцам, "
        f"удалите повтор ID в записи №{duplicate_index}, очистите лишние пробелы, "
        f"исправьте формат даты в записи №{malformed_date_index}, установите корректные "
        "типы даты и условно выделите CRITICAL. На листе «Пример» показан образец очистки.",
    )
    _add_example_sheet(
        wb, 6, "Пример очистки записи лога",
        ["ID", "Дата и время", "Уровень", "Модуль", "Сообщение"],
        [[" LOG-01-004 ", "2026-03-04 10:15", " CRITICAL ", " API ", " timeout "],
         ["LOG-01-004", "04.03.2026 10:15", "CRITICAL", "API", "timeout"]],
    )
    sheet = _add_table_sheet(
        wb,
        "Сырые данные",
        f"Вариант {variant:02d} · исходный импорт",
        ["ID", "Временная метка", "Уровень", "Модуль", "Сообщение"],
        _raw_logs(variant),
    )
    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 24
    sheet.column_dimensions["C"].width = 18
    sheet.column_dimensions["D"].width = 22
    sheet.column_dimensions["E"].width = 34
    for row in range(3, sheet.max_row + 1):
        sheet.row_dimensions[row].height = 32
    result = _add_table_sheet(
        wb,
        "Мой вариант",
        f"Вариант {variant:02d} · очищенный журнал",
        ["ID", "Дата и время", "Уровень", "Модуль", "Сообщение"],
        [[None, None, None, None, None] for _ in range(row_count)],
    )
    for column, width in {"A": 22, "B": 24, "C": 18, "D": 22, "E": 34}.items():
        result.column_dimensions[column].width = width
    for row in range(3, result.max_row + 1):
        result.row_dimensions[row].height = 32
    result.conditional_formatting.add(
        "A3:E20",
        FormulaRule(
            formula=['$C3="CRITICAL"'],
            fill=PatternFill("solid", fgColor=RED),
            font=Font(color="9C0006", bold=True),
        ),
    )
    return _save_workbook(wb)


def _build_xlsx_7(variant: int) -> bytes:
    product, _users, _feature = _scenario(variant)
    workflow, rule, acceptance = _scenario_requirements(variant)
    start = date(2026, 3, 2) + timedelta(days=variant - 1)
    phases = [
        ("Сбор требований", 3 + variant % 3, "Аналитик"),
        ("Проектирование", 3 + variant % 4, "Архитектор"),
        ("Разработка интерфейса", 5 + variant % 4, "Разработчик"),
        (f"Реализация: {_feature}", 6 + variant % 5, "Разработчик"),
        (f"Проверка: {rule}", 4 + variant % 3, "Тестировщик"),
        ("Исправление замечаний", 3 + variant % 4, "Команда"),
        ("Подготовка релиза", 2 + variant % 2, "DevOps"),
    ]
    wb = Workbook()
    _style_workbook(wb)
    _add_xlsx_instructions(
        wb.active, 7, variant, "Диаграмма Ганта жизненного цикла ПО",
        f"Создайте план проекта «{product}» по сценарию «{workflow}». "
        f"Критерий готовности: {acceptance}. На листе «Мой вариант» заполните "
        "даты начала и окончания этапов, настройте условное форматирование "
        "по датам диаграммы и проследите зависимости между этапами. "
        "Образец готовой диаграммы находится на листе «Пример».",
    )
    _add_example_sheet(
        wb, 7, "Пример плана работ",
        ["Этап", "Начало", "Длительность, дн.", "Окончание"],
        [["Требования", date(2026, 3, 2), 3, "=B4+C4-1"],
         ["Проектирование", date(2026, 3, 5), 4, "=B5+C5-1"]],
    )
    example = wb["Пример"]
    example["B4"].number_format = "dd.mm.yyyy"
    example["D4"].number_format = "dd.mm.yyyy"
    example["B5"].number_format = "dd.mm.yyyy"
    example["D5"].number_format = "dd.mm.yyyy"

    sheet = wb.create_sheet("Мой вариант")
    sheet.append([f"Вариант {variant:02d} · проект «{product}»"])
    for day in range(21):
        sheet.cell(3, 8 + day, start + timedelta(days=day))
        sheet.cell(3, 8 + day).number_format = "dd.mm"
    headers = ["Этап", "Ответственный", "Начало", "Длительность, дн.", "Окончание", "Предшественник", "Статус"]
    for column, header in enumerate(headers, 1):
        sheet.cell(2, column, header)
    for row_index, (phase, duration, owner) in enumerate(phases, 4):
        sheet.row_dimensions[row_index].height = 42
        predecessor = phases[row_index - 5][0] if row_index > 4 else "—"
        sheet.cell(row_index, 1, phase)
        sheet.cell(row_index, 2, owner)
        sheet.cell(row_index, 3, start + timedelta(days=(row_index - 4) * 2))
        sheet.cell(row_index, 4, duration)
        sheet.cell(row_index, 5, f"=C{row_index}+D{row_index}-1")
        sheet.cell(row_index, 6, predecessor)
        sheet.cell(row_index, 7, "Не начато")
        sheet.cell(row_index, 3).number_format = "dd.mm.yyyy"
        sheet.cell(row_index, 5).number_format = "dd.mm.yyyy"
    for column in range(1, 8):
        sheet.cell(2, column).fill = PatternFill("solid", fgColor=NAVY)
        sheet.cell(2, column).font = Font(bold=True, color=WHITE)
    sheet["A3"] = "Календарь"
    for column in range(8, 29):
        sheet.cell(2, column, "План")
        sheet.cell(2, column).fill = PatternFill("solid", fgColor=NAVY)
        sheet.cell(2, column).font = Font(bold=True, color=WHITE)
    sheet.conditional_formatting.add(
        "H4:AB10",
        FormulaRule(
            formula=["AND(H$3>=$C4,H$3<=$E4)"],
            fill=PatternFill("solid", fgColor=TEAL),
        ),
    )
    sheet.freeze_panes = "H4"
    sheet.column_dimensions["A"].width = 31
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 15
    sheet.column_dimensions["D"].width = 17
    sheet.column_dimensions["E"].width = 15
    sheet.column_dimensions["F"].width = 29
    sheet.column_dimensions["G"].width = 16
    for column in range(8, 29):
        sheet.column_dimensions[get_column_letter(column)].width = 11
    return _save_workbook(wb)


def _bug_data(variant: int) -> list[list]:
    _product, _users, feature = _scenario(variant)
    _workflow, rule, _acceptance = _scenario_requirements(variant)
    statuses = ["Закрыта", "Закрыта", "В работе", "Закрыта", "Открыта"]
    priorities = ["Низкий", "Средний", "Высокий", "Критический"]
    sprints = [f"Спринт {(variant + index) % 4 + 1}" for index in range(4)]
    rows = []
    for index in range(1, 31):
        status = statuses[(index + variant) % len(statuses)]
        opened = date(2026, 2, 1) + timedelta(days=(index * 2 + variant) % 55)
        fix_days = (index + variant) % 11 + 1 if status == "Закрыта" else None
        developer = DEVELOPERS[(index + variant) % len(DEVELOPERS)][0]
        rows.append([
            f"BUG-{variant:02d}-{index:03d}",
            f"Сбой при операции «{feature}»: запись {index}; проверить правило: {rule}",
            developer,
            priorities[(index + variant) % len(priorities)],
            sprints[index % len(sprints)],
            status,
            opened,
            fix_days,
        ])
    return rows


def _build_xlsx_8(variant: int) -> bytes:
    _product, _users, feature = _scenario(variant)
    workflow, rule, acceptance = _scenario_requirements(variant)
    priorities = ["Низкий", "Средний", "Высокий", "Критический"]
    focus_priority = priorities[(variant - 1) % len(priorities)]
    focus_sprint = (variant - 1) % 4 + 1
    wb = Workbook()
    _style_workbook(wb)
    _add_xlsx_instructions(
        wb.active, 8, variant, "Сводный анализ выгрузки баг-трекера",
        "На листе «Данные» находится выгрузка задач. Создайте сводную таблицу: "
        "разработчик — строки, статус и приоритет — фильтры, количество ID — значения. "
        "Добавьте среднее время исправления по разработчикам. "
        f"Отдельно проанализируйте приоритет «{focus_priority}» в спринте {focus_sprint} "
        f"и сформулируйте вывод по операции «{feature}». Учтите правило: {rule}. "
        f"Критерий проверки вывода: {acceptance}. На листе «Пример» приведён образец.",
    )
    _add_example_sheet(
        wb, 8, "Пример анализа задач",
        ["Разработчик", "Закрыто задач", "Среднее исправление, дней"],
        [["Анна", 8, 4.5], ["Борис", 11, 3.8], ["Вера", 9, 2.9]],
    )
    data = _add_table_sheet(
        wb,
        "Данные",
        f"Вариант {variant:02d} · выгрузка баг-трекера",
        ["ID", "Описание", "Разработчик", "Приоритет", "Спринт", "Статус", "Дата открытия", "Дней до закрытия"],
        _bug_data(variant),
    )
    for column, width in {
        "A": 19, "B": 42, "C": 18, "D": 17, "E": 15, "F": 17, "G": 18, "H": 20,
    }.items():
        data.column_dimensions[column].width = width
    for row in range(3, data.max_row + 1):
        data.cell(row, 7).number_format = "dd.mm.yyyy"
        data.row_dimensions[row].height = 32
    table = Table(displayName=f"BugTrackerV{variant:02d}", ref=f"A2:H{data.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    data.add_table(table)
    return _save_workbook(wb)


def _build_xlsx_9(variant: int) -> bytes:
    project, _users, _feature = _scenario(variant)
    workflow, rule, acceptance = _scenario_requirements(variant)
    capacities = [50 + variant % 10, 80 + variant % 12, 55 + variant % 9]
    rates = [850 + variant * 4, 1400 + variant * 5, 2100 + variant * 7]
    tasks = [
        (f"Анализ сценария: {workflow}", 16 + variant % 8, "Junior"),
        (f"Экран операции: {_feature}", 34 + variant % 12, "Junior"),
        (f"Реализация правила: {rule}", 26 + variant % 10, "Middle"),
        ("Хранение и проверка данных", 18 + variant % 8, "Middle"),
        ("Интеграционное тестирование", 14 + variant % 7, "Senior"),
        (f"Проверка приёмки: {acceptance}", 18 + variant % 8, "Senior"),
    ]
    wb = Workbook()
    _style_workbook(wb)
    _add_xlsx_instructions(
        wb.active, 9, variant, "Оптимизация распределения задач",
        f"Для проекта «{project}» распределите работы по сценарию «{workflow}» "
        f"и правилу «{rule}». Заполните матрицу часов по квалификациям. "
        "Используйте надстройку Excel «Поиск решения»: минимизируйте общую стоимость, "
        "сумма распределённых часов по каждой задаче должна равняться её трудоёмкости, "
        "а суммарные часы по специалистам не должны превышать доступные мощности. "
        f"Не распределяйте задачу специалистам ниже минимального уровня. "
        f"Готовность подтвердите критерием: {acceptance}.",
    )
    _add_example_sheet(
        wb, 9, "Пример распределения задач",
        ["Задача", "Часы", "Junior, ч", "Middle, ч", "Senior, ч", "Стоимость"],
        [["Макет", 8, 8, 0, 0, "=C4*900+D4*1500+E4*2200"],
         ["Ревью", 6, 0, 2, 4, "=C5*900+D5*1500+E5*2200"],
         ["Итого", "=SUM(B4:B5)", "=SUM(C4:C5)", "=SUM(D4:D5)", "=SUM(E4:E5)", "=SUM(F4:F5)"]],
    )
    example = wb["Пример"]
    for cell in example[3]:
        cell.fill = PatternFill("solid", fgColor=TEAL)
        cell.font = Font(bold=True, color=WHITE)

    capacity = _add_table_sheet(
        wb,
        "Ресурсы",
        f"Вариант {variant:02d} · ресурсы проекта",
        ["Квалификация", "Доступно часов", "Стоимость часа, руб."],
        [["Junior", capacities[0], rates[0]], ["Middle", capacities[1], rates[1]], ["Senior", capacities[2], rates[2]]],
    )
    capacity.column_dimensions["A"].width = 24
    capacity.column_dimensions["B"].width = 22
    capacity.column_dimensions["C"].width = 27

    sheet = _add_table_sheet(
        wb,
        "Мой вариант",
        f"Вариант {variant:02d} · «{project}»",
        ["Задача", "Трудоёмкость, ч", "Мин. квалификация", "Junior, ч", "Middle, ч", "Senior, ч", "Распределено", "Стоимость, руб."],
        [[name, hours, level, 0, 0, 0, None, None] for name, hours, level in tasks],
    )
    for row in range(3, 3 + len(tasks)):
        sheet.row_dimensions[row].height = 48
        sheet.cell(row, 7, f"=SUM(D{row}:F{row})")
        sheet.cell(row, 8, f"=D{row}*Ресурсы!$C$3+E{row}*Ресурсы!$C$4+F{row}*Ресурсы!$C$5")
        minimum = sheet.cell(row, 3).value
        if minimum == "Middle":
            sheet.cell(row, 4).fill = PatternFill("solid", fgColor=RED)
        elif minimum == "Senior":
            sheet.cell(row, 4).fill = PatternFill("solid", fgColor=RED)
            sheet.cell(row, 5).fill = PatternFill("solid", fgColor=RED)
    total = 3 + len(tasks)
    sheet.cell(total, 1, "ИТОГО")
    for col in range(4, 9):
        letter = get_column_letter(col)
        sheet.cell(total, col, f"=SUM({letter}3:{letter}{total - 1})")
    sheet.append(["", "", "Максимум по квалификации:", "=Ресурсы!B3", "=Ресурсы!B4", "=Ресурсы!B5"])
    for col, width in {
        "A": 29, "B": 19, "C": 23, "D": 16, "E": 16, "F": 16, "G": 19, "H": 21,
    }.items():
        sheet.column_dimensions[col].width = width
    sheet.freeze_panes = "D3"
    sheet["A12"] = "Solver: целевая ячейка H9 — минимум; изменяемые D3:F8."
    sheet["A13"] = "Ограничения: G3:G8 = B3:B8; D9:F9 <= D10:F10; переменные >= 0."
    return _save_workbook(wb)


def _save_workbook(workbook: Workbook) -> bytes:
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_practice_file(
    practice_number: int, variant: int | None
) -> tuple[str, bytes]:
    if practice_number not in PRACTICE_NUMBERS:
        raise ValueError("Для этого занятия шаблон не предусмотрен.")
    if variant is None and practice_number != 2:
        raise ValueError("Для этого занятия необходимо указать номер варианта.")
    if variant is not None:
        _scenario(variant)
    if practice_number == 2:
        return ".docx", _build_word_2(variant)
    if variant is None:
        raise ValueError("Для этого занятия необходимо указать номер варианта.")
    if practice_number == 1:
        return ".docx", _build_word_1(variant)
    if practice_number == 3:
        return ".docx", _build_word_3(variant)
    builders = {
        5: _build_xlsx_5,
        6: _build_xlsx_6,
        7: _build_xlsx_7,
        8: _build_xlsx_8,
        9: _build_xlsx_9,
    }
    return ".xlsx", builders[practice_number](variant)


def build_practice_csv(variant: int) -> bytes:
    _scenario(variant)
    output = StringIO(newline="")
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["ID", "Временная метка", "Уровень", "Модуль", "Сообщение"])
    writer.writerows(_raw_logs(variant))
    return output.getvalue().encode("utf-8-sig")
