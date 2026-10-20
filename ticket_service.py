import logging
import os
import random
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from docx import Document
from openpyxl import Workbook, load_workbook

logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

RESULTS_HEADERS = [
    "Группа",
    "Фамилия",
    "Имя",
    "Номер билета",
    "Дата и время",
    "Повтор",
]


@dataclass
class Student:
    last_name: str
    first_name: str

    @property
    def full_name(self) -> str:
        return f"{self.last_name} {self.first_name}"


@dataclass
class Ticket:
    number: int
    questions: List[str]


def load_students(filepath: str = "students.xlsx") -> Dict[str, List[Student]]:
    """
    Читает students.xlsx.
    Каждый лист — группа (имя листа = имя группы).
    Строка 1 — заголовок, со строки 2: колонка A — Фамилия, колонка B — Имя.
    Возвращает словарь: {имя_группы: [Student, ...]}.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Файл со списком студентов '{filepath}' не найден.")

    wb = load_workbook(filepath, read_only=True, data_only=True)
    groups: Dict[str, List[Student]] = {}

    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            students: List[Student] = []
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            for row in rows:
                if not row or len(row) < 2:
                    continue
                last_name = str(row[0]).strip() if row[0] is not None else ""
                first_name = str(row[1]).strip() if row[1] is not None else ""
                if last_name and first_name:
                    students.append(Student(last_name=last_name, first_name=first_name))
            groups[sheet_name] = students
    finally:
        wb.close()

    return groups


def parse_tickets(filepath: str = "tickets.docx") -> Dict[int, Ticket]:
    """
    Парсит файл tickets.docx.
    Заголовок билета: 'Билет N'
    Вопросы: нумерованные абзацы '1. Текст вопроса' (или абзацы со стилем списка).
    Если у билета меньше 3 вопросов или заголовок некорректен — билет пропускается,
    ошибка логируется, работа продолжается с остальными билетами.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Файл с билетами '{filepath}' не найден.")

    doc = Document(filepath)
    tickets: Dict[int, Ticket] = {}

    header_pattern = re.compile(r"^Билет\s+(\d+)\s*$", re.IGNORECASE)
    question_pattern = re.compile(r"^(\d+)\.\s+(.+)$")

    current_number: Optional[int] = None
    current_questions: List[str] = []

    def finalize_ticket(num: Optional[int], questions: List[str]) -> None:
        if num is None:
            return
        if len(questions) < 3:
            logger.warning(
                f"Билет {num} пропущен: найдено вопросов ({len(questions)}), требуется минимум 3."
            )
            return
        tickets[num] = Ticket(number=num, questions=questions[:3])

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Проверяем заголовок вида "Билет N"
        header_match = header_pattern.match(text)
        if header_match:
            finalize_ticket(current_number, current_questions)
            current_number = int(header_match.group(1))
            current_questions = []
            continue

        # Если строка начинается со слова "Билет", но не подходит под формат "Билет N"
        if text.lower().startswith("билет"):
            finalize_ticket(current_number, current_questions)
            logger.warning(f"Некорректный заголовок билета пропущен: '{text}'")
            current_number = None
            current_questions = []
            continue

        # Если мы находимся внутри валидного заголовка билета, проверяем вопросы
        if current_number is not None:
            q_match = question_pattern.match(text)
            if q_match:
                current_questions.append(f"{q_match.group(1)}. {q_match.group(2).strip()}")
            else:
                logger.warning(
                    f"Нераспознанная строка внутри Билета {current_number}: '{text}'"
                )
        else:
            logger.warning(f"Строка вне контекста билета пропущена: '{text}'")

    finalize_ticket(current_number, current_questions)
    return tickets


def find_existing_ticket(
    group: str,
    last_name: str,
    first_name: str,
    results_filepath: str = "results.xlsx",
) -> Optional[int]:
    """
    Ищет самую первую запись студента (группа + Фамилия + Имя) в results.xlsx.
    Если найдена — возвращает номер билета (int), иначе None.
    """
    if not os.path.exists(results_filepath):
        return None

    wb = load_workbook(results_filepath, read_only=True, data_only=True)
    try:
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 4:
                continue
            r_group = str(row[0]).strip() if row[0] is not None else ""
            r_last = str(row[1]).strip() if row[1] is not None else ""
            r_first = str(row[2]).strip() if row[2] is not None else ""
            r_ticket = row[3]

            if (
                r_group == group
                and r_last.lower() == last_name.lower()
                and r_first.lower() == first_name.lower()
            ):
                try:
                    return int(r_ticket)
                except (ValueError, TypeError):
                    continue
    finally:
        wb.close()

    return None


def append_result_row(
    group: str,
    last_name: str,
    first_name: str,
    ticket_number: int,
    is_repeat: bool,
    results_filepath: str = "results.xlsx",
    timestamp: Optional[str] = None,
) -> None:
    """
    Дописывает строку в results.xlsx и сразу сохраняет файл.
    Если файла нет — создаёт его с шапкой.
    Может выбросить PermissionError, если файл открыт в Excel.
    """
    if timestamp is None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    repeat_str = "да" if is_repeat else "нет"

    if os.path.exists(results_filepath):
        wb = load_workbook(results_filepath)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(RESULTS_HEADERS)

    ws.append([group, last_name, first_name, ticket_number, timestamp, repeat_str])
    wb.save(results_filepath)
    wb.close()


def assign_ticket(
    group: str,
    last_name: str,
    first_name: str,
    tickets: Dict[int, Ticket],
    results_filepath: str = "results.xlsx",
) -> Tuple[Ticket, bool]:
    """
    Определяет билет для студента (без записи в файл, чтобы запись можно было повторить при блокировке Excel):
    - Если студент уже есть в results.xlsx, возвращает его первый билет и флаг is_repeat=True.
    - Если студента нет, случайно выбирает один из доступных билетов и возвращает is_repeat=False.
    """
    if not tickets:
        raise ValueError("Список доступных билетов пуст.")

    existing_num = find_existing_ticket(group, last_name, first_name, results_filepath)
    if existing_num is not None and existing_num in tickets:
        return tickets[existing_num], True
    elif existing_num is not None:
        # Если в журнале был билет, который сейчас отсутствует в docx, создаём объект с заглушкой вопросов
        fallback_ticket = Ticket(
            number=existing_num,
            questions=[
                "1. (Вопрос не найден в текущем tickets.docx)",
                "2. (Вопрос не найден в текущем tickets.docx)",
                "3. (Вопрос не найден в текущем tickets.docx)",
            ],
        )
        return fallback_ticket, True

    chosen_num = random.choice(list(tickets.keys()))
    return tickets[chosen_num], False
