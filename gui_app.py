import tkinter as tk
from tkinter import messagebox, ttk
from typing import Dict, List

from ticket_service import (
    Student,
    Ticket,
    append_result_row,
    assign_ticket,
    load_students,
    parse_tickets,
)

STUDENTS_FILE = "students.xlsx"
TICKETS_FILE = "tickets.docx"
RESULTS_FILE = "results.xlsx"


class TicketResultDialog(tk.Toplevel):
    """Модальное окно с результатом генерации билета (закрывается по ESC или кнопке)."""

    def __init__(
        self,
        parent: tk.Tk,
        group: str,
        student: Student,
        ticket: Ticket,
        is_repeat: bool,
    ) -> None:
        super().__init__(parent)
        self.title(f"Билет № {ticket.number}")
        self.geometry("560x360")
        self.minsize(480, 300)
        self.transient(parent)
        self.grab_set()

        # Привязываем ESC к закрытию только этого окна результата
        self.bind("<Escape>", lambda event: self.destroy())

        container = ttk.Frame(self, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        header_text = f"Ваш билет № {ticket.number}"
        header_label = ttk.Label(
            container,
            text=header_text,
            font=("Segoe UI", 16, "bold"),
        )
        header_label.pack(anchor=tk.W, pady=(0, 6))

        repeat_badge = "Повторная выдача (да)" if is_repeat else "Первичная выдача (нет)"
        sub_label = ttk.Label(
            container,
            text=f"Студент: {student.full_name}  |  Группа: {group}  |  {repeat_badge}",
            font=("Segoe UI", 10),
            foreground="#555555",
        )
        sub_label.pack(anchor=tk.W, pady=(0, 14))

        ttk.Separator(container, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 12))

        questions_frame = ttk.Frame(container)
        questions_frame.pack(fill=tk.BOTH, expand=True)

        for q in ticket.questions:
            q_label = ttk.Label(
                questions_frame,
                text=q,
                font=("Segoe UI", 11),
                wraplength=510,
                justify=tk.LEFT,
            )
            q_label.pack(anchor=tk.W, pady=6)

        footer_frame = ttk.Frame(container)
        footer_frame.pack(fill=tk.X, pady=(12, 0))

        hint_label = ttk.Label(
            footer_frame,
            text="Нажмите ESC для закрытия окна",
            font=("Segoe UI", 9, "italic"),
            foreground="#666666",
        )
        hint_label.pack(side=tk.LEFT)

        close_btn = ttk.Button(footer_frame, text="Закрыть (ESC)", command=self.destroy)
        close_btn.pack(side=tk.RIGHT)

        self.focus_set()


class TicketGeneratorApp(tk.Tk):
    """Главное окно десктопного приложения генератора экзаменационных билетов."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Генератор экзаменационных билетов")
        self.geometry("480x270")
        self.resizable(False, False)

        self.groups_data: Dict[str, List[Student]] = {}
        self.tickets_data: Dict[int, Ticket] = {}
        self.current_students: List[Student] = []

        self._build_ui()
        self._load_sources()

    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self, padding=24)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_lbl = ttk.Label(
            main_frame,
            text="Выдача экзаменационного билета",
            font=("Segoe UI", 13, "bold"),
        )
        title_lbl.pack(anchor=tk.W, pady=(0, 16))

        # Выбор группы
        group_lbl = ttk.Label(main_frame, text="Выберите группу:", font=("Segoe UI", 10))
        group_lbl.pack(anchor=tk.W)

        self.group_var = tk.StringVar()
        self.group_combo = ttk.Combobox(
            main_frame,
            textvariable=self.group_var,
            state="readonly",
            font=("Segoe UI", 10),
        )
        self.group_combo.pack(fill=tk.X, pady=(4, 14))
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_selected)

        # Выбор студента
        student_lbl = ttk.Label(
            main_frame, text="Выберите студента:", font=("Segoe UI", 10)
        )
        student_lbl.pack(anchor=tk.W)

        self.student_var = tk.StringVar()
        self.student_combo = ttk.Combobox(
            main_frame,
            textvariable=self.student_var,
            state="disabled",
            font=("Segoe UI", 10),
        )
        self.student_combo.pack(fill=tk.X, pady=(4, 20))
        self.student_combo.bind("<<ComboboxSelected>>", self._on_student_selected)

        # Кнопка генерации
        self.generate_btn = ttk.Button(
            main_frame,
            text="Сгенерировать билет",
            state="disabled",
            command=self._on_generate_clicked,
        )
        self.generate_btn.pack(fill=tk.X, ipady=6)

    def _load_sources(self) -> None:
        """Загружает и валидирует входные файлы students.xlsx и tickets.docx."""
        errors: List[str] = []

        try:
            self.groups_data = load_students(STUDENTS_FILE)
        except FileNotFoundError as e:
            errors.append(str(e))
        except Exception as e:
            errors.append(f"Ошибка чтения '{STUDENTS_FILE}': {e}")

        try:
            self.tickets_data = parse_tickets(TICKETS_FILE)
            if not self.tickets_data:
                errors.append(
                    f"В файле '{TICKETS_FILE}' не найдено ни одного корректного билета (минимум 3 вопроса)."
                )
        except FileNotFoundError as e:
            errors.append(str(e))
        except Exception as e:
            errors.append(f"Ошибка чтения '{TICKETS_FILE}': {e}")

        if errors:
            messagebox.showerror(
                "Ошибка загрузки данных",
                "Обнаружены проблемы с входными файлами:\n\n"
                + "\n".join(f"• {err}" for err in errors)
                + "\n\nПроверьте наличие и корректность файлов в папке с приложением.",
            )
            self.group_combo.configure(state="disabled")
            self.student_combo.configure(state="disabled")
            self.generate_btn.configure(state="disabled")
            return

        group_names = list(self.groups_data.keys())
        self.group_combo["values"] = group_names
        if not group_names:
            self.group_combo.configure(state="disabled")

    def _on_group_selected(self, event=None) -> None:
        selected_group = self.group_var.get()
        self.current_students = self.groups_data.get(selected_group, [])

        # Сбрасываем текущий выбор студента и деактивируем кнопку
        self.student_var.set("")
        self.generate_btn.configure(state="disabled")

        if self.current_students:
            student_names = [s.full_name for s in self.current_students]
            self.student_combo["values"] = student_names
            self.student_combo.configure(state="readonly")
        else:
            # Пустая группа: второй список пуст, кнопка неактивна, без исключений
            self.student_combo["values"] = []
            self.student_combo.configure(state="disabled")

    def _on_student_selected(self, event=None) -> None:
        if self.group_var.get() and self.student_combo.current() >= 0:
            if self.tickets_data:
                self.generate_btn.configure(state="normal")
        else:
            self.generate_btn.configure(state="disabled")

    def _on_generate_clicked(self) -> None:
        group = self.group_var.get()
        student_idx = self.student_combo.current()
        if not group or student_idx < 0 or student_idx >= len(self.current_students):
            return

        student = self.current_students[student_idx]

        # Определяем билет с учётом истории в results.xlsx (с повтором при блокировке чтения)
        while True:
            try:
                ticket, is_repeat = assign_ticket(
                    group=group,
                    last_name=student.last_name,
                    first_name=student.first_name,
                    tickets=self.tickets_data,
                    results_filepath=RESULTS_FILE,
                )
                break
            except PermissionError:
                retry = messagebox.askretrycancel(
                    "Файл заблокирован",
                    f"Файл '{RESULTS_FILE}' сейчас открыт в другой программе (например, в Excel).\n\n"
                    "Пожалуйста, закройте его и нажмите «Повтор».",
                )
                if not retry:
                    return
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось назначить билет: {e}")
                return

        # Записываем и сразу сохраняем результат в results.xlsx (с повтором при блокировке записи)
        while True:
            try:
                append_result_row(
                    group=group,
                    last_name=student.last_name,
                    first_name=student.first_name,
                    ticket_number=ticket.number,
                    is_repeat=is_repeat,
                    results_filepath=RESULTS_FILE,
                )
                break
            except PermissionError:
                retry = messagebox.askretrycancel(
                    "Файл заблокирован",
                    f"Не удалось сохранить результат: файл '{RESULTS_FILE}' открыт в Excel.\n\n"
                    "Закройте файл в Excel и нажмите «Повтор», чтобы сохранить запись.",
                )
                if not retry:
                    return
            except Exception as e:
                messagebox.showerror(
                    "Ошибка записи", f"Не удалось записать результат в '{RESULTS_FILE}': {e}"
                )
                return

        # Открываем окно результата (закрывается по ESC, возвращая к выбору)
        TicketResultDialog(
            parent=self,
            group=group,
            student=student,
            ticket=ticket,
            is_repeat=is_repeat,
        )


def main() -> None:
    app = TicketGeneratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
