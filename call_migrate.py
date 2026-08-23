from __future__ import annotations

import os
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext

# =============================================================================
# Пути
# =============================================================================

PROJECT_DIR = Path(__file__).resolve().parent

MIGRATE_SCRIPT = (
    PROJECT_DIR
    / "migrate_v21_to_v22.py"
)

# Миграция всегда запускается Python из проектного .venv,
# независимо от того, каким интерпретатором запущен сам launcher.
MIGRATE_PYTHON = Path(
    r"C:\Users\Dmitriy.Kayurin\.venvs\sqlite-excel\Scripts\python.exe"
)


# =============================================================================
# Окно запуска
# =============================================================================

class MigrateLauncher(tk.Tk):

    def __init__(self) -> None:
        super().__init__()

        self.title("Миграция форм v21 → v22")
        self.geometry("900x620")
        self.minsize(720, 480)

        self.project_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=False)

        self.process: subprocess.Popen[str] | None = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        main_frame = tk.Frame(
            self,
            padx=12,
            pady=12,
        )
        main_frame.pack(
            fill="both",
            expand=True,
        )

        tk.Label(
            main_frame,
            text="Проект или диапазон проектов:",
            anchor="w",
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 5),
        )

        self.project_entry = tk.Entry(
            main_frame,
            textvariable=self.project_var,
        )
        self.project_entry.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 8),
        )

        tk.Label(
            main_frame,
            text=(
                "Примеры: 26-F501, 26-F501...26-F510\n"
                "Несколько проектов можно разделить пробелом, "
                "запятой или точкой с запятой."
            ),
            justify="left",
            anchor="w",
        ).grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 12),
        )

        tk.Checkbutton(
            main_frame,
            text="Перезаписать уже существующие файлы v22",
            variable=self.overwrite_var,
        ).grid(
            row=3,
            column=0,
            sticky="w",
            pady=(0, 12),
        )

        button_frame = tk.Frame(main_frame)
        button_frame.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(0, 10),
        )

        self.run_button = tk.Button(
            button_frame,
            text="Запустить миграцию",
            width=22,
            command=self._start_migration,
        )
        self.run_button.pack(side="left")

        self.clear_button = tk.Button(
            button_frame,
            text="Очистить лог",
            width=15,
            command=self._clear_log,
        )
        self.clear_button.pack(
            side="left",
            padx=(10, 0),
        )

        self.status_label = tk.Label(
            button_frame,
            text="Готово к запуску",
            anchor="w",
        )
        self.status_label.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(15, 0),
        )

        self.log = scrolledtext.ScrolledText(
            main_frame,
            wrap="word",
            height=25,
            font=("Consolas", 10),
        )
        self.log.grid(
            row=5,
            column=0,
            sticky="nsew",
        )

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        self.project_entry.focus_set()

        # Enter запускает миграцию.
        self.bind(
            "<Return>",
            lambda _event: self._start_migration(),
        )

    # =========================================================================
    # Подготовка запуска
    # =========================================================================

    def _start_migration(self) -> None:
        if self.process is not None:
            messagebox.showwarning(
                "Миграция уже выполняется",
                "Дождись завершения текущей миграции.",
            )
            return

        projects = self.project_var.get().strip()

        if not projects:
            messagebox.showerror(
                "Не указан проект",
                "Укажи проект, несколько проектов или диапазон проектов.",
            )
            self.project_entry.focus_set()
            return

        if not MIGRATE_SCRIPT.is_file():
            messagebox.showerror(
                "Скрипт не найден",
                "Не найден скрипт миграции:\n\n"
                f"{MIGRATE_SCRIPT}",
            )
            return

        if not MIGRATE_PYTHON.is_file():
            messagebox.showerror(
                "Python не найден",
                "Не найден интерпретатор проектного "
                "виртуального окружения:\n\n"
                f"{MIGRATE_PYTHON}\n\n"
                "Проверь наличие файла:\n"
                ".venv\\Scripts\\python.exe",
            )
            return

        command = [
            str(MIGRATE_PYTHON),

            # Не буферизовать stdout/stderr, чтобы строки
            # появлялись в окне сразу.
            "-u",

            str(MIGRATE_SCRIPT),
            "--projects",
            projects,
        ]

        if self.overwrite_var.get():
            command.append("--overwrite")

        self._append_log("=" * 100)
        self._append_log("Запуск миграции")
        self._append_log(f"Проекты: {projects}")
        self._append_log(
            "Перезапись: "
            + (
                "включена"
                if self.overwrite_var.get()
                else "выключена"
            )
        )
        self._append_log(
            f"Интерпретатор: {MIGRATE_PYTHON}"
        )
        self._append_log(
            f"Скрипт: {MIGRATE_SCRIPT}"
        )
        self._append_log(
            "Команда: "
            + self._format_command(command)
        )
        self._append_log("=" * 100)

        self.run_button.configure(state="disabled")
        self.project_entry.configure(state="disabled")

        self.status_label.configure(
            text="Миграция выполняется..."
        )

        worker = threading.Thread(
            target=self._run_process,
            args=(command,),
            daemon=True,
        )
        worker.start()

    # =========================================================================
    # Выполнение процесса
    # =========================================================================

    def _run_process(
        self,
        command: list[str],
    ) -> None:
        try:
            env = os.environ.copy()

            # Принудительно включаем UTF-8 в дочернем Python.
            env["PYTHONUTF8"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUNBUFFERED"] = "1"

            creationflags = 0

            # Не открываем отдельное консольное окно.
            if os.name == "nt":
                creationflags = subprocess.CREATE_NO_WINDOW

            self.process = subprocess.Popen(
                command,
                cwd=str(PROJECT_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
                creationflags=creationflags,
            )

            if self.process.stdout is not None:
                for line in self.process.stdout:
                    self.after(
                        0,
                        self._append_log,
                        line.rstrip("\r\n"),
                    )

            return_code = self.process.wait()

            self.after(
                0,
                self._migration_finished,
                return_code,
            )

        except Exception as exc:
            self.after(
                0,
                self._migration_failed,
                str(exc),
            )

        finally:
            self.process = None

    # =========================================================================
    # Завершение
    # =========================================================================

    def _migration_finished(
        self,
        return_code: int,
    ) -> None:
        self.run_button.configure(state="normal")
        self.project_entry.configure(state="normal")

        if return_code == 0:
            self.status_label.configure(
                text="Миграция завершена успешно"
            )

            self._append_log("")
            self._append_log(
                "Миграция завершена успешно."
            )

            messagebox.showinfo(
                "Готово",
                "Миграция завершена успешно.",
            )

        else:
            self.status_label.configure(
                text=f"Ошибка, код завершения: {return_code}"
            )

            self._append_log("")
            self._append_log(
                "Миграция завершилась с ошибкой. "
                f"Код: {return_code}"
            )

            messagebox.showerror(
                "Ошибка миграции",
                "Миграция завершилась с ошибкой.\n\n"
                f"Код завершения: {return_code}\n\n"
                "Подробности показаны в журнале окна.",
            )

        self.project_entry.focus_set()

    def _migration_failed(
        self,
        error_message: str,
    ) -> None:
        self.run_button.configure(state="normal")
        self.project_entry.configure(state="normal")

        self.status_label.configure(
            text="Не удалось запустить миграцию"
        )

        self._append_log("")
        self._append_log(
            "Не удалось запустить migrate_v21_to_v22.py:"
        )
        self._append_log(error_message)

        messagebox.showerror(
            "Ошибка запуска",
            error_message,
        )

        self.project_entry.focus_set()

    # =========================================================================
    # Вспомогательные методы
    # =========================================================================

    def _append_log(
        self,
        text: str,
    ) -> None:
        self.log.insert(
            "end",
            text + "\n",
        )
        self.log.see("end")

    def _clear_log(self) -> None:
        self.log.delete(
            "1.0",
            "end",
        )

    @staticmethod
    def _format_command(
        command: list[str],
    ) -> str:
        return " ".join(
            f'"{part}"'
            if any(char.isspace() for char in part)
            else part
            for part in command
        )

    def _on_close(self) -> None:
        if self.process is not None:
            close_confirmed = messagebox.askyesno(
                "Миграция выполняется",
                "Миграция ещё выполняется.\n\n"
                "Закрыть окно и прервать процесс?",
            )

            if not close_confirmed:
                return

            try:
                self.process.terminate()
            except Exception:
                pass

        self.destroy()


if __name__ == "__main__":
    app = MigrateLauncher()
    app.mainloop()
