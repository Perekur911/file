# sync_launcher.py
from __future__ import annotations

import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext


PROJECT_DIR = Path(__file__).resolve().parent
SYNC_SCRIPT = PROJECT_DIR / "sync.py"
DEFAULT_DB_PATH = r"C:\sqlite\results.db"
DEFAULT_STUDIES = "OP AP GC SSF GOR BP REC EMV"


class SyncLauncher(tk.Tk):
    """Простое окно запуска sync.py с параметрами."""

    def __init__(self):
        super().__init__()

        self.title("Запуск синхронизации Excel ↔ SQLite")
        self.geometry("820x560")

        self.mode_var = tk.StringVar(value="save")
        self.workbook_var = tk.StringVar(value="")
        self.db_var = tk.StringVar(value=DEFAULT_DB_PATH)
        self.project_var = tk.StringVar(value="26-F083")
        self.studies_var = tk.StringVar(value=DEFAULT_STUDIES)

        self._build_ui()

    def _build_ui(self) -> None:
        row = 0

        tk.Label(self, text="Режим:").grid(row=row, column=0, sticky="w", padx=10, pady=6)

        mode_frame = tk.Frame(self)
        mode_frame.grid(row=row, column=1, sticky="w", padx=10, pady=6)

        tk.Radiobutton(
            mode_frame,
            text="save: Excel → SQLite",
            variable=self.mode_var,
            value="save",
            command=self._update_project_state,
        ).pack(side="left")

        tk.Radiobutton(
            mode_frame,
            text="load: SQLite → Excel",
            variable=self.mode_var,
            value="load",
            command=self._update_project_state,
        ).pack(side="left", padx=20)

        row += 1
        tk.Label(self, text="Excel-книга:").grid(row=row, column=0, sticky="w", padx=10, pady=6)
        tk.Entry(self, textvariable=self.workbook_var, width=85).grid(row=row, column=1, sticky="we", padx=10, pady=6)
        tk.Button(self, text="Выбрать", command=self._choose_workbook).grid(row=row, column=2, padx=10, pady=6)

        row += 1
        tk.Label(self, text="SQLite БД:").grid(row=row, column=0, sticky="w", padx=10, pady=6)
        tk.Entry(self, textvariable=self.db_var, width=85).grid(row=row, column=1, sticky="we", padx=10, pady=6)
        tk.Button(self, text="Выбрать", command=self._choose_db).grid(row=row, column=2, padx=10, pady=6)

        row += 1
        tk.Label(self, text="Проект для load:").grid(row=row, column=0, sticky="w", padx=10, pady=6)
        self.project_entry = tk.Entry(self, textvariable=self.project_var, width=30)
        self.project_entry.grid(row=row, column=1, sticky="w", padx=10, pady=6)

        row += 1
        tk.Label(self, text="Studies:").grid(row=row, column=0, sticky="w", padx=10, pady=6)
        tk.Entry(self, textvariable=self.studies_var, width=85).grid(row=row, column=1, sticky="we", padx=10, pady=6)

        row += 1
        btn_frame = tk.Frame(self)
        btn_frame.grid(row=row, column=1, sticky="w", padx=10, pady=10)

        tk.Button(btn_frame, text="Запустить", width=18, command=self._run_sync).pack(side="left")
        tk.Button(btn_frame, text="Очистить лог", width=18, command=self._clear_log).pack(side="left", padx=10)

        row += 1
        tk.Label(self, text="Команда / результат:").grid(row=row, column=0, sticky="nw", padx=10, pady=6)

        self.log = scrolledtext.ScrolledText(self, width=100, height=20)
        self.log.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=10, pady=6)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(row, weight=1)

        self._update_project_state()

    def _choose_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери открытую Excel-книгу",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xlsb"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.workbook_var.set(path)

    def _choose_db(self) -> None:
        path = filedialog.askopenfilename(
            title="Выбери SQLite БД",
            filetypes=[
                ("SQLite database", "*.db *.sqlite *.sqlite3"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.db_var.set(path)

    def _update_project_state(self) -> None:
        if self.mode_var.get() == "load":
            self.project_entry.configure(state="normal")
        else:
            self.project_entry.configure(state="disabled")

    def _run_sync(self) -> None:
        mode = self.mode_var.get().strip()
        workbook = self.workbook_var.get().strip()
        db_path = self.db_var.get().strip()
        project = self.project_var.get().strip()
        studies = self.studies_var.get().strip()

        if not workbook:
            messagebox.showerror("Ошибка", "Укажи путь к Excel-книге")
            return

        if not db_path:
            messagebox.showerror("Ошибка", "Укажи путь к SQLite БД")
            return

        if mode == "load" and not project:
            messagebox.showerror("Ошибка", "Для режима load нужен номер проекта")
            return

        if not SYNC_SCRIPT.exists():
            messagebox.showerror("Ошибка", f"Не найден sync.py:\n{SYNC_SCRIPT}")
            return

        cmd = [
            sys.executable,
            str(SYNC_SCRIPT),
            "--mode",
            mode,
            "--workbook",
            workbook,
            "--db",
            db_path,
        ]

        if studies:
            cmd.append("--studies")
            cmd.extend(studies.split())

        if mode == "load":
            cmd.extend(["--project", project])

        self._log("\n" + "=" * 80)
        self._log("Команда:")
        self._log(" ".join(f'"{x}"' if " " in x else x for x in cmd))
        self._log("=" * 80)

        try:
            completed = subprocess.run(
                cmd,
                cwd=str(PROJECT_DIR),
                text=True,
                capture_output=True,
                encoding="utf-8",
                errors="replace",
            )
        except Exception as exc:
            messagebox.showerror("Ошибка запуска", str(exc))
            return

        if completed.stdout:
            self._log("\nSTDOUT:")
            self._log(completed.stdout)

        if completed.stderr:
            self._log("\nSTDERR:")
            self._log(completed.stderr)

        self._log(f"\nExit code: {completed.returncode}")

        if completed.returncode == 0:
            messagebox.showinfo("Готово", "sync.py завершился успешно")
        else:
            messagebox.showerror("Ошибка", "sync.py завершился с ошибкой. Смотри лог.")

    def _log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def _clear_log(self) -> None:
        self.log.delete("1.0", "end")


if __name__ == "__main__":
    app = SyncLauncher()
    app.mainloop()