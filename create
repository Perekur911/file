from __future__ import annotations

import argparse
import json
import logging
import re
import shutil
import sqlite3
import sys
import traceback
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import openpyxl
import xlwings as xw


# =============================================================================
# Настройки проекта
# =============================================================================

PROJECT_DIR = Path(__file__).resolve().parent

DB_PATH = r"C:\sqlite\results.db"

# Корневая папка, внутри которой создавать папки проектов.
# Если USE_YEAR_SUBFOLDER=True, итог будет:
#   BASE_PROJECTS_DIR\2026\26-F001-XXX-YYY\...
BASE_PROJECTS_DIR = r"L:\LRC\common_data\ФЛЮИДЫ\ГТИ\Работа"

USE_YEAR_SUBFOLDER = True

# Чистая форма v22
CLEAN_V22_TEMPLATE = r"L:\LRC\common_data\ФЛЮИДЫ\ПТИ\sqlite-excel\clean_form_v22.xlsx"

# Внешняя книга с заданиями, та же, которую используешь для refresh_tasks
TASKS_WORKBOOK = r"L:\LRC\exchange\КСП Лайт\Журнал_заданий_флюиды.xlsx"

# Если во внешнем файле задания лежат в умной таблице:
TASKS_PARENT_SHEET = "task"
TASKS_PARENT_TABLE = "task"

# Если таблицы нет, скрипт попробует прочитать просто used range листа TASKS_PARENT_SHEET.
TASK_SAMPLE_CODE_COL = "Код проекта"
TASK_DATETIME_COL = "Дата и время"

# refresh_tasks.py должен лежать рядом или быть доступен через PYTHONPATH.
REFRESH_TASKS_SOURCE_PARENT_SHEET = "task"
REFRESH_TASKS_SOURCE_PARENT_TABLE = "task"
REFRESH_TASKS_SOURCE_CHILD_SHEET = "task_mix"
REFRESH_TASKS_SOURCE_CHILD_TABLE = "task_mix"

REFRESH_TASKS_TARGET_PARENT_SHEET = "Task"
REFRESH_TASKS_TARGET_PARENT_TABLE = "Task"
REFRESH_TASKS_TARGET_CHILD_SHEET = "Task_mix"
REFRESH_TASKS_TARGET_CHILD_TABLE = "Task_mix"

# Надстройка и макрос
ADDIN_PATH = r"L:\LRC\common_data\ФЛЮИДЫ\ПТИ\sqlite-excel\надстройка.xlam"
AFTER_REFRESH_MACRO = "wrappers.QueryFilterSilent_wrap"

# Куда писать номер проекта в новой форме
PROJECT_CELL_SHEET = "OP"
PROJECT_CELL_ADDRESS = "B6"

# Сколько последних проектов брать из внешнего журнала, если пользователь оставил ввод пустым
RECENT_PROJECT_LIMIT = 50

# Как выделять имя папки проекта из sampleCode:
# 25-F218-SRU-204-GS1 -> при 4 будет 25-F218-SRU-204
# 25-F218-SRU-204-GS1 -> при 2 будет 25-F218
FOLDER_PROJECT_PARTS = 4

# Какой проект передавать в refresh_tasks:
# обычно 25-F218, 26-F001 и т.д.
REFRESH_PROJECT_PARTS = 2

LOG_PATH = PROJECT_DIR / "create_project_forms.log"


logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filemode="a",
    encoding="utf-8",
)


# =============================================================================
# Reports
# =============================================================================

@dataclass
class ProjectCandidate:
    folder_project: str
    refresh_project: str
    latest_datetime: str | None
    example_sample_code: str


@dataclass
class ProjectCreateReport:
    project: str
    ok: bool
    message: str
    folder: str | None = None
    workbook: str | None = None
    refresh_project: str | None = None
    error_type: str | None = None


@dataclass
class CreateFormsReport:
    ok: bool
    message: str
    candidates_found: int
    already_started: int
    created: int
    skipped_existing_files: int
    projects: list[ProjectCreateReport]


# =============================================================================
# Console progress
# =============================================================================

def progress(message: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}", flush=True)


# =============================================================================
# Project parsing
# =============================================================================

def normalize_project_number(value: str) -> str:
    value = value.strip().upper()

    match = re.fullmatch(r"(\d{2})-F(\d+)", value)

    if not match:
        raise ValueError(f"Некорректный номер проекта: {value}")

    year, number = match.groups()
    return f"{year}-F{int(number):03d}"


def expand_project_range(start: str, end: str) -> list[str]:
    start = normalize_project_number(start)
    end = normalize_project_number(end)

    m1 = re.fullmatch(r"(\d{2}-F)(\d{3})", start)
    m2 = re.fullmatch(r"(\d{2}-F)(\d{3})", end)

    if not m1 or not m2:
        raise ValueError(f"Некорректный диапазон: {start}...{end}")

    prefix1, n1 = m1.groups()
    prefix2, n2 = m2.groups()

    if prefix1 != prefix2:
        raise ValueError(f"Диапазон должен быть внутри одного года: {start}...{end}")

    n1_i = int(n1)
    n2_i = int(n2)

    if n2_i < n1_i:
        raise ValueError(f"Конец диапазона меньше начала: {start}...{end}")

    return [f"{prefix1}{i:03d}" for i in range(n1_i, n2_i + 1)]


def parse_projects_input(text: str) -> list[str]:
    text = text.strip()

    if not text:
        return []

    text = re.sub(r"\s*\.{2,3}\s*", "...", text)

    parts = re.split(r"[,\s;]+", text)

    result: list[str] = []

    for part in parts:
        part = part.strip()

        if not part:
            continue

        if "..." in part:
            start, end = part.split("...", 1)
            result.extend(expand_project_range(start, end))
        else:
            result.append(normalize_project_number(part))

    seen = set()
    unique: list[str] = []

    for project in result:
        if project not in seen:
            unique.append(project)
            seen.add(project)

    return unique


def ask_projects_text() -> str:
    try:
        import tkinter as tk
        from tkinter import simpledialog

        root = tk.Tk()
        root.withdraw()

        value = simpledialog.askstring(
            title="Создание форм проектов",
            prompt=(
                "Введите проект или диапазон, например:\n"
                "26-F001 или 26-F001...26-F015\n\n"
                "Оставьте пустым, чтобы взять последние 50 проектов из журнала заданий."
            ),
        )

        root.destroy()

        return "" if value is None else value.strip()

    except Exception:
        return input(
            "Введите проект/диапазон или оставьте пустым для последних 50 проектов: "
        ).strip()


# =============================================================================
# Sample/project helpers
# =============================================================================

def canonicalize_f_part(value: str) -> str:
    """
    F218, F218a, F218b -> F218
    """
    text = str(value).strip().upper()

    match = re.fullmatch(r"F(\d{3})([A-ZА-Я]*)?", text)

    if not match:
        return text

    return f"F{match.group(1)}"

def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def derive_project_from_sample_code(sample_code: Any, parts_count: int) -> str:
    text = normalize_text(sample_code)

    if not text:
        return ""

    parts = text.split("-")

    if len(parts) < 2:
        return text

    # 25-F218a-SRU-204-GS2 -> 25-F218-SRU-204-GS2
    parts[1] = canonicalize_f_part(parts[1])

    return "-".join(parts[:min(parts_count, len(parts))])


def normalize_excel_date(value: Any) -> datetime | None:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    if isinstance(value, (int, float)):
        if value <= 0:
            return None
        return datetime(1899, 12, 30) + timedelta(days=float(value))

    text = str(value).strip()

    if not text:
        return None

    for fmt in (
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass

    return None


def project_year_folder(project_name: str) -> str:
    match = re.match(r"^(\d{2})-F", project_name)

    if not match:
        return ""

    return f"20{match.group(1)}"


def project_folder_path(base_dir: Path, folder_project: str) -> Path:
    if USE_YEAR_SUBFOLDER:
        return base_dir / project_year_folder(folder_project) / folder_project

    return base_dir / folder_project


# =============================================================================
# Read external tasks workbook
# =============================================================================

def read_external_task_rows(
    workbook_path: str,
    sheet_name: str,
    table_name: str,
) -> list[dict[str, Any]]:
    wb = openpyxl.load_workbook(
        workbook_path,
        read_only=False,
        data_only=True,
    )

    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Во внешней книге нет листа '{sheet_name}'")

        ws = wb[sheet_name]

        rows_values: list[list[Any]]

        if table_name in ws.tables:
            table = ws.tables[table_name]
            cell_range = ws[table.ref]
            rows_values = [[cell.value for cell in row] for row in cell_range]
        else:
            # fallback: читаем used range листа
            rows_values = [
                list(row)
                for row in ws.iter_rows(values_only=True)
                if any(cell is not None for cell in row)
            ]

        if not rows_values:
            return []

        headers = [
            str(value).strip() if value is not None else ""
            for value in rows_values[0]
        ]

        result: list[dict[str, Any]] = []

        for raw_row in rows_values[1:]:
            row = {
                headers[i]: raw_row[i] if i < len(raw_row) else ""
                for i in range(len(headers))
            }

            if all(normalize_text(v) == "" for v in row.values()):
                continue

            result.append(row)

        return result

    finally:
        wb.close()


def collect_recent_project_candidates(
    *,
    rows: list[dict[str, Any]],
    wanted_refresh_projects: set[str] | None,
    limit: int,
) -> list[ProjectCandidate]:
    grouped: dict[str, dict[str, Any]] = {}

    for row in rows:
        sample_code = normalize_text(row.get(TASK_SAMPLE_CODE_COL))

        if not sample_code:
            continue

        refresh_project = derive_project_from_sample_code(
            sample_code,
            REFRESH_PROJECT_PARTS,
        )

        folder_project = derive_project_from_sample_code(
            sample_code,
            FOLDER_PROJECT_PARTS,
        )

        if wanted_refresh_projects is not None and refresh_project not in wanted_refresh_projects:
            continue

        if not refresh_project or not folder_project:
            continue

        dt = normalize_excel_date(row.get(TASK_DATETIME_COL))

        old = grouped.get(refresh_project)

        if old is None:
            grouped[refresh_project] = {
                "folder_project": folder_project,
                "refresh_project": refresh_project,
                "latest_datetime": dt,
                "example_sample_code": sample_code,
            }
            continue

        old_dt = old.get("latest_datetime")

        # Берём самый свежий sampleCode внутри проекта,
        # чтобы из него получить актуальное имя папки.
        if old_dt is None or (dt is not None and dt > old_dt):
            old["folder_project"] = folder_project
            old["latest_datetime"] = dt
            old["example_sample_code"] = sample_code

    candidates: list[ProjectCandidate] = []

    for item in grouped.values():
        dt = item["latest_datetime"]

        candidates.append(
            ProjectCandidate(
                folder_project=item["folder_project"],
                refresh_project=item["refresh_project"],
                latest_datetime=dt.strftime("%Y-%m-%d %H:%M:%S") if dt else None,
                example_sample_code=item["example_sample_code"],
            )
        )

    candidates.sort(
        key=lambda c: c.latest_datetime or "",
        reverse=True,
    )

    if wanted_refresh_projects is None:
        candidates = candidates[:limit]

    return candidates

# =============================================================================
# SQLite
# =============================================================================

def get_started_projects_from_db(db_path: str) -> set[str]:
    started: set[str] = set()

    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT DISTINCT sampleCode
            FROM OP_results
            WHERE sampleCode IS NOT NULL
              AND TRIM(sampleCode) <> ''
            """
        ).fetchall()

    for (sample_code,) in rows:
        folder_project = derive_project_from_sample_code(
            sample_code,
            FOLDER_PROJECT_PARTS,
        )

        if folder_project:
            started.add(folder_project)

    return started


# =============================================================================
# Excel / xlwings
# =============================================================================

def open_or_get_addin(app: xw.App, addin_path: str) -> xw.Book:
    addin_name = Path(addin_path).name.lower()

    for book in app.books:
        if book.name.lower() == addin_name:
            return book

    return app.books.open(addin_path)


def set_project_cell(book: xw.Book, refresh_project: str) -> None:
    sheet = book.sheets[PROJECT_CELL_SHEET]
    sheet.range(PROJECT_CELL_ADDRESS).value = refresh_project


def run_refresh_tasks_for_book(
    *,
    workbook_path: Path,
    refresh_project: str,
) -> None:
    from refresh_tasks import refresh_tasks

    progress(f"Обновляю задания refresh_tasks для {refresh_project}")

    refresh_tasks(
        workbook=str(workbook_path),
        db_path=DB_PATH,
        source_workbook=TASKS_WORKBOOK,
        project_number=refresh_project,
        source_parent_sheet=REFRESH_TASKS_SOURCE_PARENT_SHEET,
        source_parent_table=REFRESH_TASKS_SOURCE_PARENT_TABLE,
        source_child_sheet=REFRESH_TASKS_SOURCE_CHILD_SHEET,
        source_child_table=REFRESH_TASKS_SOURCE_CHILD_TABLE,
        target_parent_sheet=REFRESH_TASKS_TARGET_PARENT_SHEET,
        target_parent_table=REFRESH_TASKS_TARGET_PARENT_TABLE,
        target_child_sheet=REFRESH_TASKS_TARGET_CHILD_SHEET,
        target_child_table=REFRESH_TASKS_TARGET_CHILD_TABLE,
    )


def run_after_refresh_macro(
    *,
    app: xw.App,
    workbook: xw.Book,
    addin_book: xw.Book,
) -> None:
    progress(f"Запускаю макрос надстройки: {AFTER_REFRESH_MACRO}")

    workbook.activate()

    macro = addin_book.macro(AFTER_REFRESH_MACRO)
    macro()


# =============================================================================
# Main workflow
# =============================================================================

def create_one_project_form(
    *,
    app: xw.App,
    candidate: ProjectCandidate,
    base_dir: Path,
    template_path: Path,
    overwrite: bool,
    addin_book: xw.Book,
) -> ProjectCreateReport:
    workbook: xw.Book | None = None

    try:
        folder = project_folder_path(base_dir, candidate.folder_project)
        folder.mkdir(parents=True, exist_ok=True)

        new_file = folder / f"{candidate.folder_project}_v22.xlsx"

        if new_file.exists():
            if not overwrite:
                return ProjectCreateReport(
                    project=candidate.folder_project,
                    ok=True,
                    message="Файл уже существует, пропускаю",
                    folder=str(folder),
                    workbook=str(new_file),
                    refresh_project=candidate.refresh_project,
                )

            new_file.unlink()

        progress(f"Копирую чистую форму: {candidate.folder_project}")
        shutil.copy2(template_path, new_file)

        progress(f"Открываю новую форму: {new_file.name}")
        workbook = app.books.open(
            str(new_file),
            update_links=False,
            ignore_read_only_recommended=True,
        )

        progress(f"Записываю номер проекта в {PROJECT_CELL_SHEET}!{PROJECT_CELL_ADDRESS}: {candidate.refresh_project}")
        set_project_cell(workbook, candidate.refresh_project)
        workbook.save()

        run_refresh_tasks_for_book(
            workbook_path=new_file,
            refresh_project=candidate.refresh_project,
        )

        run_after_refresh_macro(
            app=app,
            workbook=workbook,
            addin_book=addin_book,
        )

        progress("Сохраняю новую форму")
        workbook.save()

        return ProjectCreateReport(
            project=candidate.folder_project,
            ok=True,
            message="Форма создана",
            folder=str(folder),
            workbook=str(new_file),
            refresh_project=candidate.refresh_project,
        )

    except Exception as exc:
        logging.exception("Ошибка создания формы проекта %s", candidate.folder_project)

        return ProjectCreateReport(
            project=candidate.folder_project,
            ok=False,
            message=str(exc),
            refresh_project=candidate.refresh_project,
            error_type=type(exc).__name__,
        )

    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass


def create_project_forms(
    *,
    projects_text: str,
    overwrite: bool = False,
) -> CreateFormsReport:
    base_dir = Path(BASE_PROJECTS_DIR)
    template_path = Path(CLEAN_V22_TEMPLATE)

    if not base_dir.exists():
        raise FileNotFoundError(f"Корневая папка проектов не найдена: {base_dir}")

    if not template_path.exists():
        raise FileNotFoundError(f"Чистая форма v22 не найдена: {template_path}")

    progress("Читаю OP_results из SQLite, определяю уже начатые проекты")
    started_projects = get_started_projects_from_db(DB_PATH)

    wanted_projects = parse_projects_input(projects_text)

    wanted_set: set[str] | None
    if wanted_projects:
        wanted_set = set(wanted_projects)
        progress(f"Пользователь указал проекты: {', '.join(wanted_projects)}")
    else:
        wanted_set = None
        progress(f"Проекты не указаны, беру последние {RECENT_PROJECT_LIMIT} из журнала заданий")

    progress("Читаю внешний журнал заданий")
    task_rows = read_external_task_rows(
        TASKS_WORKBOOK,
        TASKS_PARENT_SHEET,
        TASKS_PARENT_TABLE,
    )

    candidates = collect_recent_project_candidates(
        rows=task_rows,
        wanted_refresh_projects=wanted_set,
        limit=RECENT_PROJECT_LIMIT,
    )

    progress(f"Найдено кандидатов проектов: {len(candidates)}")

    missing_candidates = [
        candidate
        for candidate in candidates
        if candidate.folder_project not in started_projects
    ]

    progress(f"Из них ещё нет в OP_results: {len(missing_candidates)}")

    app = xw.App(visible=True, add_book=False)

    reports: list[ProjectCreateReport] = []
    skipped_existing_files = 0

    try:
        app.display_alerts = False
        app.screen_updating = False

        progress("Открываю надстройку")
        addin_book = open_or_get_addin(app, ADDIN_PATH)

        for candidate in missing_candidates:
            progress("-" * 80)
            progress(f"Создаю форму проекта: {candidate.folder_project}")

            report = create_one_project_form(
                app=app,
                candidate=candidate,
                base_dir=base_dir,
                template_path=template_path,
                overwrite=overwrite,
                addin_book=addin_book,
            )

            if report.message == "Файл уже существует, пропускаю":
                skipped_existing_files += 1

            reports.append(report)

    finally:
        try:
            app.quit()
        except Exception:
            pass

    created_count = sum(1 for report in reports if report.ok and report.message == "Форма создана")
    ok = all(report.ok for report in reports)

    return CreateFormsReport(
        ok=ok,
        message="Создание форм завершено" if ok else "Создание форм завершено с ошибками",
        candidates_found=len(candidates),
        already_started=len(candidates) - len(missing_candidates),
        created=created_count,
        skipped_existing_files=skipped_existing_files,
        projects=reports,
    )


# =============================================================================
# CLI
# =============================================================================

def run_json(
    *,
    projects_text: str,
    overwrite: bool,
) -> str:
    try:
        report = create_project_forms(
            projects_text=projects_text,
            overwrite=overwrite,
        )

        return json.dumps(asdict(report), ensure_ascii=False, indent=2)

    except Exception as exc:
        logging.exception("Критическая ошибка создания форм")

        report = CreateFormsReport(
            ok=False,
            message=f"{type(exc).__name__}: {exc}. Подробности в логе: {LOG_PATH}",
            candidates_found=0,
            already_started=0,
            created=0,
            skipped_existing_files=0,
            projects=[],
        )

        return json.dumps(asdict(report), ensure_ascii=False, indent=2)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Автоматическое создание clean_form_v22 для новых проектов"
    )

    parser.add_argument(
        "--projects",
        default=None,
        help=(
            "Проект или диапазон проектов, например 26-F001 или 26-F001...26-F015. "
            "Если не указано, будет показано окно ввода. Пустой ввод = последние 50 проектов."
        ),
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Перезаписывать уже существующие *_v22.xlsx",
    )

    args = parser.parse_args(argv)

    projects_text = args.projects

    if projects_text is None:
        projects_text = ask_projects_text()

    json_result = run_json(
        projects_text=projects_text,
        overwrite=args.overwrite,
    )

    print(json_result)

    try:
        parsed = json.loads(json_result)
        return 0 if parsed.get("ok") else 1
    except Exception:
        return 1


if __name__ == "__main__":
    sys.exit(main())
