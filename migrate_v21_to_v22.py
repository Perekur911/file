from __future__ import annotations

import argparse
import json
import logging
import re
import shutil
import sys
import traceback
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import xlwings as xw


# =============================================================================
# Настройки
# =============================================================================

BASE_PROJECTS_DIR = r"L:\LRC\common_data\ФЛЮИДЫ\ПТИ\PROJECTS"
CLEAN_V22_TEMPLATE = r"L:\LRC\common_data\ФЛЮИДЫ\ПТИ\sqlite-excel\clean_form_v22.xlsx"

PROJECT_DIR = Path(__file__).resolve().parent
LOG_PATH = PROJECT_DIR / "migrate_v21_to_v22.log"

STUDIES = ["AP", "BP", "GC", "GOR", "SSF", "OP", "EMV"]

UNPROTECT_PASSWORDS = ("1984", "9184", "")
PROTECT_PASSWORD = "1984"

SKIP_RESULT_FIELDS = {
    "datetimesync",
    "page",
}

SKIP_FIELD_PREFIXES = (
    "resultid",
    "rowid",
)

OP_SKIP_FIELDS = {
    "page",
    "datetimesync",
    "typeSampler",
    "volumeSampler",
    "numberSampler",
    "regime",
    "transferDate",
    "PsamplingMPa",
    "Tsampling",
    "resultIdOP",
    "PopenMPa",
    "PendMPa",
    "BP_bpMPa",
    "BP_deltaPCT",
    "natureLiq",
    "deltaPopenPCT",
    "PopenMPaT",
    "typeSample"
}

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
class ProjectMigrationReport:
    project: str
    ok: bool
    message: str
    folder: str | None = None
    old_file: str | None = None
    new_file: str | None = None
    forms_copied: int = 0
    named_values_copied: int = 0
    table_values_copied: int = 0
    warnings: list[str] | None = None
    error_type: str | None = None


@dataclass
class MigrationReport:
    ok: bool
    message: str
    projects: list[ProjectMigrationReport]


# =============================================================================
# Project parsing / search
# =============================================================================

def parse_project_input(text: str) -> list[str]:
    text = text.strip()

    if not text:
        raise ValueError("Не указан проект или диапазон проектов")

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
            result.append(normalize_project_code(part))

    # Убираем дубли, порядок сохраняем
    seen = set()
    unique: list[str] = []

    for project in result:
        if project not in seen:
            unique.append(project)
            seen.add(project)

    return unique


def normalize_project_code(value: str) -> str:
    value = value.strip().upper()

    if not re.fullmatch(r"\d{2}-F\d+", value):
        raise ValueError(f"Некорректный формат проекта: {value}")

    prefix, num = value.split("F", 1)
    return f"{prefix}F{int(num):0{len(num)}d}"


def expand_project_range(start: str, end: str) -> list[str]:
    start = normalize_project_code(start)
    end = normalize_project_code(end)

    m1 = re.fullmatch(r"(\d{2}-F)(\d+)", start)
    m2 = re.fullmatch(r"(\d{2}-F)(\d+)", end)

    if not m1 or not m2:
        raise ValueError(f"Некорректный диапазон проектов: {start}...{end}")

    prefix1, n1_text = m1.groups()
    prefix2, n2_text = m2.groups()

    if prefix1 != prefix2:
        raise ValueError(f"Диапазон проектов должен быть внутри одного года/префикса: {start}...{end}")

    n1 = int(n1_text)
    n2 = int(n2_text)
    width = max(len(n1_text), len(n2_text))

    if n2 < n1:
        raise ValueError(f"Конец диапазона меньше начала: {start}...{end}")

    return [f"{prefix1}{n:0{width}d}" for n in range(n1, n2 + 1)]


def find_project_folder(base_dir: Path, project_code: str, recursive: bool = False) -> Path:
    if not base_dir.exists():
        raise FileNotFoundError(f"Основная папка проектов не найдена: {base_dir}")

    candidates = []

    iterator = base_dir.rglob("*") if recursive else base_dir.iterdir()

    for path in iterator:
        if path.is_dir() and path.name.upper().startswith(project_code.upper()):
            candidates.append(path)

    if not candidates:
        raise FileNotFoundError(f"Не найдена папка проекта, начинающаяся на {project_code}")

    candidates = sorted(candidates, key=lambda p: len(p.name))

    if len(candidates) > 1:
        logging.warning(
            "Для проекта %s найдено несколько папок, беру первую: %s; все варианты: %s",
            project_code,
            candidates[0],
            candidates,
        )

    return candidates[0]


def find_v21_file(project_folder: Path) -> Path:
    folder_name = project_folder.name.lower()

    files = []

    for path in project_folder.iterdir():
        if not path.is_file():
            continue

        if path.name.startswith("~$"):
            continue

        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            continue

        name_lower = path.name.lower()

        if name_lower.startswith(folder_name) and "_v21" in name_lower:
            files.append(path)

    if not files:
        raise FileNotFoundError(
            f"В папке '{project_folder}' не найден файл, начинающийся с имени папки и содержащий '_v21'"
        )

    if len(files) > 1:
        raise ValueError(
            f"В папке '{project_folder}' найдено несколько v21-файлов: {files}"
        )

    return files[0]


def make_v22_file(project_folder: Path, template_path: Path, overwrite: bool = False) -> Path:
    new_path = project_folder / f"{project_folder.name}_v22.xlsx"

    if new_path.exists():
        if not overwrite:
            raise FileExistsError(
                f"Файл v22 уже существует: {new_path}. "
                f"Для перезаписи запусти с --overwrite"
            )

        new_path.unlink()

    shutil.copy2(template_path, new_path)

    return new_path


# =============================================================================
# Excel helpers
# =============================================================================

def open_book(app: xw.App, path: Path, read_only: bool) -> xw.Book:
    return app.books.open(
        str(path),
        read_only=read_only,
        update_links=False,
        ignore_read_only_recommended=True,
    )


def unprotect_all_sheets(book: xw.Book) -> None:
    for sheet in book.sheets:
        for password in UNPROTECT_PASSWORDS:
            try:
                sheet.api.Unprotect(Password=password)
                break
            except Exception:
                pass

def protect_all_sheets(book: xw.Book) -> None:
    for sheet in book.sheets:
        try:
            sheet.api.Protect(
                Password=PROTECT_PASSWORD,
                DrawingObjects=True,
                Contents=True,
                Scenarios=True,
                UserInterfaceOnly=True,
                AllowFiltering=True,
            )
        except Exception as exc:
            logging.warning("Не удалось защитить лист %s: %s", sheet.name, exc)

def get_named_range(book: xw.Book, name: str):
    try:
        return book.api.Names.Item(name).RefersToRange
    except Exception:
        try:
            return book.names[name].refers_to_range.api
        except Exception:
            return None


def named_range_exists(book: xw.Book, name: str) -> bool:
    return get_named_range(book, name) is not None


def get_list_object(book: xw.Book, sheet_name: str, table_name: str):
    try:
        sheet = book.sheets[sheet_name]
    except Exception as exc:
        raise ValueError(f"В книге '{book.name}' нет листа '{sheet_name}'") from exc

    try:
        return sheet.api.ListObjects.Item(table_name)
    except Exception as exc:
        raise ValueError(
            f"На листе '{sheet_name}' нет умной таблицы '{table_name}'"
        ) from exc


def try_get_list_object(book: xw.Book, sheet_name: str, table_name: str):
    try:
        return get_list_object(book, sheet_name, table_name)
    except Exception:
        return None


def list_object_exists(book: xw.Book, sheet_name: str, table_name: str) -> bool:
    return try_get_list_object(book, sheet_name, table_name) is not None


def read_table_headers(table) -> list[str]:
    values = table.HeaderRowRange.Value

    if values is None:
        return []

    if isinstance(values, tuple) and values and isinstance(values[0], tuple):
        raw = values[0]
    else:
        raw = values

    return [str(v).strip() if v is not None else "" for v in raw]


def header_index_map(table) -> dict[str, int]:
    headers = read_table_headers(table)
    return {h: idx + 1 for idx, h in enumerate(headers) if h}


def ensure_table_rows(table, row_count: int) -> None:
    if row_count <= 0:
        return

    while table.ListRows.Count < row_count:
        table.ListRows.Add()


def cell_has_formula(cell) -> bool:
    try:
        if bool(cell.MergeCells):
            top_left = cell.MergeArea.Cells(1, 1)
            return bool(top_left.HasFormula)

        return bool(cell.HasFormula)
    except Exception:
        return False


def get_cell_value_from_range(rng):
    if rng is None:
        return None

    try:
        if bool(rng.MergeCells):
            return rng.MergeArea.Cells(1, 1).Value

        return rng.Cells(1, 1).Value
    except Exception:
        return rng.Value


def write_value_to_range_if_not_formula(target_rng, value: Any) -> bool:
    if target_rng is None:
        return False

    try:
        if bool(target_rng.MergeCells):
            write_cell = target_rng.MergeArea.Cells(1, 1)
        else:
            write_cell = target_rng.Cells(1, 1)

        if cell_has_formula(write_cell):
            return False

        write_cell.Value = value
        return True

    except Exception:
        return False


def copy_named_range_value(
    old_book: xw.Book,
    new_book: xw.Book,
    range_name: str,
) -> bool:
    old_rng = get_named_range(old_book, range_name)
    new_rng = get_named_range(new_book, range_name)

    if old_rng is None or new_rng is None:
        return False

    value = get_cell_value_from_range(old_rng)

    return write_value_to_range_if_not_formula(new_rng, value)


def should_skip_field(field_name: str) -> bool:
    f = str(field_name).strip().lower()

    if not f:
        return True

    if f in SKIP_RESULT_FIELDS:
        return True

    return any(f.startswith(prefix) for prefix in SKIP_FIELD_PREFIXES)


def copy_table_values_by_common_columns(
    old_table,
    new_table,
    *,
    skip_fields: set[str] | None = None,
) -> int:
    if skip_fields is None:
        skip_fields = set()

    old_map = header_index_map(old_table)
    new_map = header_index_map(new_table)

    common_headers = [
        header
        for header in new_map.keys()
        if header in old_map
        and header.lower() not in skip_fields
        and not should_skip_field(header)
    ]

    if not common_headers:
        return 0

    old_rows_count = old_table.ListRows.Count

    if old_rows_count == 0:
        return 0

    ensure_table_rows(new_table, old_rows_count)

    old_body = old_table.DataBodyRange
    new_body = new_table.DataBodyRange

    if old_body is None or new_body is None:
        return 0

    copied = 0

    for row_idx in range(1, old_rows_count + 1):
        for header in common_headers:
            old_col = old_map[header]
            new_col = new_map[header]

            src_cell = old_body.Cells(row_idx, old_col)
            dst_cell = new_body.Cells(row_idx, new_col)

            if cell_has_formula(dst_cell):
                continue

            dst_cell.Value = src_cell.Value
            copied += 1

    return copied


# =============================================================================
# Study migration
# =============================================================================

def source_form_table_name(study: str, form_index: int) -> str:
    if study in ("GOR", "SSF"):
        return f"{study}_{form_index}"

    return f"{study}_source_{form_index}"


def migrate_named_range_study(
    *,
    old_book: xw.Book,
    new_book: xw.Book,
    study: str,
    warnings: list[str],
) -> tuple[int, int, int]:
    """
    Переносит обычные формы:
    AP, BP, GC, GOR, SSF, EMV.

    Parent-поля берутся по заголовкам новой таблицы <study>_results.
    Child/source-таблицы копируются по общим заголовкам таблиц формы.
    """
    forms_copied = 0
    named_values_copied = 0
    table_values_copied = 0

    result_table = try_get_list_object(
        new_book,
        f"{study}_results",
        f"{study}_results",
    )

    if result_table is None:
        warnings.append(f"{study}: не найдена новая таблица {study}_results, пропускаю")
        return 0, 0, 0

    result_fields = read_table_headers(result_table)
    form_index = 1

    while named_range_exists(old_book, f"{study}_{form_index}_sampleCode"):
        if not named_range_exists(new_book, f"{study}_{form_index}_sampleCode"):
            warnings.append(
                f"{study}_{form_index}: в новой форме нет именованного диапазона sampleCode, форма пропущена"
            )
            form_index += 1
            continue

        form_has_any_data = False

        for field in result_fields:
            if should_skip_field(field):
                continue

            range_name = f"{study}_{form_index}_{field}"

            if not named_range_exists(old_book, range_name):
                continue

            if not named_range_exists(new_book, range_name):
                continue

            if copy_named_range_value(old_book, new_book, range_name):
                named_values_copied += 1
                form_has_any_data = True

        old_src_table_name = source_form_table_name(study, form_index)
        new_src_table_name = source_form_table_name(study, form_index)

        old_src_table = try_get_list_object(old_book, study, old_src_table_name)
        new_src_table = try_get_list_object(new_book, study, new_src_table_name)

        if old_src_table is not None and new_src_table is not None:
            copied = copy_table_values_by_common_columns(
                old_src_table,
                new_src_table,
                skip_fields={"page"},
            )
            table_values_copied += copied

            if copied > 0:
                form_has_any_data = True

        elif old_src_table is not None and new_src_table is None:
            warnings.append(
                f"{study}_{form_index}: в старой форме есть таблица {old_src_table_name}, "
                f"а в новой не найдена"
            )

        if form_has_any_data:
            forms_copied += 1

        form_index += 1

    return forms_copied, named_values_copied, table_values_copied


def migrate_op_table(
    *,
    old_book: xw.Book,
    new_book: xw.Book,
    warnings: list[str],
) -> tuple[int, int]:
    """
    OP — особый случай: форма OP сама является умной таблицей.
    """
    old_table = try_get_list_object(old_book, "OP", "OP")
    new_table = try_get_list_object(new_book, "OP", "OP")

    if old_table is None:
        warnings.append("OP: в старой книге не найдена таблица OP")
        return 0, 0

    if new_table is None:
        warnings.append("OP: в новой книге не найдена таблица OP")
        return 0, 0

    copied = copy_table_values_by_common_columns(
        old_table,
        new_table,
        skip_fields=OP_SKIP_FIELDS,
    )

    forms_copied = old_table.ListRows.Count if copied > 0 else 0

    return forms_copied, copied


def migrate_workbook_data(
    *,
    old_book: xw.Book,
    new_book: xw.Book,
) -> tuple[int, int, int, list[str]]:
    warnings: list[str] = []

    total_forms = 0
    total_named = 0
    total_table_values = 0

    unprotect_all_sheets(new_book)

    for study in STUDIES:
        logging.info("Миграция исследования %s", study)

        if study == "OP":
            forms, table_values = migrate_op_table(
                old_book=old_book,
                new_book=new_book,
                warnings=warnings,
            )

            total_forms += forms
            total_table_values += table_values
            continue

        forms, named_values, table_values = migrate_named_range_study(
            old_book=old_book,
            new_book=new_book,
            study=study,
            warnings=warnings,
        )

        total_forms += forms
        total_named += named_values
        total_table_values += table_values

    return total_forms, total_named, total_table_values, warnings


# =============================================================================
# Project processing
# =============================================================================

def migrate_one_project(
    *,
    app: xw.App,
    project_code: str,
    base_dir: Path,
    template_path: Path,
    recursive: bool,
    overwrite: bool,
) -> ProjectMigrationReport:
    warnings: list[str] = []

    old_book = None
    new_book = None

    try:
        folder = find_project_folder(base_dir, project_code, recursive=recursive)
        old_file = find_v21_file(folder)
        new_file = make_v22_file(folder, template_path, overwrite=overwrite)

        logging.info("Проект %s: folder=%s", project_code, folder)
        logging.info("Проект %s: old=%s", project_code, old_file)
        logging.info("Проект %s: new=%s", project_code, new_file)

        old_book = open_book(app, old_file, read_only=True)
        new_book = open_book(app, new_file, read_only=False)

        forms, named_values, table_values, warnings = migrate_workbook_data(
            old_book=old_book,
            new_book=new_book,
        )

        protect_all_sheets(new_book)

        new_book.save()

        return ProjectMigrationReport(
            project=project_code,
            ok=True,
            message="Миграция выполнена",
            folder=str(folder),
            old_file=str(old_file),
            new_file=str(new_file),
            forms_copied=forms,
            named_values_copied=named_values,
            table_values_copied=table_values,
            warnings=warnings,
        )

    except Exception as exc:
        logging.exception("Ошибка миграции проекта %s", project_code)

        return ProjectMigrationReport(
            project=project_code,
            ok=False,
            message=str(exc),
            warnings=warnings,
            error_type=type(exc).__name__,
        )

    finally:
        if old_book is not None:
            try:
                old_book.close()
            except Exception:
                pass

        if new_book is not None:
            try:
                new_book.close()
            except Exception:
                pass


def run_migration(
    *,
    projects_text: str,
    base_dir: str = BASE_PROJECTS_DIR,
    template_path: str = CLEAN_V22_TEMPLATE,
    recursive: bool = False,
    overwrite: bool = False,
) -> MigrationReport:
    projects = parse_project_input(projects_text)

    base_path = Path(base_dir)
    template = Path(template_path)

    if not template.exists():
        raise FileNotFoundError(f"Чистая форма v22 не найдена: {template}")

    app = xw.App(visible=False, add_book=False)

    try:
        app.display_alerts = False
        app.screen_updating = False

        reports: list[ProjectMigrationReport] = []

        for project_code in projects:
            report = migrate_one_project(
                app=app,
                project_code=project_code,
                base_dir=base_path,
                template_path=template,
                recursive=recursive,
                overwrite=overwrite,
            )
            reports.append(report)

        ok = all(r.ok for r in reports)

        return MigrationReport(
            ok=ok,
            message="Миграция завершена" if ok else "Миграция завершена с ошибками",
            projects=reports,
        )

    finally:
        try:
            app.quit()
        except Exception:
            pass


# =============================================================================
# CLI / JSON
# =============================================================================

def run_json(
    *,
    projects_text: str,
    base_dir: str = BASE_PROJECTS_DIR,
    template_path: str = CLEAN_V22_TEMPLATE,
    recursive: bool = False,
    overwrite: bool = False,
) -> str:
    try:
        report = run_migration(
            projects_text=projects_text,
            base_dir=base_dir,
            template_path=template_path,
            recursive=recursive,
            overwrite=overwrite,
        )

        return json.dumps(asdict(report), ensure_ascii=False, indent=2)

    except Exception as exc:
        logging.exception("Критическая ошибка миграции")

        report = MigrationReport(
            ok=False,
            message=f"{type(exc).__name__}: {exc}. Подробности в логе: {LOG_PATH}",
            projects=[],
        )

        return json.dumps(asdict(report), ensure_ascii=False, indent=2)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Миграция данных из старой формы v21 в новую чистую форму v22"
    )

    parser.add_argument(
        "--projects",
        default=None,
        help="Проект или диапазон проектов. Например: 26-F001 или 26-F001...26-F015",
    )

    parser.add_argument(
        "--base-dir",
        default=BASE_PROJECTS_DIR,
        help="Папка, внутри которой искать папки проектов",
    )

    parser.add_argument(
        "--template",
        default=CLEAN_V22_TEMPLATE,
        help="Путь к чистой форме v22",
    )

    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Искать папки проектов рекурсивно внутри base-dir",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Перезаписывать уже существующий файл *_v22.xlsx",
    )

    args = parser.parse_args(argv)

    projects_text = args.projects

    if not projects_text:
        projects_text = input(
            "Введите проект или диапазон проектов, например 26-F001 или 26-F001...26-F015: "
        ).strip()

    json_result = run_json(
        projects_text=projects_text,
        base_dir=args.base_dir,
        template_path=args.template,
        recursive=args.recursive,
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
