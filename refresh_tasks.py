from __future__ import annotations

import argparse
import json
import sqlite3
import sys
import traceback
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from datetime import datetime, timedelta

import openpyxl
import xlwings as xw


# =============================================================================
# Настройки колонок
# =============================================================================

PARENT_TASK_ID_COL = "Номер задания"
PARENT_SAMPLE_CODE_COL = "Код проекта"
PARENT_TASK_TYPE_COL = "Задание"
PARENT_STATUS_COL = "Статус"
PARENT_ROW_ID_COL = "ID"
PARENT_DATETIME_COL = "Дата и время"
CHILD_TASK_ID_COL = "Номер ГТИ"


# =============================================================================
# Отчёт
# =============================================================================

@dataclass
class RefreshTasksReport:
    ok: bool
    message: str
    project: str
    parent_rows_read: int = 0
    parent_rows_written: int = 0
    child_rows_read: int = 0
    child_rows_written: int = 0
    sqlite_statuses_applied: int = 0
    external_statuses_kept: int = 0
    error_type: str | None = None
    traceback: str | None = None


# =============================================================================
# openpyxl: чтение внешней книги с заданиями
# =============================================================================

def read_excel_table_openpyxl(
    workbook_path: str,
    *,
    sheet_name: str,
    table_name: str,
) -> tuple[list[str], list[dict[str, Any]]]:
    """
    Читает умную таблицу из внешней Excel-книги через openpyxl.

    Внешняя книга открывается только на чтение:
    - read_only=False нужен, потому что openpyxl в read_only-режиме хуже работает с таблицами.
    - data_only=True читает рассчитанные значения формул, а не формулы.
    """
    wb = openpyxl.load_workbook(
        workbook_path,
        read_only=False,
        data_only=True,
    )

    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Во внешней книге нет листа '{sheet_name}'")

        ws = wb[sheet_name]

        if table_name not in ws.tables:
            raise ValueError(
                f"На листе '{sheet_name}' во внешней книге нет таблицы '{table_name}'"
            )

        table = ws.tables[table_name]
        table_range = ws[table.ref]

        rows_values = [
            [cell.value for cell in row]
            for row in table_range
        ]

        if not rows_values:
            return [], []

        headers = [
            str(value).strip() if value is not None else ""
            for value in rows_values[0]
        ]

        rows: list[dict[str, Any]] = []

        for raw_row in rows_values[1:]:
            row = {
                headers[i]: normalize_excel_value(raw_row[i])
                for i in range(len(headers))
            }

            if is_empty_row(row):
                continue

            rows.append(row)

        return headers, rows

    finally:
        wb.close()


def normalize_excel_value(value: Any) -> Any:
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        return int(value)

    return value


def is_empty_row(row: dict[str, Any]) -> bool:
    return all(str(v).strip() == "" for v in row.values())


# =============================================================================
# xlwings: запись в открытую книгу с формами
# =============================================================================

def find_open_book(workbook: str) -> xw.Book:
    target = str(Path(workbook)).lower()
    target_name = Path(workbook).name.lower()

    for app in xw.apps:
        for book in app.books:
            full_name = str(book.fullname).lower() if book.fullname else ""
            book_name = book.name.lower()

            if full_name == target or book_name == target_name:
                return book

    raise FileNotFoundError(
        f"Открытая книга Excel не найдена: {workbook}. "
        f"Передавай путь к уже открытой книге с формами."
    )


def get_list_object(book: xw.Book, sheet_name: str, table_name: str):
    try:
        sheet = book.sheets[sheet_name]
    except Exception as exc:
        raise ValueError(f"В книге '{book.name}' нет листа '{sheet_name}'") from exc

    list_objects = sheet.api.ListObjects

    if list_objects.Count == 0:
        raise ValueError(f"На листе '{sheet_name}' нет умных таблиц")

    try:
        return list_objects.Item(table_name)
    except Exception as exc:
        raise ValueError(
            f"На листе '{sheet_name}' нет умной таблицы '{table_name}'"
        ) from exc


def read_headers_from_list_object(table) -> list[str]:
    values = table.HeaderRowRange.Value

    if values is None:
        return []

    if isinstance(values, tuple) and values and isinstance(values[0], tuple):
        raw_headers = values[0]
    else:
        raw_headers = values

    return [str(h).strip() if h is not None else "" for h in raw_headers]


def resize_list_object(table, data_rows: int, col_count: int) -> None:
    """
    Меняет размер умной таблицы.

    data_rows — количество строк данных без заголовка.
    Оставляем минимум одну строку данных.
    """
    header = table.HeaderRowRange
    start_cell = header.Cells(1, 1)

    total_rows = max(data_rows, 1) + 1

    end_cell = start_cell.Worksheet.Cells(
        start_cell.Row + total_rows - 1,
        start_cell.Column + col_count - 1,
    )

    new_range = start_cell.Worksheet.Range(start_cell, end_cell)
    table.Resize(new_range)


def replace_table_rows(table, rows: list[dict[str, Any]]) -> int:
    """
    Полностью заменяет данные умной таблицы в открытой книге Excel.

    Порядок колонок берётся из текущих заголовков целевой таблицы.
    """
    headers = read_headers_from_list_object(table)
    col_count = len(headers)
    row_count = len(rows)

    resize_list_object(table, data_rows=max(row_count, 1), col_count=col_count)

    body = table.DataBodyRange
    if body is None:
        return 0

    if row_count == 0:
        body.ClearContents()
        return 0

    matrix: list[list[Any]] = []

    for row in rows:
        matrix.append([row.get(header, "") for header in headers])

    body.Value = matrix

    return row_count


# =============================================================================
# Фильтрация и нормализация
# =============================================================================

def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_task_id(value: Any) -> int | None:
    if value is None:
        return None

    text = str(value).strip()

    if text == "":
        return None

    if text.endswith(".0") or text.endswith(",0"):
        text = text[:-2]

    return int(float(text.replace(",", ".")))

def normalize_excel_datetime(value: Any) -> datetime | None:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    if isinstance(value, (int, float)):
        if value <= 0:
            return None
        return datetime(1899, 12, 30) + timedelta(days=float(value))

    text = str(value).strip()

    if text == "":
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

def external_parent_latest_sort_key(
    row: dict[str, Any],
    original_index: int,
) -> tuple[int, float, int]:
    """
    Чем больше ключ — тем строка считается новее.

    Приоритет:
    1. ID строки внешней таблицы
    2. Дата и время
    3. Порядок строки в файле
    """
    row_id = normalize_task_id(row.get(PARENT_ROW_ID_COL))
    dt = normalize_excel_datetime(row.get(PARENT_DATETIME_COL))

    row_id_key = row_id if row_id is not None else -1
    dt_key = dt.timestamp() if dt is not None else 0.0

    return row_id_key, dt_key, original_index


def keep_latest_parent_rows_by_task_id(
    rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Во внешней таблице заданий может быть несколько строк с одним taskId.
    Оставляем только последнюю строку по каждому 'Номер задания'.

    Дубли ищем только по taskId.
    """
    latest_by_task_id: dict[int, tuple[tuple[int, float, int], int, dict[str, Any]]] = {}
    rows_without_task_id: list[tuple[int, dict[str, Any]]] = []

    for index, row in enumerate(rows):
        task_id = normalize_task_id(row.get(PARENT_TASK_ID_COL))

        if task_id is None:
            rows_without_task_id.append((index, dict(row)))
            continue

        sort_key = external_parent_latest_sort_key(row, index)

        old = latest_by_task_id.get(task_id)

        if old is None or sort_key > old[0]:
            latest_by_task_id[task_id] = (sort_key, index, dict(row))

    selected: list[tuple[int, dict[str, Any]]] = [
        (index, row)
        for _, index, row in latest_by_task_id.values()
    ]

    selected.extend(rows_without_task_id)
    selected.sort(key=lambda item: item[0])

    return [row for _, row in selected]

def filter_parent_rows(
    rows: list[dict[str, Any]],
    project_number: str,
) -> list[dict[str, Any]]:
    """
    Родительскую таблицу фильтруем по колонке 'Код проекта'.

    Пример:
        project_number = '25-F218'
        подходит '25-F218-SRU-204-GS1'
    """
    result: list[dict[str, Any]] = []

    for row in rows:
        sample_code = normalize_text(row.get(PARENT_SAMPLE_CODE_COL))

        if sample_code.startswith(project_number):
            result.append(dict(row))

    return result


def filter_child_rows(
    rows: list[dict[str, Any]],
    task_ids: set[int],
) -> list[dict[str, Any]]:
    """
    Дочернюю таблицу фильтруем по 'Номер ГТИ',
    который соответствует taskId / Номер задания.
    """
    result: list[dict[str, Any]] = []

    for row in rows:
        task_id = normalize_task_id(row.get(CHILD_TASK_ID_COL))

        if task_id in task_ids:
            result.append(dict(row))

    return result


# =============================================================================
# SQLite: последние статусы
# =============================================================================

def fetch_latest_statuses(
    db_path: str,
    task_ids: set[int],
) -> dict[int, str]:
    """
    Возвращает последний статус из SQLite TaskStatus по taskid.

    taskId считается уникальным независимо от типа исследования.
    """
    if not task_ids:
        return {}

    task_ids_sorted = sorted(task_ids)
    placeholders = ", ".join("?" for _ in task_ids_sorted)

    sql = f"""
        SELECT
            taskid,
            status
        FROM TaskStatus
        WHERE taskid IN ({placeholders})
        ORDER BY statusid
    """

    result: dict[int, str] = {}

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row

        cur = conn.execute(sql, task_ids_sorted)

        # ORDER BY statusid ASC:
        # более поздний статус перезапишет ранний.
        for row in cur.fetchall():
            task_id = int(row["taskid"])
            status = "" if row["status"] is None else str(row["status"]).strip()

            if status:
                result[task_id] = status

    return result


def apply_sqlite_statuses(
    parent_rows: list[dict[str, Any]],
    sqlite_statuses: dict[int, str],
) -> tuple[int, int]:
    """
    Заменяет только колонку 'Статус' в родительских строках.

    Если в SQLite есть последний статус по taskId — ставим его.
    Если нет — оставляем статус из внешней Excel-книги.
    """
    sqlite_statuses_applied = 0
    external_statuses_kept = 0

    for row in parent_rows:
        task_id = normalize_task_id(row.get(PARENT_TASK_ID_COL))

        if task_id is None:
            external_statuses_kept += 1
            continue

        sqlite_status = sqlite_statuses.get(task_id)

        if sqlite_status:
            row[PARENT_STATUS_COL] = sqlite_status
            sqlite_statuses_applied += 1
        else:
            external_statuses_kept += 1

    return sqlite_statuses_applied, external_statuses_kept


# =============================================================================
# Главная логика
# =============================================================================

def refresh_tasks(
    *,
    workbook: str,
    db_path: str,
    source_workbook: str,
    project_number: str,
    source_parent_sheet: str,
    source_parent_table: str,
    source_child_sheet: str,
    source_child_table: str,
    target_parent_sheet: str,
    target_parent_table: str,
    target_child_sheet: str,
    target_child_table: str,
) -> RefreshTasksReport:
    target_book = find_open_book(workbook)

    parent_headers, external_parent_rows_all = read_excel_table_openpyxl(
        source_workbook,
        sheet_name=source_parent_sheet,
        table_name=source_parent_table,
    )

    child_headers, external_child_rows_all = read_excel_table_openpyxl(
        source_workbook,
        sheet_name=source_child_sheet,
        table_name=source_child_table,
    )

    require_columns(
        parent_headers,
        [
            PARENT_TASK_ID_COL,
            PARENT_SAMPLE_CODE_COL,
            PARENT_TASK_TYPE_COL,
            PARENT_STATUS_COL,
        ],
        f"внешняя таблица {source_parent_table}",
    )

    require_columns(
        child_headers,
        [
            CHILD_TASK_ID_COL,
        ],
        f"внешняя таблица {source_child_table}",
    )

    parent_rows = filter_parent_rows(
        external_parent_rows_all,
        project_number,
    )

    parent_rows_before_dedupe = len(parent_rows)

    parent_rows = keep_latest_parent_rows_by_task_id(parent_rows)

    parent_duplicates_removed = parent_rows_before_dedupe - len(parent_rows)

    task_ids: set[int] = set()

    for row in parent_rows:
        task_id = normalize_task_id(row.get(PARENT_TASK_ID_COL))

        if task_id is None:
            continue

        task_ids.add(task_id)

    sqlite_statuses = fetch_latest_statuses(db_path, task_ids)

    sqlite_statuses_applied, external_statuses_kept = apply_sqlite_statuses(
        parent_rows,
        sqlite_statuses,
    )

    child_rows = filter_child_rows(
        external_child_rows_all,
        task_ids,
    )

    target_parent_table_obj = get_list_object(
        target_book,
        target_parent_sheet,
        target_parent_table,
    )

    target_child_table_obj = get_list_object(
        target_book,
        target_child_sheet,
        target_child_table,
    )

    parent_written = replace_table_rows(target_parent_table_obj, parent_rows)
    child_written = replace_table_rows(target_child_table_obj, child_rows)

    target_book.save()

    return RefreshTasksReport(
        ok=True,
        message=(
            "Задания успешно обновлены. "
            f"Дубликатов по taskId удалено: {parent_duplicates_removed}"
        ),
        project=project_number,
        parent_rows_read=len(external_parent_rows_all),
        parent_rows_written=parent_written,
        child_rows_read=len(external_child_rows_all),
        child_rows_written=child_written,
        sqlite_statuses_applied=sqlite_statuses_applied,
        external_statuses_kept=external_statuses_kept,
    )


def require_columns(
    actual_headers: list[str],
    required_headers: list[str],
    source_name: str,
) -> None:
    missing = [
        header
        for header in required_headers
        if header not in actual_headers
    ]

    if missing:
        raise ValueError(
            f"Не найдены обязательные колонки в {source_name}: {missing}. "
            f"Фактические колонки: {actual_headers}"
        )


def run_json(**kwargs: Any) -> str:
    try:
        report = refresh_tasks(**kwargs)
        return json.dumps(asdict(report), ensure_ascii=False, indent=2)

    except Exception as exc:
        report = RefreshTasksReport(
            ok=False,
            message=str(exc),
            project=str(kwargs.get("project_number", "")),
            error_type=type(exc).__name__,
            traceback=traceback.format_exc(),
        )
        return json.dumps(asdict(report), ensure_ascii=False, indent=2)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Обновление таблиц task/task_mix из внешней Excel-книги "
            "с подмешиванием последних статусов из SQLite"
        )
    )

    parser.add_argument(
        "--workbook",
        required=True,
        help="Текущая открытая книга с формами",
    )

    parser.add_argument(
        "--db",
        required=True,
        help="Путь к SQLite БД",
    )

    parser.add_argument(
        "--source-workbook",
        required=True,
        help="Внешняя книга Excel с заданиями",
    )

    parser.add_argument(
        "--project",
        required=True,
        help="Номер проекта, например 25-F218",
    )

    parser.add_argument("--source-parent-sheet", default="task")
    parser.add_argument("--source-parent-table", default="task")
    parser.add_argument("--source-child-sheet", default="task_mix")
    parser.add_argument("--source-child-table", default="task_mix")

    parser.add_argument("--target-parent-sheet", default="Task")
    parser.add_argument("--target-parent-table", default="Task")
    parser.add_argument("--target-child-sheet", default="Task_mix")
    parser.add_argument("--target-child-table", default="Task_mix")

    args = parser.parse_args(argv)

    json_result = run_json(
        workbook=args.workbook,
        db_path=args.db,
        source_workbook=args.source_workbook,
        project_number=args.project,
        source_parent_sheet=args.source_parent_sheet,
        source_parent_table=args.source_parent_table,
        source_child_sheet=args.source_child_sheet,
        source_child_table=args.source_child_table,
        target_parent_sheet=args.target_parent_sheet,
        target_parent_table=args.target_parent_table,
        target_child_sheet=args.target_child_sheet,
        target_child_table=args.target_child_table,
    )

    print(json_result)

    try:
        parsed = json.loads(json_result)
        return 0 if parsed.get("ok") else 1
    except Exception:
        return 1


if __name__ == "__main__":
    sys.exit(main())
