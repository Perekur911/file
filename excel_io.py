from __future__ import annotations

import gc
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlwings as xw

from schema import TableSchema
from value_converter import ValueConverter


@dataclass
class ExcelTableData:
    """
    Данные, прочитанные из умной таблицы Excel.

    rows:
        Список строк в виде dict с ключами db_name.
    row_numbers:
        Номера строк внутри DataBodyRange Excel. Нужны для точной обратной записи id.
    headers:
        Заголовки умной таблицы в текущем порядке.
    """

    rows: list[dict[str, Any]]
    row_numbers: list[int]
    headers: list[str]


class ExcelWorkbookGateway:
    """
    Работа с уже открытой книгой Excel через xlwings.

    Важное правило производительности:
    все массовые записи делаются блоками, а не по ячейкам.
    """

    def __init__(self, workbook: str):
        """
        workbook:
            Полный путь к открытой книге или имя открытой книги.
            Из VBA лучше передавать ThisWorkbook.FullName.
        """
        self.workbook_arg = workbook
        self.book = self._find_open_book(workbook)

    def read_table(self, schema: TableSchema) -> ExcelTableData:
        """
        Читает умную таблицу Excel.

        Возвращает только непустые строки. Строки возвращаются с ключами db_name,
        даже если в Excel заголовки называются иначе.
        """
        table = self._get_list_object(schema)
        headers = self._read_headers(table)
        self._validate_headers(schema, table.Name, headers)

        body = table.DataBodyRange
        if body is None or body.Value is None:
            return ExcelTableData(rows=[], row_numbers=[], headers=headers)

        values = self._to_matrix(body.Value, rows_count=body.Rows.Count, cols_count=body.Columns.Count)
        excel_to_db = schema.excel_to_db

        rows: list[dict[str, Any]] = []
        row_numbers: list[int] = []

        for row_idx, excel_row in enumerate(values, start=1):
            row_by_excel = {headers[col_idx]: excel_row[col_idx] for col_idx in range(len(headers))}
            if self._is_empty_excel_row(row_by_excel):
                continue

            db_row: dict[str, Any] = {}
            for excel_name, db_name in excel_to_db.items():
                db_row[db_name] = self._normalize_excel_value(row_by_excel.get(excel_name))

            rows.append(db_row)
            row_numbers.append(row_idx)

        return ExcelTableData(rows=rows, row_numbers=row_numbers, headers=headers)

    def write_ids(self, schema: TableSchema, row_numbers: list[int], ids: list[int]) -> None:
        """
        Записывает id обратно в id-столбец умной таблицы Excel.

        Запись выполняется одним блоком в весь id-столбец, а не по ячейкам.
        Связь строка данных -> новый id сохраняется через row_numbers.
        """
        if len(row_numbers) != len(ids):
            raise ValueError("Количество строк Excel и количество id не совпадает")

        if not row_numbers:
            return

        table = self._get_list_object(schema)
        headers = self._read_headers(table)
        pk_col_index = self._get_excel_column_index(headers, schema.pk.xlsx_name, table.Name)

        body = table.DataBodyRange
        if body is None:
            return

        id_column_range = body.Columns(pk_col_index)
        current_values = self._to_matrix(
            id_column_range.Value,
            rows_count=body.Rows.Count,
            cols_count=1,
        )

        for row_number, row_id in zip(row_numbers, ids):
            current_values[row_number - 1][0] = row_id

        id_column_range.Value = current_values

    def replace_table_rows(
            self,
            schema: TableSchema,
            db_rows: list[dict[str, Any]],
    ) -> int:
        """
        Полностью заменяет данные умной таблицы Excel данными из БД.

        Если таблица не содержит ни одной строки данных, сначала явно
        создаётся ListRow. Обычный Resize до одной пустой строки не всегда
        создаёт DataBodyRange в Excel.
        """
        table = self._get_list_object(schema)
        headers = self._read_headers(table)
        self._validate_headers(schema, table.Name, headers)

        row_count = len(db_rows)
        col_count = len(headers)

        if row_count == 0:
            body = table.DataBodyRange

            if body is not None:
                body.ClearContents()

            return 0

        # Если таблица полностью пустая, Resize на одну пустую строку
        # может оставить DataBodyRange = None.
        # Поэтому сначала явно создаём настоящую строку таблицы.
        if table.DataBodyRange is None:
            table.ListRows.Add()

            # После изменения структуры заново получаем ListObject.
            table = self._get_list_object(schema)

        # Теперь DataBodyRange уже существует, и Resize работает штатно.
        self._resize_list_object(
            table,
            data_rows=row_count,
            col_count=col_count,
        )

        # После Resize заново получаем объект таблицы.
        table = self._get_list_object(schema)
        body = table.DataBodyRange

        if body is None:
            raise RuntimeError(
                f"Не удалось создать строки данных в таблице "
                f"'{schema.excel_table_name}'. "
                f"Ожидалось строк: {row_count}"
            )

        if body.Rows.Count != row_count:
            raise RuntimeError(
                f"Таблица '{schema.excel_table_name}' получила неверный размер: "
                f"ожидалось строк {row_count}, "
                f"фактически {body.Rows.Count}"
            )

        excel_rows = ValueConverter.rows_db_to_excel(
            schema,
            db_rows,
        )

        matrix = self._build_excel_matrix(
            schema,
            headers,
            excel_rows,
        )

        body.Value = matrix

        return body.Rows.Count

    def save(self) -> None:
        """Сохраняет книгу Excel."""
        self.book.save()

    def close(self) -> None:
        """
        Освобождает ссылки xlwings/COM на пользовательский Excel.

        Книгу и приложение не закрывает: ими управляет пользователь.
        """
        self.book = None
        self.workbook_arg = None

        gc.collect()

    def _find_open_book(self, workbook: str) -> xw.Book:
        """Находит уже открытую книгу по полному пути или имени файла."""
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
            f"Передавай ThisWorkbook.FullName из VBA и убедись, что книга открыта."
        )

    def _get_list_object(self, schema: TableSchema):
        """Возвращает COM-объект ListObject для таблицы."""
        try:
            sheet = self.book.sheets[schema.sheet_name]
        except Exception as exc:
            raise ValueError(f"В книге нет листа '{schema.sheet_name}'") from exc

        list_objects = sheet.api.ListObjects
        if list_objects.Count == 0:
            raise ValueError(f"На листе '{schema.sheet_name}' нет умных таблиц")

        if schema.excel_table_name:
            try:
                return list_objects.Item(schema.excel_table_name)
            except Exception as exc:
                raise ValueError(
                    f"На листе '{schema.sheet_name}' нет умной таблицы '{schema.excel_table_name}'"
                ) from exc

        return list_objects.Item(1)

    def _resize_list_object(self, table, data_rows: int, col_count: int) -> None:
        """
        Меняет размер умной таблицы.

        data_rows — количество строк данных без строки заголовков.
        Итоговый размер = 1 строка заголовков + data_rows.
        """
        header = table.HeaderRowRange
        start_cell = header.Cells(1, 1)

        total_rows = data_rows + 1

        end_cell = start_cell.Worksheet.Cells(
            start_cell.Row + total_rows - 1,
            start_cell.Column + col_count - 1,
        )

        new_range = start_cell.Worksheet.Range(start_cell, end_cell)

        table.Resize(new_range)

    def _read_headers(self, table) -> list[str]:
        """Читает заголовки умной таблицы."""
        values = table.HeaderRowRange.Value
        if values is None:
            return []
        if isinstance(values, tuple) and values and isinstance(values[0], tuple):
            raw_headers = values[0]
        else:
            raw_headers = values
        return [str(h).strip() if h is not None else "" for h in raw_headers]

    def _validate_headers(self, schema: TableSchema, table_name: str, headers: list[str]) -> None:
        """Проверяет, что в Excel есть все столбцы из схемы."""
        missing = [c.xlsx_name for c in schema.columns if c.xlsx_name not in headers]
        if missing:
            raise ValueError(
                f"На листе '{schema.sheet_name}' в таблице '{table_name}' нет столбцов: {missing}. "
                f"Фактические столбцы: {headers}"
            )

    @staticmethod
    def _get_excel_column_index(headers: list[str], excel_name: str, table_name: str) -> int:
        """Возвращает 1-based номер столбца по заголовку Excel."""
        try:
            return headers.index(excel_name) + 1
        except ValueError as exc:
            raise ValueError(f"В таблице '{table_name}' нет столбца '{excel_name}'") from exc

    @staticmethod
    def _to_matrix(values: Any, rows_count: int, cols_count: int) -> list[list[Any]]:
        """
        Нормализует COM-значение Excel в list[list].

        xlwings/COM по-разному отдаёт одну ячейку, одну строку, один столбец и диапазон.
        Этот метод приводит все варианты к прямоугольной матрице.
        """
        if rows_count == 1 and cols_count == 1:
            return [[values]]

        if rows_count == 1:
            if isinstance(values, tuple):
                if values and isinstance(values[0], tuple):
                    return [list(values[0])]
                return [list(values)]
            return [[values]]

        if cols_count == 1:
            if isinstance(values, tuple):
                if values and isinstance(values[0], tuple):
                    return [[row[0]] for row in values]
                return [[v] for v in values]
            return [[values]]

        return [list(row) for row in values]

    @staticmethod
    def _build_excel_matrix(schema: TableSchema, headers: list[str], db_rows: list[dict[str, Any]]) -> list[list[Any]]:
        """
        Собирает матрицу для записи в Excel по текущему порядку заголовков таблицы.
        """
        excel_to_db = schema.excel_to_db
        matrix: list[list[Any]] = []

        for db_row in db_rows:
            excel_row: list[Any] = []
            for header in headers:
                db_name = excel_to_db.get(header)
                excel_row.append(db_row.get(db_name) if db_name else None)
            matrix.append(excel_row)

        return matrix

    @staticmethod
    def _is_empty_excel_row(row: dict[str, Any]) -> bool:
        """True, если строка полностью пустая."""
        return all(value is None or str(value).strip() == "" for value in row.values())

    @staticmethod
    def _normalize_excel_value(value: Any) -> Any:
        """Минимальная нормализация Excel-значений. Типы приводит ValueConverter."""
        if value is None:
            return None

        if isinstance(value, str):
            value = value.strip()
            return value if value != "" else None

        if hasattr(value, "isoformat"):
            try:
                return value.isoformat(sep=" ", timespec="seconds")
            except TypeError:
                return value.isoformat()

        return value
