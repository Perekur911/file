from __future__ import annotations

from datetime import date, datetime, time
from typing import Any

from schema import ColumnSchema, TableSchema


class ValueConverter:
    """
    Единое место для преобразования значений между Excel-текстом, Python и SQLite.
    """

    TRUE_VALUES = {"1", "true", "yes", "да", "истина", "y", "+"}
    FALSE_VALUES = {"0", "false", "no", "нет", "ложь", "n", "-"}

    @classmethod
    def row_excel_to_db(cls, schema: TableSchema, row: dict[str, Any]) -> dict[str, Any]:
        """
        Преобразует строку, прочитанную из Excel, в строку для SQLite.
        """
        result: dict[str, Any] = {}

        columns_by_name = {col.db_name: col for col in schema.columns}

        for db_name, value in row.items():
            col = columns_by_name.get(db_name)
            if col is None:
                result[db_name] = value
            else:
                result[db_name] = cls.excel_to_db_value(value, col)

        return result

    @classmethod
    def rows_excel_to_db(cls, schema: TableSchema, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Преобразует список строк Excel -> SQLite."""
        return [cls.row_excel_to_db(schema, row) for row in rows]

    @classmethod
    def row_db_to_excel(cls, schema: TableSchema, row: dict[str, Any]) -> dict[str, Any]:
        """
        Преобразует строку из SQLite в текстовые значения для Excel.
        """
        result: dict[str, Any] = {}

        columns_by_name = {col.db_name: col for col in schema.columns}

        for db_name, value in row.items():
            col = columns_by_name.get(db_name)
            if col is None:
                result[db_name] = cls.db_to_excel_value(value)
            else:
                result[db_name] = cls.db_to_excel_value(value, col)

        return result

    @classmethod
    def rows_db_to_excel(cls, schema: TableSchema, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Преобразует список строк SQLite -> Excel."""
        return [cls.row_db_to_excel(schema, row) for row in rows]

    @classmethod
    def excel_to_db_value(cls, value: Any, col: ColumnSchema) -> Any:
        """
        Преобразует значение из Excel в тип, подходящий для SQLite.
        """
        value = cls._empty_to_none(value)

        if value is None:
            return None

        # temp_1, temp_2 должны спокойно жить до обработки sync.py.
        # Но до repository они обычно уже заменены на None или реальный id.
        if isinstance(value, str) and value.strip().lower().startswith("temp_"):
            return value.strip().lower()

        value_type = col.value_type.lower()

        if value_type == "text":
            return str(value).strip()

        if value_type == "integer":
            return cls._to_int(value, col.db_name)

        if value_type == "real":
            return cls._to_float(value, col.db_name)

        if value_type == "boolean":
            return cls._to_bool_int(value, col.db_name)

        if value_type == "date":
            return cls._to_date_text(value, col.db_name)

        if value_type == "time":
            return cls._to_time_text(value, col.db_name)

        if value_type == "datetime":
            return cls._to_datetime_text(value, col.db_name)

        raise ValueError(f"Неизвестный value_type='{col.value_type}' для колонки {col.db_name}")

    @classmethod
    def db_to_excel_value(cls, value: Any, col: ColumnSchema | None = None) -> Any:
        """
        Преобразует значение из SQLite в текст для Excel.
        """
        if value is None:
            return ""

        if col is None:
            return str(value)

        value_type = col.value_type.lower()

        if value_type == "boolean":
            if int(value) == 1:
                return "1"
            return "0"

        # date/time/datetime у нас уже хранятся в SQLite как TEXT в нормальном формате.
        return str(value)

    @staticmethod
    def _empty_to_none(value: Any) -> Any:
        """Пустые строки Excel -> None."""
        if value is None:
            return None

        if isinstance(value, str):
            value = value.strip()
            if value == "":
                return None
            return value

        return value

    @staticmethod
    def _to_int(value: Any, col_name: str) -> int:
        """Преобразует значение в int."""
        if isinstance(value, int):
            return value

        if isinstance(value, float):
            if not value.is_integer():
                raise ValueError(f"Колонка {col_name}: ожидалось целое число, получено {value}")
            return int(value)

        text = str(value).strip().replace(" ", "")

        text = text.removesuffix(".0")

        return int(text)

    @staticmethod
    def _to_float(value: Any, col_name: str) -> float:
        """Преобразует значение в float. Разрешает запятую как десятичный разделитель."""
        if isinstance(value, int | float):
            return float(value)

        text = str(value).strip().replace(" ", "").replace(",", ".")

        return float(text)

    @classmethod
    def _to_bool_int(cls, value: Any, col_name: str) -> int:
        """Преобразует boolean-значение в 1/0 для SQLite."""
        if isinstance(value, bool):
            return 1 if value else 0

        if isinstance(value, int):
            if value in (0, 1):
                return value

        text = str(value).strip().lower()

        if text in cls.TRUE_VALUES:
            return 1

        if text in cls.FALSE_VALUES:
            return 0

        raise ValueError(f"Колонка {col_name}: невозможно преобразовать '{value}' в boolean")

    @staticmethod
    def _to_date_text(value: Any, col_name: str) -> str:
        """
        Преобразует дату в текст yyyy-mm-dd.
        Из Excel лучше передавать именно такой формат.
        """
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")

        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")

        text = str(value).strip()

        # Если уже нормальный формат — оставляем.
        try:
            parsed = datetime.strptime(text, "%Y-%m-%d")
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            pass

        # Поддержка русского формата на всякий случай.
        try:
            parsed = datetime.strptime(text, "%d.%m.%Y")
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            pass

        raise ValueError(f"Колонка {col_name}: невозможно преобразовать '{value}' в дату")

    @staticmethod
    def _to_time_text(value: Any, col_name: str) -> str:
        """Преобразует время в текст HH:MM:SS."""
        if isinstance(value, datetime):
            return value.strftime("%H:%M:%S")

        if isinstance(value, time):
            return value.strftime("%H:%M:%S")

        text = str(value).strip()

        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                parsed = datetime.strptime(text, fmt)
                return parsed.strftime("%H:%M:%S")
            except ValueError:
                pass

        raise ValueError(f"Колонка {col_name}: невозможно преобразовать '{value}' во время")

    @staticmethod
    def _to_datetime_text(value: Any, col_name: str) -> str:
        """Преобразует дату-время в текст yyyy-mm-dd HH:MM:SS."""
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M:%S")

        text = str(value).strip()

        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d.%m.%Y %H:%M:%S"):
            try:
                parsed = datetime.strptime(text, fmt)
                return parsed.strftime("%Y-%m-%d %H:%M:%S")
            except ValueError:
                pass

        raise ValueError(f"Колонка {col_name}: невозможно преобразовать '{value}' в дату-время")