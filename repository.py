from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from typing import Any, Iterable

from schema import TableSchema
from value_converter import ValueConverter


@dataclass
class SaveTableResult:
    """
    Результат сохранения одной таблицы Excel -> SQLite.

    inserted:
        Сколько строк добавлено.
    updated:
        Сколько строк обновлено.
    skipped:
        Сколько строк пропущено.
    ids:
        id для каждой входной строки в том же порядке.
    """

    inserted: int = 0
    updated: int = 0
    skipped: int = 0
    ids: list[int] | None = None

    def __post_init__(self) -> None:
        if self.ids is None:
            self.ids = []


class SQLiteRepository:
    """
    Слой работы с SQLite.

    Поддерживает два направления:
    1. Excel -> SQLite: INSERT новых строк и UPDATE существующих по id.
    2. SQLite -> Excel: чтение актуальных строк по номеру проекта.
    """

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")

    def close(self) -> None:
        """Закрывает соединение с БД."""
        self.conn.close()

    def begin(self) -> None:
        """Начинает транзакцию."""
        self.conn.execute("BEGIN")

    def commit(self) -> None:
        """Фиксирует транзакцию."""
        self.conn.commit()

    def rollback(self) -> None:
        """Откатывает транзакцию."""
        self.conn.rollback()

    def save_rows(self, schema: TableSchema, rows: list[dict[str, Any]]) -> SaveTableResult:
        """
        Добавляет или обновляет строки в SQLite.

        Если id пустой — INSERT, новый id возвращается через lastrowid.
        Если id заполнен — UPDATE по id.
        Если UPDATE не нашёл строку — INSERT с указанным id.
        """
        result = SaveTableResult()

        for index, row in enumerate(rows, start=1):
            row = ValueConverter.row_excel_to_db(schema, row)
            self._validate_row(schema, row, index)
            pk_name = schema.pk.db_name
            pk_value = self._empty_to_none(row.get(pk_name))

            try:
                if pk_value is None:
                    new_id = self._insert_row(schema, row, include_pk=False)
                    result.inserted += 1
                    result.ids.append(new_id)
                else:
                    updated = self._update_row(schema, row)

                    if updated:
                        result.updated += 1
                        result.ids.append(int(pk_value))
                    else:
                        new_id = self._insert_row(schema, row, include_pk=True)
                        result.inserted += 1
                        result.ids.append(new_id)

            except sqlite3.IntegrityError as exc:
                raise ValueError(
                    self._build_integrity_error_message(
                        schema=schema,
                        row=row,
                        row_index=index,
                        exc=exc,
                    )
                ) from exc

        return result

    def _build_integrity_error_message(
        self,
        *,
        schema: TableSchema,
        row: dict[str, Any],
        row_index: int,
        exc: sqlite3.IntegrityError,
    ) -> str:
        raw_message = str(exc)

        unique_columns = self._extract_unique_columns_from_integrity_error(
            raw_message
        )

        if unique_columns:
            return self._build_unique_constraint_message(
                schema=schema,
                row=row,
                row_index=row_index,
                columns=unique_columns,
                raw_message=raw_message,
            )

        return (
            f"Ошибка ограничения БД при сохранении таблицы {schema.db_table}, "
            f"строка данных №{row_index}.\n\n"
            f"Техническая ошибка SQLite:\n{raw_message}"
        )

    @staticmethod
    def _extract_unique_columns_from_integrity_error(
        message: str,
    ) -> list[str]:
        """
        Из сообщения SQLite:
            UNIQUE constraint failed: AP_results.sampleCode, AP_results.analyseNumber

        возвращает:
            ["sampleCode", "analyseNumber"]
        """
        match = re.search(
            r"UNIQUE constraint failed:\s*(.+)$",
            message,
            flags=re.IGNORECASE,
        )

        if not match:
            return []

        raw_columns = match.group(1).split(",")

        result: list[str] = []

        for raw_column in raw_columns:
            raw_column = raw_column.strip()

            if "." in raw_column:
                _, column_name = raw_column.rsplit(".", 1)
            else:
                column_name = raw_column

            column_name = column_name.strip().strip('"').strip("'")

            if column_name:
                result.append(column_name)

        return result

    def _build_unique_constraint_message(
        self,
        *,
        schema: TableSchema,
        row: dict[str, Any],
        row_index: int,
        columns: list[str],
        raw_message: str,
    ) -> str:
        key_parts: list[str] = []

        for column in columns:
            excel_name = schema.db_to_excel.get(column, column)
            value = row.get(column)
            key_parts.append(f"{excel_name} = {self._format_value_for_message(value)}")

        key_text = ", ".join(key_parts)

        existing = self._find_existing_row_by_columns(
            schema=schema,
            columns=columns,
            row=row,
        )

        existing_text = ""

        if existing is not None:
            pk_name = schema.pk.db_name
            existing_pk = existing.get(pk_name)

            if existing_pk is not None:
                existing_text = (
                    f"\n\nВ БД уже есть строка с таким ключом: "
                    f"{pk_name} = {existing_pk}."
                )

        # Специальное человекочитаемое сообщение для AP.
        if (
            schema.db_table == "AP_results"
            and [c.lower() for c in columns] == ["samplecode", "analysenumber"]
        ):
            return (
                "Нельзя сохранить AP: в БД уже есть результат с такой же парой "
                "«шифр пробы + номер анализа».\n\n"
                f"Проблемная пара:\n{key_text}"
                f"{existing_text}\n\n"
                "Проверь поле sampleCode и analyseNumber в форме. "
                "Для AP эта пара должна быть уникальной."
            )

        return (
            f"Нельзя сохранить таблицу {schema.db_table}: нарушено ограничение уникальности.\n\n"
            f"Проблемная комбинация:\n{key_text}"
            f"{existing_text}\n\n"
            f"Строка данных №{row_index}.\n\n"
            f"Техническая ошибка SQLite:\n{raw_message}"
        )

    def _find_existing_row_by_columns(
        self,
        *,
        schema: TableSchema,
        columns: list[str],
        row: dict[str, Any],
    ) -> dict[str, Any] | None:
        where_parts: list[str] = []
        params: list[Any] = []

        for column in columns:
            value = self._empty_to_none(row.get(column))

            if value is None:
                where_parts.append(f"{self._q(column)} IS NULL")
            else:
                where_parts.append(f"{self._q(column)} = ?")
                params.append(value)

        select_columns = [schema.pk.db_name] + [
            column for column in columns
            if column != schema.pk.db_name
        ]

        columns_sql = self._columns_sql(select_columns)
        where_sql = " AND ".join(where_parts)

        sql = (
            f"SELECT {columns_sql} "
            f"FROM {self._q(schema.db_table)} "
            f"WHERE {where_sql} "
            f"LIMIT 1"
        )

        cur = self.conn.execute(sql, params)
        existing = cur.fetchone()

        if existing is None:
            return None

        return dict(existing)

    @staticmethod
    def _format_value_for_message(value: Any) -> str:
        if value is None:
            return "<пусто>"

        text = str(value).strip()

        if text == "":
            return "<пусто>"

        return f"'{text}'"

    def fetch_results_by_project(self, schema: TableSchema, project_number: str) -> list[dict[str, Any]]:
        """
        Возвращает строки *_results по номеру проекта.

        Условие:
            sampleCode LIKE '<project_number>-%'

        Пример:
            project_number='25-F123'
            найдёт sampleCode вида '25-F123-...', но не '25-F1234-...'.
        """
        if "sampleCode" not in schema.db_columns:
            raise ValueError(f"В таблице {schema.logical_name} нет столбца sampleCode")

        columns_sql = self._columns_sql(schema.db_columns)
        sql = (
            f"SELECT {columns_sql} "
            f"FROM {self._q(schema.db_table)} "
            f"WHERE {self._q('sampleCode')} LIKE ? "
            f"ORDER BY {self._q('sampleCode')}, {self._q(schema.pk.db_name)}"
        )
        return self._fetch_dicts(sql, [f"{project_number}%"])

    def fetch_child_rows_by_parent_ids(self, schema: TableSchema, parent_ids: list[int]) -> list[dict[str, Any]]:
        """
        Возвращает дочерние строки по id родительских результатов.

        Для AP_sourceData это обычно:
            WHERE resultId IN (...)
        """
        if not schema.fk_column:
            raise ValueError(f"Для дочерней таблицы {schema.logical_name} не указан fk_column")

        if not parent_ids:
            return []

        columns_sql = self._columns_sql(schema.db_columns)
        placeholders = ", ".join("?" for _ in parent_ids)
        sql = (
            f"SELECT {columns_sql} "
            f"FROM {self._q(schema.db_table)} "
            f"WHERE {self._q(schema.fk_column)} IN ({placeholders}) "
            f"ORDER BY {self._q(schema.fk_column)}, {self._q(schema.pk.db_name)}"
        )
        return self._fetch_dicts(sql, parent_ids)

    def _insert_row(self, schema: TableSchema, row: dict[str, Any], *, include_pk: bool) -> int:
        """Выполняет INSERT и возвращает id строки."""
        pk_name = schema.pk.db_name
        columns = schema.insert_columns()
        if include_pk:
            columns = [pk_name] + columns

        values = [self._empty_to_none(row.get(c)) for c in columns]
        placeholders = ", ".join("?" for _ in columns)
        column_sql = self._columns_sql(columns)

        sql = f"INSERT INTO {self._q(schema.db_table)} ({column_sql}) VALUES ({placeholders})"
        cur = self.conn.execute(sql, values)

        if include_pk:
            return int(row[pk_name])
        return int(cur.lastrowid)

    def _update_row(self, schema: TableSchema, row: dict[str, Any]) -> bool:
        """Выполняет UPDATE по primary key."""
        pk_name = schema.pk.db_name
        columns = schema.update_columns()
        if not columns:
            return False

        set_sql = ", ".join(f"{self._q(c)} = ?" for c in columns)
        values = [self._empty_to_none(row.get(c)) for c in columns]
        values.append(row[pk_name])

        sql = f"UPDATE {self._q(schema.db_table)} SET {set_sql} WHERE {self._q(pk_name)} = ?"
        cur = self.conn.execute(sql, values)
        return cur.rowcount > 0

    def _validate_row(self, schema: TableSchema, row: dict[str, Any], index: int) -> None:
        """Проверяет обязательные поля перед записью Excel -> SQLite."""
        missing: list[str] = []
        for col in schema.columns:
            if col.required and self._empty_to_none(row.get(col.db_name)) is None:
                missing.append(col.db_name)

        if missing:
            raise ValueError(
                f"Таблица {schema.logical_name}, строка данных #{index}: "
                f"не заполнены обязательные поля: {missing}. Строка: {row}"
            )

    def _fetch_dicts(self, sql: str, params: Iterable[Any]) -> list[dict[str, Any]]:
        """Выполняет SELECT и возвращает строки как list[dict]."""
        cur = self.conn.execute(sql, list(params))
        return [dict(row) for row in cur.fetchall()]

    def _columns_sql(self, columns: list[str]) -> str:
        """Формирует SQL-список экранированных колонок."""
        return ", ".join(self._q(c) for c in columns)

    def delete_rows_by_pk(self, schema: TableSchema, ids: list[Any]) -> int:
        """
        Удаляет строки таблицы по primary key.

        Используется для строк вида:
            rowIdAP есть, но данных уже нет.
        """
        if not ids:
            return 0

        pk_name = schema.pk.db_name
        placeholders = ", ".join("?" for _ in ids)

        sql = (
            f"DELETE FROM {self._q(schema.db_table)} "
            f"WHERE {self._q(pk_name)} IN ({placeholders})"
        )

        cur = self.conn.execute(sql, ids)
        return cur.rowcount

    @staticmethod
    def _empty_to_none(value: Any) -> Any:
        """Преобразует пустую строку в NULL."""
        if value is None:
            return None
        if isinstance(value, str) and value.strip() == "":
            return None
        return value

    @staticmethod
    def _q(identifier: str) -> str:
        """Экранирует имя таблицы/колонки двойными кавычками."""
        return '"' + identifier.replace('"', '""') + '"'
