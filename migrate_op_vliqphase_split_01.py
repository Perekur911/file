from __future__ import annotations

import sqlite3
from collections import Counter
from datetime import datetime
from pathlib import Path


# =============================================================================
# НАСТРОЙКИ
# =============================================================================

# ВСТАВЬ СЮДА ПУТЬ К ТЕСТОВОЙ КОПИИ БД.
DB_PATH = Path(
    r"L:\LRC\common_data\ФЛЮИДЫ\ГТИ\sqlite-excel\sqlite\results_test.db"
)

# Перед миграцией скрипт сам создаст консистентную резервную копию БД
# через sqlite3 backup API.
CREATE_BACKUP = True

OLD_TABLE = "OP_results"
NEW_TABLE = "OP_results_new"


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================

def quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def normalize_old_vliqphase(value: object) -> tuple[float | None, str | None]:
    """
    Преобразует старое VLiqPhase в новую пару:
        (VLiqPhase REAL, LiqPhaseQualifier TEXT)

    Правила:
      NULL                    -> (NULL, NULL)
      "Не выполнялось"        -> (NULL, "Не выполнялось")
      "Следы"                 -> (NULL, "Менее 1 мл")
      0 < число < 1           -> (NULL, "Менее 1 мл")
      0                       -> (0.0, NULL)
      число >= 1              -> (число, NULL)

    Любое неизвестное/отрицательное значение считается ошибкой миграции.
    """

    if value is None:
        return None, None

    text = str(value).strip()

    if text == "":
        raise ValueError(
            "Обнаружена пустая строка '' в VLiqPhase. "
            "Она не имеет согласованной семантики; миграция остановлена."
        )

    if text == "Не выполнялось":
        return None, "Не выполнялось"

    if text == "Следы":
        return None, "Менее 1 мл"

    normalized = text.replace(",", ".")

    try:
        number = float(normalized)
    except ValueError as exc:
        raise ValueError(
            f"Неизвестное значение VLiqPhase: {value!r}"
        ) from exc

    if number < 0:
        raise ValueError(
            f"Отрицательное значение VLiqPhase недопустимо: {value!r}"
        )

    if 0 < number < 1:
        return None, "Менее 1 мл"

    return number, None


def ensure_table_exists(conn: sqlite3.Connection, table_name: str) -> None:
    row = conn.execute(
        """
        SELECT 1
        FROM sqlite_schema
        WHERE type = 'table'
          AND name = ?
        """,
        (table_name,),
    ).fetchone()

    if row is None:
        raise RuntimeError(f"Таблица {table_name!r} не найдена.")


def ensure_table_absent(conn: sqlite3.Connection, table_name: str) -> None:
    row = conn.execute(
        """
        SELECT 1
        FROM sqlite_schema
        WHERE type = 'table'
          AND name = ?
        """,
        (table_name,),
    ).fetchone()

    if row is not None:
        raise RuntimeError(
            f"Таблица {table_name!r} уже существует. "
            "Скрипт специально не удаляет её автоматически. "
            "Проверь БД и удали/переименуй её вручную, если это остаток "
            "предыдущей неудачной попытки."
        )


def read_table_columns(conn: sqlite3.Connection, table_name: str) -> list[tuple]:
    return conn.execute(
        f"PRAGMA table_info({quote_ident(table_name)})"
    ).fetchall()


def validate_original_schema(conn: sqlite3.Connection) -> None:
    columns = read_table_columns(conn, OLD_TABLE)

    if not columns:
        raise RuntimeError(f"Не удалось прочитать структуру {OLD_TABLE}.")

    column_types = {
        str(row[1]): str(row[2]).upper()
        for row in columns
    }

    required_columns = {
        "resultIdOp",
        "TaskId",
        "sampleCode",
        "operator",
        "date",
        "Popen",
        "Punit",
        "Pabsolute",
        "PopenMPa",
        "Topen",
        "VH2O",
        "Pend",
        "PendMPa",
        "VLiqPhase",
        "natureLiq",
        "state",
        "comment",
        "deltaPopenPCT",
        "PopenMPaT",
        "dateTimeSync",
    }

    missing = sorted(required_columns - set(column_types))
    if missing:
        raise RuntimeError(
            "В OP_results отсутствуют ожидаемые столбцы: "
            + ", ".join(missing)
        )

    if "LiqPhaseQualifier" in column_types:
        raise RuntimeError(
            "В OP_results уже существует LiqPhaseQualifier. "
            "Похоже, миграция уже выполнялась."
        )

    if column_types["VLiqPhase"] != "TEXT":
        raise RuntimeError(
            "Ожидался старый VLiqPhase типа TEXT, "
            f"но сейчас его тип: {column_types['VLiqPhase']!r}. "
            "Скрипт остановлен, чтобы не мигрировать неизвестную версию схемы."
        )


def validate_source_values(conn: sqlite3.Connection) -> Counter:
    rows = conn.execute(
        """
        SELECT VLiqPhase, COUNT(*) AS cnt
        FROM OP_results
        GROUP BY VLiqPhase
        ORDER BY VLiqPhase
        """
    ).fetchall()

    summary: Counter = Counter()

    print("\nСтарые значения VLiqPhase:")
    print("-" * 72)

    for value, count in rows:
        new_value, qualifier = normalize_old_vliqphase(value)

        if value is None:
            old_display = "NULL"
        else:
            old_display = repr(value)

        print(
            f"{old_display:<24} -> "
            f"VLiqPhase={new_value!r:<10} "
            f"Qualifier={qualifier!r:<20} "
            f"count={count}"
        )

        if new_value is None and qualifier is None:
            summary["NULL / NULL"] += count
        elif new_value is None:
            summary[f"NULL / {qualifier}"] += count
        else:
            summary["число / NULL"] += count

    return summary


def create_backup(conn: sqlite3.Connection, db_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = db_path.with_name(
        f"{db_path.stem}_backup_before_OP_VLiqPhase_{timestamp}{db_path.suffix}"
    )

    if backup_path.exists():
        raise RuntimeError(
            f"Файл резервной копии уже существует: {backup_path}"
        )

    backup_conn = sqlite3.connect(backup_path)
    try:
        conn.backup(backup_conn)
    finally:
        backup_conn.close()

    return backup_path


def capture_schema_objects(conn: sqlite3.Connection) -> list[tuple[str, str, str]]:
    """
    Сохраняет пользовательские индексы/триггеры OP_results.
    После DROP старой таблицы они исчезнут, поэтому восстановим их.
    """
    rows = conn.execute(
        """
        SELECT type, name, sql
        FROM sqlite_schema
        WHERE tbl_name = ?
          AND type IN ('index', 'trigger')
          AND sql IS NOT NULL
        ORDER BY
            CASE type WHEN 'index' THEN 1 ELSE 2 END,
            name
        """,
        (OLD_TABLE,),
    ).fetchall()

    return [
        (str(obj_type), str(name), str(sql))
        for obj_type, name, sql in rows
    ]


def print_post_migration_distribution(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        """
        SELECT
            VLiqPhase,
            LiqPhaseQualifier,
            COUNT(*) AS cnt
        FROM OP_results
        GROUP BY VLiqPhase, LiqPhaseQualifier
        ORDER BY LiqPhaseQualifier, VLiqPhase
        """
    ).fetchall()

    print("\nРаспределение после миграции:")
    print("-" * 72)
    for vliq, qualifier, count in rows:
        print(
            f"VLiqPhase={vliq!r:<10} "
            f"Qualifier={qualifier!r:<20} "
            f"count={count}"
        )


# =============================================================================
# МИГРАЦИЯ
# =============================================================================

def migrate(conn: sqlite3.Connection) -> None:
    ensure_table_exists(conn, OLD_TABLE)
    ensure_table_absent(conn, NEW_TABLE)
    validate_original_schema(conn)

    source_count = conn.execute(
        f"SELECT COUNT(*) FROM {quote_ident(OLD_TABLE)}"
    ).fetchone()[0]

    schema_objects = capture_schema_objects(conn)

    print(f"\nСтрок в {OLD_TABLE}: {source_count}")

    if schema_objects:
        print("\nБудут восстановлены объекты схемы:")
        for obj_type, name, _sql in schema_objects:
            print(f"  {obj_type}: {name}")
    else:
        print("\nПользовательских индексов/триггеров у OP_results не найдено.")

    # Важная защита от SQLite CAST('мусор' AS REAL) -> 0.
    # До любых изменений убеждаемся, что все старые значения нам понятны.
    summary = validate_source_values(conn)

    print("\nОжидаемая сводка преобразования:")
    for key, count in summary.items():
        print(f"  {key}: {count}")

    conn.execute("BEGIN IMMEDIATE")

    try:
        conn.execute(
            """
            CREATE TABLE OP_results_new (
                resultIdOp INTEGER PRIMARY KEY AUTOINCREMENT,
                TaskId INTEGER,
                sampleCode TEXT,
                operator TEXT,
                date TEXT,
                Popen REAL,
                Punit TEXT,
                Pabsolute INTEGER,
                PopenMPa REAL,
                Topen REAL,
                VH2O REAL,
                Pend REAL,
                PendMPa REAL,
                VLiqPhase REAL,
                LiqPhaseQualifier TEXT,
                natureLiq TEXT,
                state INTEGER,
                comment TEXT,
                deltaPopenPCT REAL,
                PopenMPaT REAL,
                dateTimeSync TEXT
            )
            """
        )

        conn.execute(
            """
            INSERT INTO OP_results_new (
                resultIdOp,
                TaskId,
                sampleCode,
                operator,
                date,
                Popen,
                Punit,
                Pabsolute,
                PopenMPa,
                Topen,
                VH2O,
                Pend,
                PendMPa,
                VLiqPhase,
                LiqPhaseQualifier,
                natureLiq,
                state,
                comment,
                deltaPopenPCT,
                PopenMPaT,
                dateTimeSync
            )
            SELECT
                resultIdOp,
                TaskId,
                sampleCode,
                operator,
                date,
                Popen,
                Punit,
                Pabsolute,
                PopenMPa,
                Topen,
                VH2O,
                Pend,
                PendMPa,

                CASE
                    WHEN VLiqPhase IS NULL
                        THEN NULL

                    WHEN TRIM(VLiqPhase) IN (
                        'Следы',
                        'Не выполнялось'
                    )
                        THEN NULL

                    WHEN CAST(
                        REPLACE(TRIM(VLiqPhase), ',', '.')
                        AS REAL
                    ) > 0
                    AND CAST(
                        REPLACE(TRIM(VLiqPhase), ',', '.')
                        AS REAL
                    ) < 1
                        THEN NULL

                    ELSE CAST(
                        REPLACE(TRIM(VLiqPhase), ',', '.')
                        AS REAL
                    )
                END AS VLiqPhase,

                CASE
                    WHEN VLiqPhase IS NULL
                        THEN NULL

                    WHEN TRIM(VLiqPhase) = 'Не выполнялось'
                        THEN 'Не выполнялось'

                    WHEN TRIM(VLiqPhase) = 'Следы'
                        THEN 'Менее 1 мл'

                    WHEN CAST(
                        REPLACE(TRIM(VLiqPhase), ',', '.')
                        AS REAL
                    ) > 0
                    AND CAST(
                        REPLACE(TRIM(VLiqPhase), ',', '.')
                        AS REAL
                    ) < 1
                        THEN 'Менее 1 мл'

                    ELSE NULL
                END AS LiqPhaseQualifier,

                natureLiq,
                state,
                comment,
                deltaPopenPCT,
                PopenMPaT,
                dateTimeSync

            FROM OP_results
            """
        )

        new_count = conn.execute(
            f"SELECT COUNT(*) FROM {quote_ident(NEW_TABLE)}"
        ).fetchone()[0]

        if new_count != source_count:
            raise RuntimeError(
                "Количество строк после копирования не совпало: "
                f"old={source_count}, new={new_count}"
            )

        # В новой модели два поля взаимоисключающие.
        both_filled = conn.execute(
            """
            SELECT COUNT(*)
            FROM OP_results_new
            WHERE VLiqPhase IS NOT NULL
              AND LiqPhaseQualifier IS NOT NULL
            """
        ).fetchone()[0]

        if both_filled != 0:
            raise RuntimeError(
                "После преобразования найдены строки, где одновременно "
                f"заполнены VLiqPhase и LiqPhaseQualifier: {both_filled}"
            )

        # Числовые значения между 0 и 1 больше не должны существовать.
        sub_one = conn.execute(
            """
            SELECT COUNT(*)
            FROM OP_results_new
            WHERE VLiqPhase > 0
              AND VLiqPhase < 1
            """
        ).fetchone()[0]

        if sub_one != 0:
            raise RuntimeError(
                "После преобразования остались числовые VLiqPhase "
                f"между 0 и 1: {sub_one}"
            )

        # Qualifier должен содержать только два согласованных значения.
        bad_qualifiers = conn.execute(
            """
            SELECT COUNT(*)
            FROM OP_results_new
            WHERE LiqPhaseQualifier IS NOT NULL
              AND LiqPhaseQualifier NOT IN (
                  'Менее 1 мл',
                  'Не выполнялось'
              )
            """
        ).fetchone()[0]

        if bad_qualifiers != 0:
            raise RuntimeError(
                "После преобразования появились неизвестные qualifier: "
                f"{bad_qualifiers}"
            )

        # Теперь можно заменить таблицу.
        conn.execute("DROP TABLE OP_results")
        conn.execute("ALTER TABLE OP_results_new RENAME TO OP_results")

        # Восстанавливаем все существовавшие пользовательские индексы/триггеры.
        for obj_type, name, sql in schema_objects:
            conn.execute(sql)
            print(f"Восстановлен {obj_type}: {name}")

        # AUTOINCREMENT: убеждаемся, что следующий ID будет больше текущего max.
        max_result_id = conn.execute(
            """
            SELECT COALESCE(MAX(resultIdOp), 0)
            FROM OP_results
            """
        ).fetchone()[0]

        seq_row = conn.execute(
            """
            SELECT 1
            FROM sqlite_sequence
            WHERE name = 'OP_results'
            """
        ).fetchone()

        if seq_row is None:
            conn.execute(
                """
                INSERT INTO sqlite_sequence(name, seq)
                VALUES('OP_results', ?)
                """,
                (max_result_id,),
            )
        else:
            conn.execute(
                """
                UPDATE sqlite_sequence
                SET seq = ?
                WHERE name = 'OP_results'
                """,
                (max_result_id,),
            )

        # Проверяем итоговую структуру до COMMIT.
        final_columns = {
            str(row[1]): str(row[2]).upper()
            for row in read_table_columns(conn, OLD_TABLE)
        }

        if final_columns.get("VLiqPhase") != "REAL":
            raise RuntimeError(
                "Итоговый тип VLiqPhase не REAL: "
                f"{final_columns.get('VLiqPhase')!r}"
            )

        if final_columns.get("LiqPhaseQualifier") != "TEXT":
            raise RuntimeError(
                "Итоговый тип LiqPhaseQualifier не TEXT: "
                f"{final_columns.get('LiqPhaseQualifier')!r}"
            )

        final_count = conn.execute(
            "SELECT COUNT(*) FROM OP_results"
        ).fetchone()[0]

        if final_count != source_count:
            raise RuntimeError(
                "Количество строк в итоговой OP_results изменилось: "
                f"before={source_count}, after={final_count}"
            )

        # Проверяем внешние ключи всей БД, если они есть.
        fk_violations = conn.execute(
            "PRAGMA foreign_key_check"
        ).fetchall()

        if fk_violations:
            raise RuntimeError(
                "PRAGMA foreign_key_check обнаружил нарушения: "
                f"{fk_violations[:10]}"
            )

        quick_check = conn.execute(
            "PRAGMA quick_check"
        ).fetchone()

        if quick_check is None or quick_check[0] != "ok":
            raise RuntimeError(
                f"PRAGMA quick_check завершился с ошибкой: {quick_check}"
            )

        conn.commit()

    except Exception:
        conn.rollback()
        raise

    print_post_migration_distribution(conn)


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    db_path = DB_PATH.expanduser()

    if not db_path.exists():
        raise FileNotFoundError(
            f"БД не найдена: {db_path}\n"
            "Измени DB_PATH в начале скрипта."
        )

    if not db_path.is_file():
        raise RuntimeError(f"DB_PATH не является файлом: {db_path}")

    print("=" * 72)
    print("Миграция OP_results: VLiqPhase TEXT -> REAL")
    print("Добавление LiqPhaseQualifier TEXT")
    print("=" * 72)
    print(f"БД: {db_path}")

    conn = sqlite3.connect(db_path)

    try:
        # Для пересоздания таблицы безопаснее не давать SQLite блокировать DROP
        # из-за существующих FK. После замены выполняется foreign_key_check.
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("PRAGMA busy_timeout = 30000")

        if CREATE_BACKUP:
            backup_path = create_backup(conn, db_path)
            print(f"Резервная копия: {backup_path}")

        migrate(conn)

    finally:
        conn.close()

    print("\n" + "=" * 72)
    print("МИГРАЦИЯ УСПЕШНО ЗАВЕРШЕНА")
    print("=" * 72)
    print(
        "Если выше не было исключений, OP_results заменена атомарно: "
        "при любой ошибке до COMMIT исходная таблица откатывается."
    )


if __name__ == "__main__":
    main()
