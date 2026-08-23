from __future__ import annotations

import re
import sqlite3
from typing import Any


class TaskStatusRepository:
    """
    Работа с историей статусов TaskStatus.

    Добавляет новую строку статуса, если последний статус по task_type + taskid
    отличается от переданного.
    """

    COMPLETED_STATUS = "Завершено"
    VALIDATED_STATUS = "Валидировано"
    CANCELED_STATUS = "Отменено"
    IN_PROGRESS_PREFIX = "В работе"

    ALLOWED_STATUSES = {
        COMPLETED_STATUS,
        VALIDATED_STATUS,
        CANCELED_STATUS,
    }

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn

    def ensure_latest_status_for_result(
            self,
            *,
            taskid: int,
            task_type: str,
            status: str,
            result_id: int | None = None,
            date_time: Any = None,
            status_comment: str | None = None,
    ) -> bool:
        status = self._normalize_status(status)

        latest_row = self.get_latest_status_row(
            taskid=taskid,
            task_type=task_type,
        )

        latest_status = None
        latest_statusid = None

        if latest_row is not None:
            latest_status = str(latest_row["status"]).strip()
            latest_statusid = int(latest_row["statusid"])

        # Если последний статус уже такой же
        # Если последний статус уже такой же
        if latest_status == status:

            # Для повторной валидации/отмены обновляем время/resultId/comment
            if status in (self.VALIDATED_STATUS, self.CANCELED_STATUS):
                if status_comment is None:
                    self.conn.execute(
                        """
                        UPDATE TaskStatus
                        SET resultId = ?,
                            dateTime = ?
                        WHERE statusid = ?
                        """,
                        (
                            result_id,
                            self._to_text(date_time),
                            latest_statusid,
                        ),
                    )
                else:
                    self.conn.execute(
                        """
                        UPDATE TaskStatus
                        SET resultId = ?,
                            dateTime = ?,
                            statusComment = ?
                        WHERE statusid = ?
                        """,
                        (
                            result_id,
                            self._to_text(date_time),
                            status_comment,
                            latest_statusid,
                        ),
                    )

                return True

            # Если снова "Завершено" после "Завершено" — TaskStatus не трогаем
            return False

        self.ensure_can_set_status(
            taskid=taskid,
            task_type=task_type,
            status=status,
        )

        self.conn.execute(
            """
            INSERT INTO TaskStatus (
                status,
                taskid,
                task_type,
                resultId,
                dateTime,
                statusComment
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                status,
                taskid,
                task_type,
                result_id,
                self._to_text(date_time),
                status_comment,
            ),
        )

        return True

    def ensure_latest_completed_status_for_result(
        self,
        *,
        taskid: int,
        task_type: str,
        result_id: int,
        date_time: Any,
    ) -> bool:
        return self.ensure_latest_status_for_result(
            taskid=taskid,
            task_type=task_type,
            status=self.COMPLETED_STATUS,
            result_id=result_id,
            date_time=date_time,
        )

    def get_latest_status(self, *, taskid: int, task_type: str) -> str | None:
        cur = self.conn.execute(
            """
            SELECT status
            FROM TaskStatus
            WHERE taskid = ?
              AND task_type = ?
            ORDER BY statusid DESC
            LIMIT 1
            """,
            (taskid, task_type),
        )

        row = cur.fetchone()
        if row is None:
            return None

        value = row["status"] if isinstance(row, sqlite3.Row) else row[0]
        if value is None:
            return None

        return str(value).strip()

    def get_latest_status_row(
            self,
            *,
            taskid: int,
            task_type: str,
    ) -> dict[str, Any] | None:
        cur = self.conn.execute(
            """
            SELECT statusid, status, resultId, dateTime
            FROM TaskStatus
            WHERE taskid = ?
              AND task_type = ?
            ORDER BY statusid DESC
            LIMIT 1
            """,
            (taskid, task_type),
        )

        row = cur.fetchone()
        if row is None:
            return None

        if isinstance(row, sqlite3.Row):
            return dict(row)

        return {
            "statusid": row[0],
            "status": row[1],
            "resultId": row[2],
            "dateTime": row[3],
        }

    def ensure_can_set_status(
            self,
            *,
            taskid: int,
            task_type: str,
            status: str,
    ) -> None:
        status = self._normalize_status(status)

        if status != self.VALIDATED_STATUS:
            return

        latest_status = self.get_latest_status(
            taskid=taskid,
            task_type=task_type,
        )

        if latest_status in (self.COMPLETED_STATUS, self.VALIDATED_STATUS):
            return

        raise ValueError(
            f"Нельзя поставить статус '{self.VALIDATED_STATUS}' "
            f"для задания {taskid} / {task_type}: "
            f"последний статус = '{latest_status}'. "
            f"Перед валидацией должен быть статус '{self.COMPLETED_STATUS}' "
            f"или уже '{self.VALIDATED_STATUS}'."
        )

    @classmethod
    def _normalize_status(cls, status: str) -> str:
        text = str(status).strip()

        aliases = {
            "completed": cls.COMPLETED_STATUS,
            "complete": cls.COMPLETED_STATUS,
            "done": cls.COMPLETED_STATUS,
            "saved": cls.COMPLETED_STATUS,
            "завершено": cls.COMPLETED_STATUS,

            "validated": cls.VALIDATED_STATUS,
            "valid": cls.VALIDATED_STATUS,
            "валидировано": cls.VALIDATED_STATUS,
            "проверено": cls.VALIDATED_STATUS,

            "canceled": cls.CANCELED_STATUS,
            "cancelled": cls.CANCELED_STATUS,
            "cancel": cls.CANCELED_STATUS,
            "отмена": cls.CANCELED_STATUS,
            "отменено": cls.CANCELED_STATUS,
        }

        normalized = aliases.get(text.lower(), text)

        if normalized in cls.ALLOWED_STATUSES:
            return normalized

        progress_match = re.fullmatch(
            r"в\s+работе\s+(\d+)\s*/\s*(\d+)",
            normalized,
            flags=re.IGNORECASE,
        )

        if progress_match is not None:
            current_step = int(progress_match.group(1))
            total_steps = int(progress_match.group(2))

            if total_steps <= 0:
                raise ValueError(
                    f"Некорректный статус '{status}': "
                    "общее количество этапов должно быть больше нуля"
                )

            if current_step < 1 or current_step > total_steps:
                raise ValueError(
                    f"Некорректный статус '{status}': "
                    f"номер этапа должен быть от 1 до {total_steps}"
                )

            return f"{cls.IN_PROGRESS_PREFIX} {current_step}/{total_steps}"

        raise ValueError(
            f"Неизвестный статус '{status}'. Допустимо: "
            f"{sorted(cls.ALLOWED_STATUSES)} или "
            f"'{cls.IN_PROGRESS_PREFIX} N/M'"
        )

    @staticmethod
    def _to_text(value: Any) -> str:
        if value is None:
            return ""

        return str(value).strip()