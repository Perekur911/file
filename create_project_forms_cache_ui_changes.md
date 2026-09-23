# `create_project_forms.py`: точный порядок добавления JSON-кэша и нового UI

Инструкция составлена по текущей структуре файла, где имеются функции и классы:

- `ProjectInspection`;
- `inspect_project`;
- `write_scan_report`;
- `run_scan`;
- `create_or_update_one_project`;
- `run_mutating_mode`;
- `ask_ui_options`;
- `execute`;
- `main`.

Основную логику поиска, обновления и пересоздания форм переписывать не нужно. Изменяются только перечисленные ниже участки.

---

## 1. Изменить импорты в самом начале файла

### Найти

```python
import argparse
import json
import logging
import re
import shutil
import sqlite3
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
```

### Заменить на

```python
import argparse
import json
import logging
import queue
import re
import shutil
import sqlite3
import sys
import threading
from dataclasses import asdict, dataclass, field, fields, replace
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable
```

`queue` и `threading` нужны, чтобы длительное сканирование не замораживало окно. `fields` нужен для безопасной загрузки JSON старого формата, а `replace` — для копирования объекта из кэша.

---

## 2. Добавить константы кэша

### Найти участок настроек

```python
ARCHIVE_DIR_NAME = "архив"
LOG_PATH = PROJECT_DIR / "project_forms.log"
```

### Сразу после него вставить

```python
SCAN_CACHE_PATH = PROJECT_DIR / "project_scan_cache.json"
SCAN_CACHE_SCHEMA_VERSION = 1

HARD_BLOCKING_ACTIONS = {
    "BLOCKED_FOLDER_MISMATCH",
    "BLOCKED_MULTIPLE_FOLDERS",
    "BLOCKED_MULTIPLE_WORKBOOKS",
    "BLOCKED_NO_FOLDER_INFO",
    "BLOCKED_METADATA_ERROR",
    "BLOCKED_WORKBOOK_OPEN",
    "BLOCKED_VERSION_AHEAD",
    "MIGRATION_REQUIRED",
}
```

В `.gitignore` отдельной строкой добавить:

```gitignore
project_scan_cache.json
project_scan_cache.json.tmp
```

---

## 3. Расширить `ProjectInspection`

### Найти класс

```python
@dataclass
class ProjectInspection:
```

### В конец класса, сразу после поля

```python
    excel_lock_present: bool = False
```

### Вставить

```python
    task_signature: str = ""
    workbook_modified_at: str = ""
    workbook_mtime_ns: int = 0
    workbook_size: int = 0
    inspected_at: str = ""
    inspection_source: str = ""
    selected_action: str = ""
    execution_result: str = ""
```

Назначение полей:

- `task_signature` — позволяет обнаружить изменение данных проекта в журнале заданий;
- `workbook_modified_at` — понятная дата изменения для UI и JSON;
- `workbook_mtime_ns` и `workbook_size` — машинная проверка изменения файла;
- `inspected_at` — когда metadata книги действительно перечитывалась;
- `inspection_source` — `fresh` или `cache`;
- `selected_action` — действие, назначенное пользователем в UI;
- `execution_result` — последний результат выполнения.

### Удалить больше не используемый класс

Полностью удалить:

```python
@dataclass
class UiOptions:
    mode: str
    projects_text: str
    year: int
    cancelled: bool = False
```

---

## 4. Добавить функции JSON-кэша

### Место вставки

Найти функцию:

```python
def progress(message: str) -> None:
```

Вставить следующий блок **после всей функции `progress`**, но до `normalize_text`.

```python
def iso_now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def task_project_signature(task_info: TaskProjectInfo | None) -> str:
    if task_info is None:
        return ""

    payload = {
        "refresh_project": task_info.refresh_project,
        "latest_datetime": (
            task_info.latest_datetime.isoformat()
            if task_info.latest_datetime is not None
            else ""
        ),
        "latest_sample_code": task_info.latest_sample_code,
        "new_folder_name": task_info.new_folder_name,
        "search_keys": list(task_info.search_keys),
    }
    return json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def workbook_file_signature(path: Path) -> tuple[str, int, int]:
    stat = path.stat()
    modified_at = datetime.fromtimestamp(
        stat.st_mtime
    ).astimezone().isoformat(timespec="seconds")
    return modified_at, stat.st_mtime_ns, stat.st_size


def fill_inspection_cache_fields(
    inspection: ProjectInspection,
    task_info: TaskProjectInfo | None,
    *,
    source: str,
) -> ProjectInspection:
    inspection.task_signature = task_project_signature(task_info)
    inspection.inspection_source = source

    if source == "fresh":
        inspection.inspected_at = iso_now()

    if inspection.selected_workbook:
        path = Path(inspection.selected_workbook)
        try:
            (
                inspection.workbook_modified_at,
                inspection.workbook_mtime_ns,
                inspection.workbook_size,
            ) = workbook_file_signature(path)
        except OSError:
            inspection.workbook_modified_at = ""
            inspection.workbook_mtime_ns = 0
            inspection.workbook_size = 0
    else:
        inspection.workbook_modified_at = ""
        inspection.workbook_mtime_ns = 0
        inspection.workbook_size = 0

    return inspection


def load_scan_cache(release: ReleaseInfo) -> dict[str, ProjectInspection]:
    if not SCAN_CACHE_PATH.exists():
        return {}

    try:
        raw = json.loads(SCAN_CACHE_PATH.read_text(encoding="utf-8"))

        if raw.get("cacheSchemaVersion") != SCAN_CACHE_SCHEMA_VERSION:
            logger.info("Кэш сканирования имеет другую версию формата")
            return {}

        cached_release = raw.get("release") or {}
        if cached_release.get("formVersion") != release.form_version:
            logger.info(
                "Кэш создан для другой версии формы: %s вместо %s",
                cached_release.get("formVersion"),
                release.form_version,
            )
            return {}

        allowed_fields = {item.name for item in fields(ProjectInspection)}
        result: dict[str, ProjectInspection] = {}

        for raw_item in raw.get("projects", []):
            if not isinstance(raw_item, dict):
                continue

            values = {
                key: value
                for key, value in raw_item.items()
                if key in allowed_fields
            }
            project = str(values.get("refresh_project", "")).strip()
            if not project:
                continue

            item = ProjectInspection(**values)
            # При загрузке с диска источник всегда кэш.
            item.inspection_source = "cache"
            result[project] = item

        return result

    except Exception:
        logger.exception("Не удалось прочитать кэш %s", SCAN_CACHE_PATH)
        return {}


def save_scan_cache(
    inspections: dict[str, ProjectInspection],
    release: ReleaseInfo,
    *,
    last_scan: dict[str, Any] | None = None,
) -> None:
    payload = {
        "cacheSchemaVersion": SCAN_CACHE_SCHEMA_VERSION,
        "generatedAt": iso_now(),
        "release": {
            "formVersion": release.form_version,
            "addinVersion": release.addin_version,
        },
        "lastScan": last_scan or {},
        "projects": [
            asdict(inspections[key])
            for key in sorted(inspections)
        ],
    }

    SCAN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = SCAN_CACHE_PATH.with_name(SCAN_CACHE_PATH.name + ".tmp")
    temp_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temp_path.replace(SCAN_CACHE_PATH)


def inspection_allows_recreate(inspection: ProjectInspection) -> bool:
    if inspection.recommended_action in HARD_BLOCKING_ACTIONS:
        return False
    if not inspection.selected_workbook:
        return False
    return inspection.safe_to_recreate
```

Кэш специально считается несовместимым при изменении `formVersion`. После релиза новой формы один полный проход потребуется снова, зато между релизами неизменённые книги перечитываться не будут.

---

## 5. Добавить проверку возможности использования записи из кэша

### Место вставки

Найти конец функции:

```python
def inspect_project(...)
```

Следующий блок вставить **после полного окончания `inspect_project` и до заголовка `# Scan report`**.

```python
def try_reuse_cached_inspection(
    *,
    refresh_project: str,
    task_info: TaskProjectInfo | None,
    release: ReleaseInfo,
    cached: ProjectInspection | None,
    force: bool,
) -> ProjectInspection | None:
    if force or cached is None:
        return None

    if cached.current_form_version != release.form_version:
        return None

    if cached.task_signature != task_project_signature(task_info):
        return None

    # Отсутствующие книги проверяются полноценно каждый раз:
    # книга или папка могли появиться после предыдущего запуска.
    if not cached.selected_workbook or not cached.selected_folder:
        return None

    # Запись без успешно прочитанной версии metadata не используем.
    if not cached.workbook_version:
        return None

    workbook = Path(cached.selected_workbook)
    folder = Path(cached.selected_folder)

    if not workbook.is_file() or not folder.is_dir():
        return None

    search_key = cached.folder_search_key or workbook_search_key(folder, task_info)
    candidates = find_project_workbooks(folder, search_key)

    if len(candidates) != 1:
        return None

    if str(candidates[0]).casefold() != str(workbook).casefold():
        return None

    lock_present = excel_lock_present(workbook)

    # Если прошлый кэш был записан при открытой книге, после её закрытия
    # metadata необходимо прочитать заново.
    if cached.excel_lock_present and not lock_present:
        return None

    try:
        modified_at, mtime_ns, size = workbook_file_signature(workbook)
    except OSError:
        return None

    if (
        mtime_ns != cached.workbook_mtime_ns
        or size != cached.workbook_size
    ):
        return None

    reused = replace(cached)
    reused.current_form_version = release.form_version
    reused.workbook_modified_at = modified_at
    reused.workbook_mtime_ns = mtime_ns
    reused.workbook_size = size
    reused.inspection_source = "cache"
    reused.excel_lock_present = lock_present

    if lock_present:
        reused.issue = "Обнаружен Excel lock-файл; книга, вероятно, открыта"
        reused.recommended_action = "BLOCKED_WORKBOOK_OPEN"
        reused.selected_action = ""

    return reused


def run_cached_scan(
    *,
    release: ReleaseInfo,
    projects_text: str,
    year: int,
    cache: dict[str, ProjectInspection] | None = None,
    force_all: bool = False,
    force_projects: set[str] | None = None,
    item_callback: Callable[[ProjectInspection, int, int], None] | None = None,
) -> tuple[list[ProjectInspection], dict[str, ProjectInspection]]:
    task_rows = read_external_task_rows(
        TASKS_WORKBOOK,
        TASKS_PARENT_SHEET,
        TASKS_PARENT_TABLE,
    )
    task_infos = build_task_project_infos(task_rows)

    requested = parse_projects_input(projects_text)
    projects = project_ids_for_scan(requested, year, task_infos)
    all_cache = dict(cache if cache is not None else load_scan_cache(release))
    forced = {value.upper() for value in (force_projects or set())}

    scan_info = {
        "startedAt": iso_now(),
        "year": year,
        "projectsText": projects_text,
        "forceAll": force_all,
    }

    scanned_items: list[ProjectInspection] = []
    total = len(projects)

    for index, project in enumerate(projects, start=1):
        task_info = task_infos.get(project)
        old_item = all_cache.get(project)
        force = force_all or project.upper() in forced

        inspection = try_reuse_cached_inspection(
            refresh_project=project,
            task_info=task_info,
            release=release,
            cached=old_item,
            force=force,
        )

        if inspection is None:
            inspection = inspect_project(project, task_info, release)
            inspection = fill_inspection_cache_fields(
                inspection,
                task_info,
                source="fresh",
            )

            # Сохраняем ранее выбранное действие только тогда, когда оно
            # остаётся допустимым после повторной проверки.
            if old_item is not None:
                inspection.execution_result = old_item.execution_result
                if old_item.selected_action == "update":
                    if inspection.recommended_action not in HARD_BLOCKING_ACTIONS:
                        inspection.selected_action = "update"
                elif old_item.selected_action == "recreate":
                    if inspection_allows_recreate(inspection):
                        inspection.selected_action = "recreate"
        else:
            inspection = fill_inspection_cache_fields(
                inspection,
                task_info,
                source="cache",
            )

        all_cache[project] = inspection
        scanned_items.append(inspection)

        # Сохраняем после каждого проекта. При сбое уже проверенные строки
        # не будут потеряны.
        save_scan_cache(all_cache, release, last_scan=scan_info)

        if item_callback is not None:
            item_callback(inspection, index, total)

    return scanned_items, all_cache
```

`inspect_project` при этом не менять. Полная проверка продолжает выполняться существующей функцией.

---

## 6. Расширить Excel-отчёт данными кэша

### Найти `PROJECT_REPORT_COLUMNS`

### Заменить его полностью на

```python
PROJECT_REPORT_COLUMNS = (
    "Project",
    "FolderSearchKey",
    "ExpectedFolder",
    "FoundFolders",
    "SelectedFolder",
    "WorkbookCandidates",
    "SelectedWorkbook",
    "WorkbookVersion",
    "CurrentFormVersion",
    "VersionStatus",
    "WorkbookModifiedAt",
    "InspectedAt",
    "InspectionSource",
    "ExcelLock",
    "SafeToRecreate",
    "BlockingStudies",
    "RecommendedAction",
    "SelectedAction",
    "ExecutionResult",
    "Issue",
)
```

### В `write_scan_report` найти `ws.append((` внутри цикла

Существующий кортеж заменить на:

```python
        ws.append((
            item.refresh_project,
            item.folder_search_key,
            item.expected_folder,
            "\n".join(item.found_folders),
            item.selected_folder,
            "\n".join(item.workbook_candidates),
            item.selected_workbook,
            item.workbook_version,
            release.form_version,
            item.version_status,
            item.workbook_modified_at,
            item.inspected_at,
            item.inspection_source,
            "YES" if item.excel_lock_present else "NO",
            "YES" if item.safe_to_recreate else "NO",
            ", ".join(item.blocking_studies),
            item.recommended_action,
            item.selected_action,
            item.execution_result,
            item.issue,
        ))
```

---

## 7. Заменить только функцию `run_scan`

### Границы замены

Заменить код начиная со строки:

```python
def run_scan(
```

и до следующего заголовка:

```python
# =============================================================================
# DB
```

### Новый код

```python
def run_scan(
    *,
    release: ReleaseInfo,
    projects_text: str,
    year: int,
    report_path: Path | None,
    force_all: bool = False,
) -> RunReport:
    progress("SCAN: Excel запускаться не будет")

    def on_item(item: ProjectInspection, index: int, total: int) -> None:
        if index == 1 or index % 50 == 0 or index == total:
            progress(
                f"SCAN: {index}/{total}; "
                f"{item.refresh_project}; source={item.inspection_source}"
            )

    inspections, _cache = run_cached_scan(
        release=release,
        projects_text=projects_text,
        year=year,
        force_all=force_all,
        item_callback=on_item,
    )

    if report_path is None:
        report_path = PROJECT_DIR / (
            f"project_scan_{year}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )

    written = write_scan_report(inspections, report_path, release)
    attention = sum(
        1
        for item in inspections
        if item.recommended_action != "OK"
    )
    from_cache = sum(
        1
        for item in inspections
        if item.inspection_source == "cache"
    )

    progress(f"SCAN завершён. Отчёт: {written}")
    return RunReport(
        ok=True,
        mode="scan",
        message=(
            f"Сканирование завершено. Проектов: {len(inspections)}, "
            f"из кэша: {from_cache}, требуют внимания: {attention}"
        ),
        form_version=release.form_version,
        addin_version=release.addin_version,
        processed=len(inspections),
        succeeded=len(inspections),
        failed=0,
        scan_report=str(written),
    )
```

---

## 8. Добавить выполнение смешанного плана `update/recreate`

### Место вставки

Найти полное окончание функции:

```python
def run_mutating_mode(...)
```

Следующий блок вставить после неё и до заголовка `# UI`.

```python
def run_execution_plan(
    *,
    release: ReleaseInfo,
    plan: dict[str, str],
    item_callback: Callable[[ProjectRunResult, int, int], None] | None = None,
) -> RunReport:
    normalized_plan: list[tuple[str, str]] = []

    for project, action in plan.items():
        normalized_project = normalize_project_number(project)
        normalized_action = action.strip().lower()
        if normalized_action not in {"update", "recreate"}:
            raise ValueError(
                f"Некорректное действие для {normalized_project}: {action}"
            )
        normalized_plan.append((normalized_project, normalized_action))

    if not normalized_plan:
        raise ValueError("Не выбрано ни одного действия")

    task_rows = read_external_task_rows(
        TASKS_WORKBOOK,
        TASKS_PARENT_SHEET,
        TASKS_PARENT_TABLE,
    )
    task_infos = build_task_project_infos(task_rows)
    projects_with_db_results = get_projects_with_db_results()

    app = xw.App(visible=True, add_book=False)
    results: list[ProjectRunResult] = []

    try:
        app.display_alerts = False
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False
        app.api.AlertBeforeOverwriting = False
        app.api.EnableEvents = True
        app.screen_updating = False

        progress("Подключаю установленную Excel-надстройку")
        addin = load_installed_addin(
            app,
            Path(release.addin_path),
            release.addin_version,
        )
        addin_name = str(addin.Name)
        disable_excel_prompts(app)

        total = len(normalized_plan)
        for index, (project, action) in enumerate(normalized_plan, start=1):
            progress("-" * 80)
            progress(
                f"PLAN: {index}/{total} {project}; "
                f"action={action.upper()}"
            )

            result = create_or_update_one_project(
                app=app,
                addin_name=addin_name,
                mode=action,
                refresh_project=project,
                task_info=task_infos.get(project),
                release=release,
                projects_with_db_results=projects_with_db_results,
            )
            results.append(result)

            if item_callback is not None:
                item_callback(result, index, total)

    finally:
        try:
            app.quit()
        except Exception:
            pass

    succeeded = sum(1 for item in results if item.ok)
    failed = len(results) - succeeded

    return RunReport(
        ok=failed == 0,
        mode="plan",
        message=(
            f"План выполнен. Успешно: {succeeded}, "
            f"ошибок/блокировок: {failed}"
        ),
        form_version=release.form_version,
        addin_version=release.addin_version,
        processed=len(results),
        succeeded=succeeded,
        failed=failed,
        projects=results,
    )
```

Не изменять `create_or_update_one_project`: перед непосредственным изменением он уже повторно проверяет книгу и запрещает небезопасное пересоздание.

---

## 9. Полностью заменить старый UI

### Границы замены

В разделе:

```python
# =============================================================================
# UI
# =============================================================================
```

полностью удалить функцию:

```python
def ask_ui_options(...)
```

Удалять нужно до следующего заголовка `# CLI / main`.

### Вместо неё вставить

```python
class ProjectFormsApp:
    def __init__(self, root, tk, ttk, filedialog, release: ReleaseInfo):
        self.root = root
        self.tk = tk
        self.ttk = ttk
        self.filedialog = filedialog
        self.release = release
        self.cache = load_scan_cache(release)
        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.busy = False

        root.title(
            f"Формы проектов ГТИ — форма {release.form_version}, "
            f"надстройка {release.addin_version}"
        )
        root.geometry("1500x820")
        root.minsize(1050, 620)

        self.projects_var = tk.StringVar()
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.status_var = tk.StringVar(value="Готово")

        self._build_controls()
        self._build_table()
        self._build_details()
        self._build_status_bar()
        self.show_cached_year()

        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll_events)

    def _build_controls(self) -> None:
        tk = self.tk
        ttk = self.ttk

        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Проекты:").grid(row=0, column=0, sticky="w")
        self.projects_entry = ttk.Entry(
            top,
            textvariable=self.projects_var,
            width=48,
        )
        self.projects_entry.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(6, 12),
        )

        ttk.Label(top, text="Год:").grid(row=0, column=2, sticky="w")
        self.year_entry = ttk.Entry(
            top,
            textvariable=self.year_var,
            width=8,
        )
        self.year_entry.grid(row=0, column=3, padx=(6, 12))

        ttk.Button(
            top,
            text="Показать кэш",
            command=self.show_cached_year,
        ).grid(row=0, column=4, padx=3)

        ttk.Button(
            top,
            text="Проверить изменения",
            command=self.scan_changed,
        ).grid(row=0, column=5, padx=3)

        ttk.Button(
            top,
            text="Пересканировать выделенные",
            command=self.force_scan_selected,
        ).grid(row=0, column=6, padx=3)

        ttk.Button(
            top,
            text="Пересканировать всё",
            command=self.force_scan_all,
        ).grid(row=0, column=7, padx=3)

        top.columnconfigure(1, weight=1)

        actions = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        actions.pack(fill="x")

        ttk.Button(
            actions,
            text="Выделить всё",
            command=self.select_all,
        ).pack(side="left", padx=3)

        ttk.Button(
            actions,
            text="Выделить устаревшие",
            command=self.select_outdated,
        ).pack(side="left", padx=3)

        ttk.Button(
            actions,
            text="Выделить доступные для пересоздания",
            command=self.select_recreatable,
        ).pack(side="left", padx=3)

        ttk.Separator(actions, orient="vertical").pack(
            side="left",
            fill="y",
            padx=8,
        )

        ttk.Button(
            actions,
            text="Назначить обновление/создание",
            command=lambda: self.assign_action("update"),
        ).pack(side="left", padx=3)

        ttk.Button(
            actions,
            text="Назначить пересоздание",
            command=lambda: self.assign_action("recreate"),
        ).pack(side="left", padx=3)

        ttk.Button(
            actions,
            text="Пропустить",
            command=lambda: self.assign_action(""),
        ).pack(side="left", padx=3)

        ttk.Button(
            actions,
            text="Выполнить план",
            command=self.execute_plan,
        ).pack(side="right", padx=3)

        ttk.Button(
            actions,
            text="Экспорт в Excel",
            command=self.export_excel,
        ).pack(side="right", padx=3)

        self._bind_entry_paste(self.projects_entry)
        self._bind_entry_paste(self.year_entry)

    def _build_table(self) -> None:
        ttk = self.ttk

        frame = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        frame.pack(fill="both", expand=True)

        columns = (
            "project",
            "version",
            "version_status",
            "modified",
            "source",
            "safe",
            "blocking",
            "recommended",
            "planned",
            "result",
        )

        self.tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
            selectmode="extended",
        )

        headings = {
            "project": "Проект",
            "version": "Версия",
            "version_status": "Статус версии",
            "modified": "Файл изменён",
            "source": "Источник",
            "safe": "Можно пересоздать",
            "blocking": "Блокирующие исследования",
            "recommended": "Рекомендация",
            "planned": "Выбранное действие",
            "result": "Результат",
        }
        widths = {
            "project": 95,
            "version": 80,
            "version_status": 150,
            "modified": 145,
            "source": 75,
            "safe": 120,
            "blocking": 230,
            "recommended": 220,
            "planned": 155,
            "result": 330,
        }

        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(
                column,
                width=widths[column],
                minwidth=60,
                stretch=column in {"blocking", "recommended", "result"},
            )

        vertical = ttk.Scrollbar(
            frame,
            orient="vertical",
            command=self.tree.yview,
        )
        horizontal = ttk.Scrollbar(
            frame,
            orient="horizontal",
            command=self.tree.xview,
        )
        self.tree.configure(
            yscrollcommand=vertical.set,
            xscrollcommand=horizontal.set,
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.tag_configure("ok", background="#e8f5e9")
        self.tree.tag_configure("attention", background="#fff8e1")
        self.tree.tag_configure("blocked", background="#ffebee")
        self.tree.tag_configure("planned", background="#e3f2fd")

        self.tree.bind("<<TreeviewSelect>>", self.show_details)
        self.tree.bind("<Control-KeyPress>", self._tree_control_key)

    def _build_details(self) -> None:
        ttk = self.ttk

        frame = ttk.LabelFrame(
            self.root,
            text="Подробности выбранного проекта",
            padding=6,
        )
        frame.pack(fill="x", padx=8, pady=(0, 8))

        self.details = self.tk.Text(
            frame,
            height=9,
            wrap="word",
            state="disabled",
        )
        self.details.pack(fill="x")

    def _build_status_bar(self) -> None:
        ttk = self.ttk

        frame = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        frame.pack(fill="x")

        ttk.Label(frame, textvariable=self.status_var).pack(
            side="left",
            fill="x",
            expand=True,
        )
        self.progress_bar = ttk.Progressbar(
            frame,
            mode="indeterminate",
            length=220,
        )
        self.progress_bar.pack(side="right")

    def _bind_entry_paste(self, widget) -> None:
        def paste(event):
            try:
                value = self.root.clipboard_get()
            except self.tk.TclError:
                return "break"

            try:
                event.widget.delete("sel.first", "sel.last")
            except self.tk.TclError:
                pass

            event.widget.insert("insert", value)
            return "break"

        def control_key(event):
            # Windows: физическая V имеет keycode 86 при любой раскладке.
            if event.keycode == 86:
                return paste(event)
            return None

        widget.bind("<Control-KeyPress>", control_key)
        widget.bind("<Shift-Insert>", paste)

    def _tree_control_key(self, event):
        # Windows: физическая A имеет keycode 65 при любой раскладке.
        if event.keycode == 65:
            self.select_all()
            return "break"
        return None

    def _read_year(self) -> int | None:
        try:
            return int(self.year_var.get().strip())
        except ValueError:
            self.status_var.set("Ошибка: год должен быть целым числом")
            return None

    def _project_visible_for_year(self, project: str, year: int) -> bool:
        return project.upper().startswith(f"{year % 100:02d}-F")

    def _row_values(self, item: ProjectInspection) -> tuple[str, ...]:
        action_names = {
            "": "",
            "update": "Обновить/создать",
            "recreate": "Пересоздать",
        }
        modified = item.workbook_modified_at.replace("T", " ")[:19]
        return (
            item.refresh_project,
            item.workbook_version,
            item.version_status,
            modified,
            item.inspection_source,
            "ДА" if item.safe_to_recreate else "НЕТ",
            ", ".join(item.blocking_studies),
            item.recommended_action,
            action_names.get(item.selected_action, item.selected_action),
            item.execution_result,
        )

    def _row_tag(self, item: ProjectInspection) -> str:
        if item.selected_action:
            return "planned"
        if item.recommended_action in HARD_BLOCKING_ACTIONS:
            return "blocked"
        if item.recommended_action != "OK":
            return "attention"
        return "ok"

    def _upsert_row(self, item: ProjectInspection) -> None:
        year = self._read_year()
        if year is None:
            return

        project = item.refresh_project
        if not self._project_visible_for_year(project, year):
            if self.tree.exists(project):
                self.tree.delete(project)
            return

        values = self._row_values(item)
        tags = (self._row_tag(item),)

        if self.tree.exists(project):
            self.tree.item(project, values=values, tags=tags)
        else:
            self.tree.insert(
                "",
                "end",
                iid=project,
                values=values,
                tags=tags,
            )

    def show_cached_year(self) -> None:
        if self.busy:
            return

        year = self._read_year()
        if year is None:
            return

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        shown = 0
        for project in sorted(self.cache):
            if self._project_visible_for_year(project, year):
                self._upsert_row(self.cache[project])
                shown += 1

        self.status_var.set(
            f"Показан кэш за {year} год: {shown} проектов"
        )

    def selected_projects(self) -> list[str]:
        return [str(value) for value in self.tree.selection()]

    def select_all(self) -> None:
        self.tree.selection_set(self.tree.get_children())

    def select_outdated(self) -> None:
        values = [
            project
            for project in self.tree.get_children()
            if self.cache[project].version_status == "OUTDATED_COMPATIBLE"
        ]
        self.tree.selection_set(values)

    def select_recreatable(self) -> None:
        values = [
            project
            for project in self.tree.get_children()
            if inspection_allows_recreate(self.cache[project])
        ]
        self.tree.selection_set(values)

    def assign_action(self, action: str) -> None:
        if self.busy:
            return

        selected = self.selected_projects()
        if not selected:
            self.status_var.set("Сначала выделите проекты")
            return

        assigned = 0
        skipped = 0

        for project in selected:
            item = self.cache[project]

            if action and item.recommended_action in HARD_BLOCKING_ACTIONS:
                skipped += 1
                continue

            if action == "recreate" and not inspection_allows_recreate(item):
                skipped += 1
                continue

            item.selected_action = action
            self._upsert_row(item)
            assigned += 1

        save_scan_cache(
            self.cache,
            self.release,
            last_scan={"kind": "ui-plan", "changedAt": iso_now()},
        )
        self.status_var.set(
            f"Действие назначено: {assigned}; пропущено: {skipped}"
        )

    def show_details(self, _event=None) -> None:
        selected = self.selected_projects()
        text = ""
        if selected:
            item = self.cache[selected[0]]
            text = json.dumps(asdict(item), ensure_ascii=False, indent=2)

        self.details.configure(state="normal")
        self.details.delete("1.0", "end")
        self.details.insert("1.0", text)
        self.details.configure(state="disabled")

    def _set_busy(self, value: bool, message: str = "") -> None:
        self.busy = value
        if message:
            self.status_var.set(message)
        if value:
            self.progress_bar.start(12)
        else:
            self.progress_bar.stop()

    def scan_changed(self) -> None:
        self._start_scan(force_all=False, selected_only=False)

    def force_scan_selected(self) -> None:
        self._start_scan(force_all=True, selected_only=True)

    def force_scan_all(self) -> None:
        self._start_scan(force_all=True, selected_only=False)

    def _start_scan(self, *, force_all: bool, selected_only: bool) -> None:
        if self.busy:
            return

        year = self._read_year()
        if year is None:
            return

        if selected_only:
            selected = self.selected_projects()
            if not selected:
                self.status_var.set("Не выбраны проекты для пересканирования")
                return
            projects_text = " ".join(selected)
        else:
            projects_text = self.projects_var.get().strip()

        cache_snapshot = {
            key: replace(value)
            for key, value in self.cache.items()
        }
        self._set_busy(True, "Сканирование запущено...")

        thread = threading.Thread(
            target=self._scan_worker,
            args=(
                year,
                projects_text,
                force_all,
                cache_snapshot,
            ),
            daemon=True,
        )
        thread.start()

    def _scan_worker(
        self,
        year: int,
        projects_text: str,
        force_all: bool,
        cache_snapshot: dict[str, ProjectInspection],
    ) -> None:
        try:
            def on_item(item, index, total):
                self.events.put(("scan_item", (item, index, total)))

            _scope, updated_cache = run_cached_scan(
                release=self.release,
                projects_text=projects_text,
                year=year,
                cache=cache_snapshot,
                force_all=force_all,
                item_callback=on_item,
            )
            self.events.put(("scan_done", updated_cache))
        except Exception as exc:
            logger.exception("Ошибка сканирования из UI")
            self.events.put(
                ("error", f"{type(exc).__name__}: {exc}")
            )

    def execute_plan(self) -> None:
        if self.busy:
            return

        plan = {
            project: item.selected_action
            for project, item in self.cache.items()
            if item.selected_action in {"update", "recreate"}
        }

        if not plan:
            self.status_var.set("План пуст: сначала назначьте действия")
            return

        cache_snapshot = {
            key: replace(value)
            for key, value in self.cache.items()
        }
        year = self._read_year()
        if year is None:
            return

        self._set_busy(
            True,
            f"Выполняется план: {len(plan)} проектов...",
        )

        thread = threading.Thread(
            target=self._execution_worker,
            args=(plan, cache_snapshot, year),
            daemon=True,
        )
        thread.start()

    def _execution_worker(
        self,
        plan: dict[str, str],
        cache_snapshot: dict[str, ProjectInspection],
        year: int,
    ) -> None:
        pythoncom = None
        try:
            try:
                import pythoncom as imported_pythoncom
                pythoncom = imported_pythoncom
                pythoncom.CoInitialize()
            except ImportError:
                pythoncom = None

            results_by_project: dict[str, ProjectRunResult] = {}

            def on_result(result, index, total):
                results_by_project[result.project] = result
                self.events.put(("execution_item", (result, index, total)))

            report = run_execution_plan(
                release=self.release,
                plan=plan,
                item_callback=on_result,
            )

            _scope, updated_cache = run_cached_scan(
                release=self.release,
                projects_text=" ".join(plan),
                year=year,
                cache=cache_snapshot,
                force_all=True,
            )

            for project, result in results_by_project.items():
                item = updated_cache.get(project)
                if item is None:
                    continue
                item.selected_action = ""
                item.execution_result = (
                    f"{result.action}: {result.message}"
                    if result.ok
                    else f"ОШИБКА: {result.message}"
                )

            save_scan_cache(
                updated_cache,
                self.release,
                last_scan={
                    "kind": "execution-plan",
                    "finishedAt": iso_now(),
                },
            )
            self.events.put(("execution_done", (report, updated_cache)))

        except Exception as exc:
            logger.exception("Ошибка выполнения плана из UI")
            self.events.put(
                ("error", f"{type(exc).__name__}: {exc}")
            )
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def export_excel(self) -> None:
        if self.busy:
            return

        year = self._read_year()
        if year is None:
            return

        visible = [
            self.cache[project]
            for project in self.tree.get_children()
        ]
        if not visible:
            self.status_var.set("Нет строк для экспорта")
            return

        path = self.filedialog.asksaveasfilename(
            title="Сохранить отчёт сканирования",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=(
                f"project_scan_{year}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            ),
        )
        if not path:
            return

        try:
            write_scan_report(visible, Path(path), self.release)
            self.status_var.set(f"Отчёт сохранён: {path}")
        except Exception as exc:
            logger.exception("Ошибка экспорта отчёта")
            self.status_var.set(f"Ошибка экспорта: {exc}")

    def poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()

                if kind == "scan_item":
                    item, index, total = payload
                    self.cache[item.refresh_project] = item
                    self._upsert_row(item)
                    self.status_var.set(
                        f"Сканирование: {index}/{total}; "
                        f"{item.refresh_project}; "
                        f"источник={item.inspection_source}"
                    )

                elif kind == "scan_done":
                    self.cache = payload
                    self._set_busy(False, "Сканирование завершено")
                    self.show_cached_year()

                elif kind == "execution_item":
                    result, index, total = payload
                    item = self.cache.get(result.project)
                    if item is not None:
                        item.execution_result = (
                            f"{result.action}: {result.message}"
                            if result.ok
                            else f"ОШИБКА: {result.message}"
                        )
                        self._upsert_row(item)
                    self.status_var.set(
                        f"Выполнение: {index}/{total}; "
                        f"{result.project}; {result.action}"
                    )

                elif kind == "execution_done":
                    report, updated_cache = payload
                    self.cache = updated_cache
                    self._set_busy(False, report.message)
                    self.show_cached_year()
                    self.status_var.set(report.message)

                elif kind == "error":
                    self._set_busy(False, f"Ошибка: {payload}")

        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.poll_events)

    def close(self) -> None:
        if self.busy:
            self.status_var.set(
                "Нельзя закрыть окно во время выполняющейся операции"
            )
            return
        self.root.destroy()


def launch_project_forms_ui(release: ReleaseInfo) -> int:
    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    ProjectFormsApp(root, tk, ttk, filedialog, release)
    root.mainloop()
    return 0
```

## 10. Изменить CLI-параметры

### В `build_parser` после `--report` добавить

```python
    parser.add_argument(
        "--force-scan",
        action="store_true",
        help="Не использовать JSON-кэш при сканировании",
    )
```

### Заменить сигнатуру `execute`

Было:

```python
def execute(*, mode: str, projects_text: str, year: int, report_path: str) -> RunReport:
```

Стало:

```python
def execute(
    *,
    mode: str,
    projects_text: str,
    year: int,
    report_path: str,
    force_scan: bool = False,
) -> RunReport:
```

### Внутри `execute`, в вызов `run_scan`, добавить

```python
            force_all=force_scan,
```

Полный вызов должен выглядеть так:

```python
        return run_scan(
            release=release,
            projects_text=projects_text,
            year=year,
            report_path=Path(report_path) if report_path else None,
            force_all=force_scan,
        )
```

---

## 11. Заменить только функцию `main`

### Границы замены

Заменить от:

```python
def main(argv: list[str] | None = None) -> int:
```

до строки перед:

```python
if __name__ == "__main__":
```

### Новый код

```python
def main(argv: list[str] | None = None) -> int:
    args_list = list(sys.argv[1:] if argv is None else argv)

    # Без параметров запускается новый интерфейс управления проектами.
    if not args_list:
        try:
            release = preflight_release()
            if not BASE_PROJECTS_DIR.exists():
                raise FileNotFoundError(
                    f"Корневая папка проектов не найдена: {BASE_PROJECTS_DIR}"
                )
            return launch_project_forms_ui(release)
        except Exception as exc:
            logger.exception("Ошибка запуска UI")
            print(
                json.dumps(
                    {
                        "ok": False,
                        "message": f"{type(exc).__name__}: {exc}",
                    },
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 1

    parser = build_parser()
    args = parser.parse_args(args_list)

    if not args.mode:
        parser.error("--mode обязателен при запуске с параметрами")

    if args.mode == "recreate" and not args.projects.strip():
        parser.error("для режима recreate нужно явно указать --projects")

    try:
        report = execute(
            mode=args.mode,
            projects_text=args.projects,
            year=args.year,
            report_path=args.report,
            force_scan=args.force_scan,
        )
    except Exception as exc:
        logger.exception("Критическая ошибка")
        report = RunReport(
            ok=False,
            mode=args.mode,
            message=(
                f"{type(exc).__name__}: {exc}. "
                f"Подробности в логе: {LOG_PATH}"
            ),
            form_version="",
            addin_version="",
            processed=0,
            succeeded=0,
            failed=1,
        )

    print(report_to_json(report))
    return 0 if report.ok else 1
```

Строки в конце файла оставить без изменений:

```python
if __name__ == "__main__":
    sys.exit(main())
```

---

## 12. Что точно не менять

Не переписывать следующие функции:

- `preflight_release`;
- `read_external_task_rows`;
- `build_task_project_infos`;
- `find_project_folders`;
- `find_project_workbooks`;
- `read_study_states` — оставить актуальную логику совместимости DLST/STABLE;
- `inspect_project`;
- `get_projects_with_db_results`;
- `build_new_workbook`;
- `update_existing_workbook`;
- `validate_existing_workbook_for_automation`;
- `create_or_update_one_project`;
- `run_mutating_mode`.

Новый UI вызывает уже проверенную существующую бизнес-логику.

---

## 13. Порядок проверки после внесения изменений

### 13.1. Синтаксис

```powershell
python -m py_compile create_project_forms.py
```

### 13.2. Первый запуск

1. Запустить скрипт без аргументов.
2. Нажать `Проверить изменения`.
3. Убедиться, что появился `project_scan_cache.json`.
4. Проверить, что первый проход показывает `inspection_source = fresh`.

### 13.3. Проверка кэша

1. Закрыть скрипт.
2. Запустить снова.
3. Нажать `Проверить изменения`.
4. У неизменённых книг должен быть `inspection_source = cache`.
5. Второй проход должен выполняться значительно быстрее первого.

### 13.4. Проверка изменения файла

1. Открыть одну тестовую форму.
2. Изменить безопасную ячейку и сохранить книгу.
3. Закрыть Excel.
4. Запустить `Проверить изменения`.
5. Только изменённая книга должна получить `inspection_source = fresh`.

### 13.5. Проверка lock-файла

1. Открыть тестовую форму в Excel.
2. Запустить проверку изменений.
3. Проект должен получить `BLOCKED_WORKBOOK_OPEN`.
4. Закрыть Excel и повторить проверку.
5. Проект должен быть полностью перечитан и получить нормальный статус.

### 13.6. Проверка массового выбора

1. Выделить несколько строк через `Ctrl` и `Shift`.
2. Назначить `Обновить/создать`.
3. Убедиться, что действие появилось только у выделенных строк.
4. Попробовать назначить пересоздание заблокированному проекту — действие не должно назначиться.

### 13.7. Проверка выполнения

Сначала использовать 1–2 тестовых проекта. После выполнения проверить:

- результат появился в колонке `Результат`;
- назначенное действие очищено;
- изменённая/созданная книга заново просканирована;
- JSON содержит новое время изменения и новый статус;
- при пересоздании старая форма оказалась в папке `архив`.

---

## 14. Команды CLI после изменений

Обычное сканирование с использованием кэша:

```powershell
python create_project_forms.py --mode scan --year 2026
```

Полное сканирование без кэша:

```powershell
python create_project_forms.py --mode scan --year 2026 --force-scan
```

Принудительная проверка отдельных проектов:

```powershell
python create_project_forms.py --mode scan --projects "26-F001 26-F002" --force-scan
```

Старые команды `update` и `recreate` продолжают работать без изменений.
