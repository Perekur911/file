# Аудит готовности релиза DLST и Stable

Дата проверки: 10.09.2026

Проверены предоставленные архивы:

- `python.zip`;
- `addin_source.zip`;
- `database.zip`.

Исходные архивы и надстройка не изменялись.

## Итог

**Текущая версия не готова к релизу.** До сборки нужно исправить четыре блокирующие ошибки и выполнить финальную проверку в Excel на настоящей книге формы.

| Приоритет | Проблема | Последствие |
|---|---|---|
| Блокер | `005_add_stable_table.sql` содержит 3 колонки и только 2 значения в `INSERT` | Миграция завершается ошибкой `2 values for 3 columns` |
| Блокер | `004_add_dlst_tables.sql` не регистрирует миграцию | Невозможно достоверно определить, была ли создана схема DLST |
| Блокер | В `schema.py` дважды объявлен `Stable.analyseNumber` | Python формирует дублирующийся столбец и использует второе, неверное описание `TEXT` |
| Блокер | Отмена Stable передаёт тип задания `Stable`, тогда как сохранение и существующая БД используют `STABLE` | Для одного задания создаются две независимые цепочки статусов |
| Обязательно | `EndMacro` не защищает `Task_mix` и обращается к несуществующему `Project_LIF` вместо `ProjectLIF` | После макросов часть листов может остаться без защиты |
| Обязательно | Ошибки `Save_Stable` видны только в Immediate Window | Пользователь получает отказ без объяснения причины |
| Решение до релиза | Кнопки добавления/удаления формы видны для DLST, но обе процедуры являются заглушками | Ribbon обещает функцию, которая не работает |
| Существующая ошибка | Обёртки CVD вызывают отсутствующие `InsertListCVD2` и `DeleteLastListCVD2` | Кнопки CVD заканчиваются ошибкой «макрос не найден» |

## 1. Исправить миграцию DLST

Файл: `database/database/migrations/004_add_dlst_tables.sql`

Перед `COMMIT;` вставить:

```sql
INSERT INTO schema_migrations (
    schema_version,
    migration_name
)
VALUES (
    '2.1.0',
    '004_add_dlst_tables'
);
```

Если версия схемы текущего релиза в `project_version.json` отличается от `2.1.0`, здесь должна быть именно версия компонента `database_schema` текущего релиза.

## 2. Исправить миграцию Stable

Файл: `database/database/migrations/005_add_stable_table.sql`

Полностью заменить текущий блок `INSERT INTO schema_migrations ... VALUES ...` на:

```sql
INSERT INTO schema_migrations (
    schema_version,
    migration_name
)
VALUES (
    '2.1.0',
    '005_add_stable_table'
);
```

Не следует вручную задавать `migration_id = 5`: существующая миграция `003` получает идентификатор автоматически, а файлы `002` и `003` являются двумя частями одной миграции. Автоматический следующий `INTEGER PRIMARY KEY` сохраняет фактический порядок применения.

Обе исправленные миграции были смоделированы в памяти. Результат проверки:

```text
003_op_vliqphase_swap -> migration_id 2, schema_version 2.0.0
004_add_dlst_tables   -> migration_id 3, schema_version 2.1.0
005_add_stable_table -> migration_id 4, schema_version 2.1.0
PRAGMA quick_check   -> ok
PRAGMA foreign_key_check -> нарушений нет
```

## 3. Удалить повторное поле Stable в Python

Файл: `python/python/schema.py`

В блоке `STABLE_RESULTS` уже есть правильное объявление:

```python
_col(
    "analyseNumber",
    "analyseNumber",
    typ="INTEGER",
    value_type="integer",
    required=True,
),
```

Ниже, после `PVTcell`, удалить ошибочный повтор:

```python
_col("analyseNumber", "analyseNumber", typ="TEXT", value_type="text"),
```

После удаления схема Stable должна содержать ровно 17 уникальных полей:

```text
resultIdStable, TaskId, sampleCode, analyseNumber,
dateStart, timeStart, dateEnd, timeEnd, operator, PVTcell,
temperature, pressure, pressureUnit, pressureAbs,
pressureMPaAbs, LiquidVolume, dateTimeSync
```

### Рекомендуемая защита от повторения ошибки

В класс `TableSchema` можно добавить проверку уникальности имён. Метод вставляется внутрь `TableSchema`, перед свойством `pk`:

```python
def __post_init__(self) -> None:
    db_names = [column.db_name.casefold() for column in self.columns]
    excel_names = [column.xlsx_name.casefold() for column in self.columns]

    duplicate_db = sorted({name for name in db_names if db_names.count(name) > 1})
    duplicate_excel = sorted({name for name in excel_names if excel_names.count(name) > 1})

    if duplicate_db:
        raise ValueError(
            f"{self.logical_name}: повторяются поля БД: {', '.join(duplicate_db)}"
        )

    if duplicate_excel:
        raise ValueError(
            f"{self.logical_name}: повторяются поля Excel: {', '.join(duplicate_excel)}"
        )
```

Это улучшение не заменяет удаление текущего дубля.

## 4. Исправить тип задания при отмене Stable

Файл: `addin_source/modules/Stable.bas`

Удалить константу:

```vb
Private Const STABLE_TASK_TYPE As String = "Stable"
```

В `cancelCurrentTaskStable` заменить:

```vb
syncResult = SetTaskStatusOnly( _
    CLng(taskCell.value), _
    STABLE_TASK_TYPE, _
    "Отмена", _
    comment _
)
```

на:

```vb
syncResult = SetTaskStatusOnly( _
    CLng(taskCell.value), _
    STABLE_STUDY_CODE, _
    "Отмена", _
    comment _
)
```

Проверка предоставленной тестовой БД показала, что существующий тип называется именно `STABLE`.

## 5. Вернуть защиту всех служебных листов

Файл: `addin_source/modules/Functions.bas`

В `EndMacro` заменить блок служебных листов на:

```vb
PrepareSheet wb, "Task", True, enableSheetCalculation
PrepareSheet wb, "Task_mix", True, enableSheetCalculation
PrepareSheet wb, "All_samples", True, enableSheetCalculation
PrepareSheet wb, "ProjectLIF", True, enableSheetCalculation
PrepareSheet wb, "StatePC", True, enableSheetCalculation
PrepareSheet wb, "Пересчет", True, enableSheetCalculation
```

Сейчас `Task_mix` пропущен, а `Project_LIF` не совпадает с именем листа `ProjectLIF`, используемым в `StartMacro` и в форме.

## 6. Показывать пользователю ошибки Stable

Файл: `addin_source/modules/Stable.bas`

В `Save_Stable`, ветка `DATA_DELETED`, перед `Exit Function` добавить:

```vb
MsgBox _
    "Удалены данные ранее синхронизированной формы Stable. " & _
    "Вызовите загрузку данных из БД.", _
    vbExclamation
```

Обработчик `ErrHandler` заменить на:

```vb
ErrHandler:
    Debug.Print "Save_Stable: " & Err.Number & ": " & Err.Description

    MsgBox _
        "Ошибка Save_Stable: " & _
        Err.Number & ": " & Err.Description, _
        vbExclamation

    Save_Stable = MACRO_FAILED
End Function
```

## 7. Определиться с добавлением и удалением страниц DLST

В `DLST.bas` процедуры `AddFormDLST` и `DeleteLastFormDLST` пока являются заглушками. При этом общие кнопки Ribbon для активного исследования видимы и вызывают эти процедуры.

До релиза нужно выбрать один вариант:

1. реализовать безопасное копирование/удаление блока из 18 столбцов со всеми объединениями, формулами и объектами;
2. временно скрыть кнопки для DLST, пока страницы создаются только заранее в шаблоне.

Без настоящей книги формы безопасно реализовать вариант 1 нельзя. Для временного варианта 2 в `addin_source/ribbon/package/customUI/customUI14.xml` добавить кнопкам:

```xml
<button id="btnAddCurrentForm"
        ...
        getVisible="Ribbon_AddCurrentFormVisible"
        onAction="Ribbon_StudyAction"/>

<button id="btnDeleteCurrentForm"
        ...
        getVisible="Ribbon_DeleteCurrentFormVisible"
        onAction="Ribbon_StudyAction"/>
```

В `modRibbonPVT.bas` добавить:

```vb
Public Sub Ribbon_AddCurrentFormVisible( _
    ByVal control As IRibbonControl, _
    ByRef returnedVal)

    Dim studyName As String
    studyName = CurrentStudyName()

    returnedVal = IsPvtResultsWorkbook(ActiveWorkbook) _
                  And Len(studyName) > 0 _
                  And studyName <> "DLST" _
                  And studyName <> "CVD"
End Sub

Public Sub Ribbon_DeleteCurrentFormVisible( _
    ByVal control As IRibbonControl, _
    ByRef returnedVal)

    Dim studyName As String
    studyName = CurrentStudyName()

    returnedVal = IsPvtResultsWorkbook(ActiveWorkbook) _
                  And Len(studyName) > 0 _
                  And studyName <> "DLST" _
                  And studyName <> "CVD" _
                  And studyName <> "SSF"
End Sub
```

Здесь также скрыты уже существующие нерабочие действия CVD и удаление SSF.

## 8. Исправить старые вызовы CVD в Ribbon

Файл: `addin_source/modules/wrappers.bas`

Текущие вызовы ссылаются на отсутствующие процедуры. Заменить:

```vb
CallMacro "InsertListCVD2", "CVD"
CallMacro "DeleteLastListCVD2", "CVD"
```

на существующие точки входа:

```vb
CallMacro "AddFormCVD", "CVD"
CallMacro "DeleteLastListCVD", "CVD"
```

Эти существующие процедуры CVD тоже пока являются заглушками. Поэтому одной замены достаточно только для устранения ошибки «макрос не найден»; для рабочего добавления/удаления следует реализовать процедуры либо скрыть кнопки способом из раздела 7.

## Что уже сделано правильно

- `DLST_FORM_COLUMN_COUNT = 18` — соответствует форме с двумя узкими столбцами-разделителями.
- DLST использует имена только первой формы `DLST_1_*`, следующие формы вычисляются через смещение.
- Stable использует имена только первой формы `Stable_1_*`, следующие формы вычисляются через смещение.
- `DLSepTest` корректно преобразуется в код исследования `DLST` в Ribbon, событиях и разблокировке листов.
- DLST зарегистрирован в `schema.py`, `StudySync.bas`, `wrappers.bas`, `modMetadata.bas`, `modRibbonPVT.bas`, `clsAppEvents.cls` и `PowerQuery.bas`.
- Stable зарегистрирован в тех же основных цепочках.
- Имена листов и таблиц согласованы:
  - `DLSepTest`, `DLSepTest_results`, `DLSepTest_sourceData`;
  - `DLST_results`, `DLST_sourceData`;
  - `Stable`, `Stable_results`.
- DLST parent: Python и SQL содержат одинаковые 21 поле и одинаковые типы.
- DLST child: Python и SQL содержат одинаковые 43 поля, включая латинское `volumetricCoefficient`.
- После удаления повторного `analyseNumber` Stable: Python и SQL будут содержать одинаковые 17 полей и типы.
- `density20` используется одинаково в Python, SQL, таблице и VBA.
- Уникальность `(sampleCode, analyseNumber)` создана для DLST и Stable.
- `sampleCodeLPh` и `sampleCodeGPh` имеют отдельные уникальные ограничения.
- Пустые шифры дочерних проб превращаются Python-кодом в `NULL`, поэтому SQLite допускает несколько пустых строк; повтор непустого шифра отклоняется.
- Для DLST действует внешний ключ с `ON DELETE CASCADE` и уникальность `(resultIdDLSP, step)`.
- Экспорт надстройки согласован с manifest: 54 из 54 компонентов присутствуют, лишних/потерянных файлов нет.
- Все ссылки VBA, зафиксированные в manifest, помечены как исправные.
- Ribbon XML корректно разбирается.
- Все Python-файлы прошли `compileall` без синтаксических ошибок.

## Что нельзя подтвердить по предоставленным архивам

В поставке нет самой актуальной книги формы и собранной XLAM, поэтому статический аудит не подтверждает:

- фактические адреса `DLST_1_*` и `Stable_1_*`;
- реальное смещение Stable в 6 столбцов и ячейку страницы `D16`;
- формулы, объединения, checkbox и защиту листов после копирования страницы;
- соответствие заголовков умных таблиц книге;
- компиляцию VBA через `Debug -> Compile VBAProject`;
- версии пользовательских свойств формы и надстройки.

Также не предоставлен `project_version.json`. Поэтому невозможно проверить, что версии `form`, `addin`, `python` и `database_schema` совпадают, а последняя запись `schema_migrations` соответствует релизу.

Предоставленный `database/test/results.db` является старой тестовой фикстурой: в нём нет таблиц DLST/Stable, а последняя запись `schema_migrations` — `001_baseline / 1.0.0`. Не использовать этот файл как доказательство готовности новой схемы; после исправления миграций обновить тестовую копию либо прогнать миграции на отдельной копии актуальной БД.

## Финальный smoke-тест перед выпуском

Выполнять только на копии формы и копии БД.

1. Исправить четыре блокера выше.
2. Применить `004`, затем `005`; обе операции должны завершиться без ошибок.
3. Проверить БД:

```sql
SELECT migration_id, schema_version, migration_name, applied_at
FROM schema_migrations
ORDER BY migration_id;

PRAGMA foreign_key_check;
PRAGMA quick_check;
```

Ожидается: последняя миграция `005_add_stable_table`, версия соответствует `project_version.json`, `foreign_key_check` пуст, `quick_check = ok`.

4. Открыть актуальную книгу с подключённой новой надстройкой и выполнить `Debug -> Compile VBAProject`.
5. Обновить проект, содержащий минимум одно задание `DLST` и одно `Stable`.
6. Проверить автоматическое заполнение `TaskId`, `sampleCode`, `analyseNumber` на первой и второй формах.
7. Stable: сохранить, загрузить, изменить, повторно сохранить, валидировать и отменить задание. В `TaskStatus.task_type` во всех случаях должно быть только `STABLE`.
8. DLST: заполнить обычный step и `Last Step`, создать LPh/GPh, сохранить, загрузить и валидировать.
9. Удалить данные уже синхронизированной формы и убедиться, что сохранение блокируется понятным сообщением.
10. Проверить отказы БД: повтор `(sampleCode, analyseNumber)`, повтор `sampleCodeLPh`, повтор `sampleCodeGPh`.
11. После каждого Ribbon-макроса проверить, что `Task_mix` и `ProjectLIF` снова защищены.
12. Закрыть и снова открыть книгу; проверить хеши/статусы синхронизации и повторную загрузку.
13. На полном репозитории запустить релизный preflight `git_commit.py`. В текущей среде его запуск невозможен из-за отсутствующих `dulwich`, `xlwings` и полного `project_version.json`; это ограничение среды аудита, а не результат проверки рабочего виртуального окружения.

## Решение о выпуске

Разрешать релиз только после выполнения разделов 1–6, выбора варианта в разделе 7 и успешного smoke-теста. После этих исправлений статические контракты DLST и Stable выглядят согласованными.
