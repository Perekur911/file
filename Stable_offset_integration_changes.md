# Подключение исследования Stable

Готовый модуль: `Stable.bas`.

Код надстройки этим комплектом не изменяется автоматически. Ниже перечислены
все вставки, которые нужно сделать после импорта модуля в проект VBA.

## 1. Требования к книге формы

Должны существовать:

- лист формы `Stable`;
- лист результатов `Stable_results`;
- умная таблица `Stable_results` на одноимённом листе;
- только первая форма с именованными диапазонами `Stable_1_*`.

Заголовки таблицы `Stable_results` должны быть такими (регистр сохранён):

```text
resultIdStable
TaskId
sampleCode
dateStart
timeStart
dateEnd
timeEnd
operator
PVTcell
temperature
pressure
pressureUnit
pressureAbs
pressureMPaAbs
LiquidVolume
dateTimeSync
```

Имена нужны только для первой формы и для всех перечисленных полей, кроме
`dateTimeSync`. Например:

```text
Stable_1_resultIdStable
Stable_1_TaskId
Stable_1_sampleCode
...
Stable_1_LiquidVolume
```

Имена должны быть уровня книги. Адрес поля формы №2 и далее модуль получает
сдвигом базового диапазона `Stable_1_*` вправо на
`(номер формы - 1) * STABLE_FORM_COLUMN_COUNT` столбцов. Создавать
`Stable_2_*`, `Stable_3_*` и другие имена не нужно.

Если такие имена уже созданы, после перехода на новый модуль их можно удалить:
код их больше не читает и при добавлении формы не создаёт.

Фактическая ячейка `pressureAbs` каждой формы является связанной ячейкой её
чекбокса.

В начале `Stable.bas` проверить четыре константы геометрии:

```vb
Private Const STABLE_FORM_COLUMN_COUNT As Long = 3
Private Const STABLE_FIRST_FORM_FIRST_COL As Long = 3
Private Const STABLE_PAGE_ROW As Long = 3
Private Const STABLE_PAGE_FIRST_COL As Long = 3
Private Const STABLE_MIN_FORM_COUNT As Long = 2
```

- `STABLE_FORM_COLUMN_COUNT` — ширина одной формы в столбцах;
- `STABLE_FIRST_FORM_FIRST_COL` — первый столбец первой формы;
- `STABLE_PAGE_ROW` — строка ячейки с номером страницы;
- `STABLE_PAGE_FIRST_COL` — столбец ячейки с номером первой страницы;
- `STABLE_MIN_FORM_COUNT` — сколько начальных форм запрещено удалять.

Последняя форма определяется по безымянным ячейкам номера страницы. Модуль
начинает с `Cells(STABLE_PAGE_ROW, STABLE_PAGE_FIRST_COL)` и движется вправо с
шагом `STABLE_FORM_COLUMN_COUNT`; первая пустая ячейка завершает список форм.

## 2. `wrappers.bas`

Во всех пяти списках исследований добавить строку `"STABLE"`:

- `Load_All_wrap`;
- `Save_All_wrap`;
- `Validate_All_wrap`;
- `Refresh_Project_wrap`;
- `silentRefresh_Project_wrap`.

Добавить процедуры:

```vb
Public Sub Add_Stable_Form_wrap()
    CallMacro "AddFormStable", "STABLE"
End Sub

Public Sub Delete_Stable_Form_wrap()
    CallMacro "DeleteLastListStable", "STABLE"
End Sub

Public Sub Cancel_Stable_Task_wrap()
    CallMacroEx _
        "cancelCurrentTaskStable", _
        REFRESH_TASKS_AFTER, _
        "STABLE"
End Sub

Public Sub Save_Stable_wrap()
    CallMacroEx _
        "Save_Stable", _
        REFRESH_TASKS_AFTER, _
        "STABLE"
End Sub

Public Sub Load_Stable_wrap()
    CallMacro "Load_Stable", "STABLE"
End Sub

Public Sub Validate_Stable_wrap()
    CallMacroEx _
        "Validate_Stable", _
        REFRESH_TASKS_AFTER, _
        "STABLE"
End Sub
```

В `Load_All_results_to_forms_Auto`, рядом с блоком OPOH, добавить:

```vb
    ' STABLE
    feedback = Load_Stable_results_to_forms( _
        mode:="sequential", _
        silence:=True _
    )

    AddLoadFeedback _
        reportText, _
        hasWarnings, _
        hasErrors, _
        "STABLE", _
        feedback
```

## 3. `modRibbonPVT.bas`

В `CurrentStudyName` добавить `"STABLE"`:

```vb
Case "OP", "OPOH", "AP", "GC", "SSF", "GOR", "BP", _
     "REC", "EMV", "JOIN", "CCE", "CVD", "STABLE"
```

В `RunStudyWrapper` добавить:

```vb
        Case "SAVE_STABLE":          Save_Stable_wrap
        Case "LOAD_STABLE":          Load_Stable_wrap
        Case "VALIDATE_STABLE":      Validate_Stable_wrap
        Case "ADD_FORM_STABLE":      Add_Stable_Form_wrap
        Case "DELETE_FORM_STABLE":   Delete_Stable_Form_wrap
        Case "CANCEL_TASK_STABLE":   Cancel_Stable_Task_wrap
```

Ribbon XML менять не требуется: кнопки текущего исследования уже используют
динамический диспетчер.

## 4. `StudySync.bas`

Добавить `STABLE` в строку исследований по умолчанию в `RunStudySync`:

```vb
studies = "OP AP GC SSF GOR BP REC EMV OPOH CVD CCE JOIN STABLE"
```

Добавить его в `DefaultStudies`:

```vb
DefaultStudies = Array( _
    "OP", "OPOH", "GC", "BP", "AP", "GOR", "SSF", _
    "REC", "EMV", "JOIN", "CCE", "CVD", "STABLE" _
)
```

В `LoadStudyForms` добавить:

```vb
        Case "STABLE"
            LoadStudyForms = _
                Load_Stable_results_to_forms(formFillMode)
```

В `IsKnownStudyCode` добавить `"STABLE"` в общий `Case`.

В `Save_All`, рядом с OPOH, добавить:

```vb
    currentResult = Save_Stable()

    If currentResult = MACRO_FAILED Then
        Save_All = MACRO_FAILED
        Exit Function
    End If

    If currentResult = MACRO_CHANGED Then
        totalResult = MACRO_CHANGED
    End If
```

В `Validate_All` добавить аналогичный блок:

```vb
    currentResult = Validate_Stable()

    If currentResult = MACRO_FAILED Then
        Validate_All = MACRO_FAILED
        Exit Function
    End If

    If currentResult = MACRO_CHANGED Then
        totalResult = MACRO_CHANGED
    End If
```

## 5. `modMetadata.bas`

Добавить константы рядом с соответствующими группами:

```vb
Public Const PROP_STABLE_DATA_HASH As String = _
    "StableDataHash"

Public Const PROP_STABLE_DATA_STATE As String = _
    "StableDataState"

Public Const PROP_STABLE_LAST_WORKFLOW_STAGE As String = _
    "StableLastWorkflowStage"
```

В `RefreshAllStudySyncStatuses` добавить:

```vb
UpdateStudySyncStatus _
    wb, "STABLE", PROP_STABLE_DATA_HASH, PROP_STABLE_DATA_STATE
```

В `RefreshChangedStudySyncStatuses` добавить:

```vb
RefreshStudyIfNeeded _
    wb, "STABLE", PROP_STABLE_DATA_HASH, PROP_STABLE_DATA_STATE
```

В `GetCurrentStudyHash` добавить:

```vb
        Case "STABLE"
            GetCurrentStudyHash = CalculateCurrentStableHash()
```

В `GetUnsyncedStudiesText` добавить `"STABLE"` в `studyNames` и
`PROP_STABLE_DATA_STATE` в ту же позицию массива `stateProperties`. Длины и
порядок двух массивов должны совпадать.

В `SetStudyBaselineAfterLoad` добавить:

```vb
        Case "STABLE"
            hashProperty = PROP_STABLE_DATA_HASH
            stateProperty = PROP_STABLE_DATA_STATE
```

В `SetAllStudyBaselinesAfterLoad` добавить:

```vb
SetStudyBaselineAfterLoad wb, "STABLE"
```

## 6. `clsAppEvents.cls`

В `GetStudyCodeFromSheet` добавить `"STABLE"` в общий `Case`. Это включает
автоматическую отметку исследования как изменённого при правке листа формы.

## 7. `PowerQuery.bas`

`Refresh_Project` распределяет задания по формам, поэтому Stable нужно добавить
и сюда.

В объявление листов добавить `Stable_list As Worksheet`, затем назначить:

```vb
Set Stable_list = ActiveWorkbook.Worksheets("Stable")
```

В объявления счётчиков добавить `countStable As Integer`, а в объявления
свободных страниц — `emptyStablePage As Integer`.

В цепочку обработки `taskType`, рядом с OPOH, добавить:

```vb
        ElseIf taskType = "Stable" Then

            Stable_list.Visible = xlSheetVisible
            countStable = countStable + 1

            Dim existingStableForm As Long
            existingStableForm = _
                FindStableFormByTaskId(cellX.value)

            If existingStableForm = 0 Then

                emptyStablePage = FindEmptyFormStable()

                Dim stableTaskRange As Range
                Dim stableSampleRange As Range

                Set stableTaskRange = _
                    GetStableFormRange(emptyStablePage, "TaskId")

                Set stableSampleRange = _
                    GetStableFormRange(emptyStablePage, "sampleCode")

                With stableTaskRange.MergeArea
                    .value = cellX.value
                    .Interior.pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark2
                    .Interior.TintAndShade = _
                        -9.99481185338908E-02
                    .Interior.PatternTintAndShade = 0
                    .Locked = True
                End With

                With stableSampleRange.MergeArea
                    .value = _
                        cellX.Offset(0, offsetToSampleCode).value
                    .Interior.pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark2
                    .Interior.TintAndShade = _
                        -9.99481185338908E-02
                    .Interior.PatternTintAndShade = 0
                    .Locked = True
                End With

            End If
```

`task_table.bas` уже знает тип `Stable`, там изменение не требуется.

## 8. Python-схема

Добавить описание таблицы:

```python
STABLE_RESULTS = TableSchema(
    logical_name="STABLE_RESULTS",
    db_table="Stable_results",
    sheet_name="Stable_results",
    excel_table_name="Stable_results",
    status_datetime_fields=("dateStart", "timeStart"),
    columns=[
        _col("resultIdStable", "resultIdStable", typ="INTEGER", value_type="integer", pk=True),
        _col("TaskId", "TaskId", typ="INTEGER", value_type="integer"),
        _col("sampleCode", "sampleCode", typ="TEXT", value_type="text", required=True),
        _col("dateStart", "dateStart", typ="TEXT", value_type="date"),
        _col("timeStart", "timeStart", typ="TEXT", value_type="time"),
        _col("dateEnd", "dateEnd", typ="TEXT", value_type="date"),
        _col("timeEnd", "timeEnd", typ="TEXT", value_type="time"),
        _col("operator", "operator", typ="TEXT", value_type="text"),
        _col("PVTcell", "PVTcell", typ="TEXT", value_type="text"),
        _col("temperature", "temperature", typ="REAL", value_type="real"),
        _col("pressure", "pressure", typ="REAL", value_type="real"),
        _col("pressureUnit", "pressureUnit", typ="TEXT", value_type="text"),
        _col("pressureAbs", "pressureAbs", typ="INTEGER", value_type="boolean"),
        _col("pressureMPaAbs", "pressureMPaAbs", typ="REAL", value_type="real"),
        _col("LiquidVolume", "LiquidVolume", typ="REAL", value_type="real"),
        _col("dateTimeSync", "dateTimeSync", typ="TEXT", value_type="datetime"),
    ],
)
```

В словарь `STUDIES` добавить:

```python
"STABLE": StudySchema("STABLE", STABLE_RESULTS, None),
```

## 9. Миграция SQLite

```sql
CREATE TABLE Stable_results (
    resultIdStable INTEGER PRIMARY KEY AUTOINCREMENT,
    TaskId INTEGER,
    sampleCode TEXT NOT NULL,
    dateStart TEXT,
    timeStart TEXT,
    dateEnd TEXT,
    timeEnd TEXT,
    operator TEXT,
    PVTcell TEXT,
    temperature REAL,
    pressure REAL,
    pressureUnit TEXT,
    pressureAbs INTEGER,
    pressureMPaAbs REAL,
    LiquidVolume REAL,
    dateTimeSync TEXT
);

CREATE UNIQUE INDEX ux_Stable_results_sampleCode
    ON Stable_results (sampleCode);

CREATE INDEX IX_Stable_results_TaskId
    ON Stable_results (TaskId);
```

Номер и версию миграции нужно оформить по текущей последовательности проекта.

## 10. Что менять не нужно

- Ribbon XML;
- `Functions.bas`: отсутствие листа `Stable_sourceData` уже обрабатывается;
- `helpersValue.bas`: поля `dateStart`, `timeStart`, `dateEnd`, `timeEnd` уже
  форматируются;
- `task_table.bas`: `Stable` уже присутствует в разрешённых заданиях.

## 11. Проверка после сборки

1. Выполнить `Debug` → `Compile VBAProject`.
2. На листе Stable проверить команды ленты: добавить/удалить форму, загрузить,
   сохранить, валидировать и отменить задание.
3. Проверить создание чекбокса `pressureAbs` на новой форме.
4. Проверить загрузку задания типа `Stable` через обновление проекта.
5. Проверить `Save all`, `Load all` и `Validate all`.
6. Убедиться, что повторный `sampleCode` отклоняется VBA и уникальным индексом
   SQLite.
