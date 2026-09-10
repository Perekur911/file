# Финальные изменения перед релизом DLST и Stable

Надстройка в этой инструкции не изменяется. Все блоки нужно перенести вручную
в рабочую копию исходников, после чего выполнить `Debug -> Compile VBAProject`.

Если какой-либо блок уже добавлен, второй раз его вставлять не нужно.

## 1. `DLST.bas`: ширина формы

В начале модуля должно быть:

```vb
Private Const DLST_FORM_COLUMN_COUNT As Long = 18
Private Const DLST_FIRST_FORM_FIRST_COL As Long = 1

Private Const DLST_PAGE_ROW As Long = 1
Private Const DLST_PAGE_FIRST_COL As Long = 7

Private Const DLST_FIRST_STEP_COL As Long = 6
Private Const DLST_LAST_REGULAR_STEP_COL As Long = 15
Private Const DLST_FINAL_STEP_COL As Long = 16
```

## 2. `DLST.bas`: автономная нумерация анализов

Вставить следующую процедуру перед `Public Function Save_DLST`:

```vb
Private Sub EnsureDLSTAnalyseNumber( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long _
)

    Dim sampleRange As Range
    Dim analyseRange As Range

    Set sampleRange = DLST_FormRange( _
        ws, formIndex, "sampleCode" _
    )

    Set analyseRange = DLST_FormRange( _
        ws, formIndex, "analyseNumber" _
    )

    If sampleRange Is Nothing Then Exit Sub
    If analyseRange Is Nothing Then Exit Sub

    Dim sampleCode As String
    sampleCode = Trim$(CStr(sampleRange.Value))

    If Len(sampleCode) = 0 Then Exit Sub
    If Len(Trim$(CStr(analyseRange.Value))) > 0 Then Exit Sub

    Dim usedNumbers As Object
    Set usedNumbers = CreateObject("Scripting.Dictionary")

    Dim otherFormIndex As Long
    Dim lastFormIndex As Long
    Dim otherSampleRange As Range
    Dim otherAnalyseRange As Range
    Dim otherAnalyseValue As Variant

    lastFormIndex = DLST_LastFormIndexOnSheet(ws)

    For otherFormIndex = 1 To lastFormIndex

        If otherFormIndex <> formIndex Then

            Set otherSampleRange = DLST_FormRange( _
                ws, otherFormIndex, "sampleCode" _
            )

            Set otherAnalyseRange = DLST_FormRange( _
                ws, otherFormIndex, "analyseNumber" _
            )

            If Not otherSampleRange Is Nothing _
               And Not otherAnalyseRange Is Nothing Then

                If StrComp( _
                    Trim$(CStr(otherSampleRange.Value)), _
                    sampleCode, _
                    vbTextCompare _
                ) = 0 Then

                    otherAnalyseValue = otherAnalyseRange.Value

                    If IsNumeric(otherAnalyseValue) Then
                        If CDbl(otherAnalyseValue) > 0 _
                           And CDbl(otherAnalyseValue) = _
                               Fix(CDbl(otherAnalyseValue)) Then

                            usedNumbers( _
                                CStr(CLng(otherAnalyseValue)) _
                            ) = True

                        End If
                    End If

                End If

            End If

        End If

    Next otherFormIndex

    Dim nextNumber As Long
    nextNumber = 1

    Do While usedNumbers.Exists(CStr(nextNumber))
        nextNumber = nextNumber + 1
    Loop

    analyseRange.Value = nextNumber

End Sub
```

В `Save_DLST`, в первом проходе по формам, найти блок с
`checkNumberAnalys2` и полностью заменить его.

Удалить:

```vb
Dim analyseRange As Range
Dim sampleRange As Range

Set analyseRange = DLST_FormRange(ws, formIndex, "analyseNumber")
Set sampleRange = DLST_FormRange(ws, formIndex, "sampleCode")

If Not analyseRange Is Nothing And Not sampleRange Is Nothing Then
    If Len(Trim$(CStr(analyseRange.Value))) = 0 Then
        Call checkNumberAnalys2( _
            sampleRange.Value, _
            DLST_STUDY_CODE _
        )
    End If
End If
```

Вставить вместо него:

```vb
EnsureDLSTAnalyseNumber ws, formIndex
```

## 3. `DLST.bas`: проверка дубликатов `sampleCode + analyseNumber`

В `Save_DLST` после:

```vb
Dim newResultIDs As New Collection
```

вставить:

```vb
Dim resultKeys As Object
Set resultKeys = CreateObject("Scripting.Dictionary")
resultKeys.CompareMode = vbTextCompare
```

В первом проходе по формам найти:

```vb
If Not Validate_DLST_Form(formIndex) Then Exit Function

validForms.Add formIndex
```

и заменить на:

```vb
If Not Validate_DLST_Form(formIndex) Then Exit Function

Dim sampleRange As Range
Dim analyseRange As Range
Dim sampleCode As String
Dim analyseNumber As String
Dim resultKey As String

Set sampleRange = DLST_FormRange( _
    ws, formIndex, "sampleCode" _
)

Set analyseRange = DLST_FormRange( _
    ws, formIndex, "analyseNumber" _
)

sampleCode = Trim$(CStr(sampleRange.Value))
analyseNumber = NormalizeKeyValue(analyseRange.Value)

resultKey = sampleCode & Chr$(30) & analyseNumber

If resultKeys.Exists(resultKey) Then

    MsgBox _
        "DLST_" & formIndex & " не сохранена." & _
        vbCrLf & vbCrLf & _
        "Сочетание шифра пробы " & sampleCode & _
        " и номера анализа " & analyseNumber & _
        " уже указано на странице " & _
        resultKeys(resultKey) & ".", _
        vbExclamation

    Exit Function

End If

resultKeys.Add resultKey, formIndex
validForms.Add formIndex
```

## 4. `DLST.bas`: номер анализа должен быть положительным целым

В `Validate_DLST_Form` найти:

```vb
Case "analyseNumber"
```

и заменить весь этот `Case` на:

```vb
Case "analyseNumber"

    If Len(valueText) = 0 Then

        AddValidationError _
            errors, errorCount, _
            "Не указан номер анализа: " & fieldLabel, _
            fieldRange

    ElseIf Not IsNumeric(fieldRange.Value) Then

        AddValidationError _
            errors, errorCount, _
            "Номер анализа должен быть числом: " & fieldLabel, _
            fieldRange

    ElseIf CDbl(fieldRange.Value) <= 0 _
           Or CDbl(fieldRange.Value) <> _
              Fix(CDbl(fieldRange.Value)) Then

        AddValidationError _
            errors, errorCount, _
            "Номер анализа должен быть положительным " & _
            "целым числом: " & fieldLabel, _
            fieldRange

    End If
```

## 5. `DLST.bas`: заменить списки числовых полей

Полностью заменить `DLST_IsNumericResultField`:

```vb
Private Function DLST_IsNumericResultField( _
    ByVal fieldName As String _
) As Boolean

    Select Case LCase$(Trim$(fieldName))

        Case "densityap", _
             "ptransfer", "ttransfer", "vtransfer", _
             "pstart", "tstart", "vstart", _
             "density20", "deltamatbalance", _
             "samplemassin", "samplemassout"

            DLST_IsNumericResultField = True

    End Select

End Function
```

Полностью заменить `DLST_IsNumericSourceField`:

```vb
Private Function DLST_IsNumericSourceField( _
    ByVal fieldName As String _
) As Boolean

    Select Case LCase$(Trim$(fieldName))

        Case "steppcell", "steptcell", _
             "vcell", "vcellliq", "vafterout", _
             "transferpcell", "transfertcell", _
             "v1", "v2", _
             "m0", "m1", "m0trap", "m1trap", _
             "patm", "initpatm", "t", _
             "vhetot", "vhecyl", _
             "vgtot", "vgcyl", "m2", _
             "apdensity", "vtransfer", _
             "vgas", "vgasst", "mgas", "dgas", _
             "densliqcell", "vliq", "mliq", _
             "volumetriccoefficient", _
             "gascontent", "gf"

            DLST_IsNumericSourceField = True

    End Select

End Function
```

## 6. `Functions.bas`: алиас листов DLST

Это изменение нужно из-за того, что код исследования называется `DLST`, а
листы начинаются с `DLSepTest`.

Перед процедурой `PrepareSheet` вставить:

```vb
Private Function GetStudySheetPrefix( _
    ByVal studyCode As String _
) As String

    Select Case UCase$(Trim$(studyCode))

        Case "DLST"
            GetStudySheetPrefix = "DLSepTest"

        Case Else
            GetStudySheetPrefix = studyCode

    End Select

End Function
```

В `StartMacro` после:

```vb
Dim studyCode As String
```

добавить:

```vb
Dim sheetPrefix As String
```

Внутри цикла заменить три вызова `PrepareSheet` на:

```vb
studyCode = CStr(studies(i))
sheetPrefix = GetStudySheetPrefix(studyCode)

PrepareSheet wb, sheetPrefix, False, enableSheetCalculation
PrepareSheet wb, sheetPrefix & "_results", False, enableSheetCalculation
PrepareSheet wb, sheetPrefix & "_sourceData", False, enableSheetCalculation
```

В `EndMacro` также добавить `Dim sheetPrefix As String` и заменить три вызова:

```vb
studyCode = CStr(studies(i))
sheetPrefix = GetStudySheetPrefix(studyCode)

PrepareSheet wb, sheetPrefix, True, enableSheetCalculation
PrepareSheet wb, sheetPrefix & "_results", True, enableSheetCalculation
PrepareSheet wb, sheetPrefix & "_sourceData", True, enableSheetCalculation
```

После этого во всех wrappers продолжать передавать код `"DLST"`, а не
`"DLSepTest"`.

## 7. `modRibbonPVT.bas`

### `CurrentStudyName`

Заменить `Select Case sheetName` на:

```vb
Select Case sheetName

    Case "DLSEPTEST"
        CurrentStudyName = "DLST"

    Case "OP", "OPOH", "AP", "GC", "SSF", "GOR", _
         "BP", "REC", "EMV", "JOIN", "CCE", "CVD", _
         "STABLE"

        CurrentStudyName = sheetName

End Select
```

### `RunStudyWrapper`

Перед `Case Else` добавить:

```vb
        Case "SAVE_DLST":           Save_DLST_wrap
        Case "LOAD_DLST":           Load_DLST_wrap
        Case "VALIDATE_DLST":       Validate_DLST_wrap
        Case "ADD_FORM_DLST":       Add_DLST_Form_wrap
        Case "DELETE_FORM_DLST":    Delete_DLST_Form_wrap
        Case "CANCEL_TASK_DLST":    Cancel_DLST_Task_wrap

        Case "SAVE_STABLE":         Save_Stable_wrap
        Case "LOAD_STABLE":         Load_Stable_wrap
        Case "VALIDATE_STABLE":     Validate_Stable_wrap
        Case "ADD_FORM_STABLE":     Add_Stable_Form_wrap
        Case "DELETE_FORM_STABLE":  Delete_Stable_Form_wrap
        Case "CANCEL_TASK_STABLE":  Cancel_Stable_Task_wrap
```

Ribbon XML менять не нужно.

## 8. `clsAppEvents.cls`

Полностью заменить `GetStudyCodeFromSheet`:

```vb
Private Function GetStudyCodeFromSheet( _
    ByVal sheetName As String _
) As String

    Select Case UCase$(Trim$(sheetName))

        Case "DLSEPTEST"
            GetStudyCodeFromSheet = "DLST"

        Case "OP", "OPOH", "AP", "BP", "GC", "EMV", _
             "GOR", "SSF", "CCE", "CVD", "REC", "JOIN", _
             "STABLE"

            GetStudyCodeFromSheet = UCase$(Trim$(sheetName))

    End Select

End Function
```

## 9. `wrappers.bas`

Добавить процедуры рядом с wrappers CVD/OPOH:

```vb
Public Sub Add_DLST_Form_wrap()
    CallMacro "AddFormDLST", "DLST"
End Sub

Public Sub Delete_DLST_Form_wrap()
    CallMacro "DeleteLastFormDLST", "DLST"
End Sub

Public Sub Cancel_DLST_Task_wrap()
    CallMacroEx _
        "cancelCurrentTaskDLST", _
        REFRESH_TASKS_AFTER, _
        "DLST"
End Sub

Public Sub Save_DLST_wrap()
    CallMacroEx _
        "Save_DLST", _
        REFRESH_TASKS_AFTER, _
        "DLST"
End Sub

Public Sub Load_DLST_wrap()
    CallMacro "Load_DLST", "DLST"
End Sub

Public Sub Validate_DLST_wrap()
    CallMacroEx _
        "Validate_DLST", _
        REFRESH_TASKS_AFTER, _
        "DLST"
End Sub

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

В списки исследований процедур `Load_All_wrap`, `Save_All_wrap` и
`Validate_All_wrap` добавить в конец:

```vb
, "DLST", "STABLE"
```

То же добавить в списки `Refresh_Project_wrap` и
`silentRefresh_Project_wrap`.

В `Load_All_results_to_forms_Auto` после блока CVD вставить:

```vb
    ' DLST
    feedback = Load_DLST_results_to_forms( _
        mode:="sequential", _
        silence:=True _
    )

    AddLoadFeedback _
        reportText, _
        hasWarnings, _
        hasErrors, _
        "DLST", _
        feedback

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

## 10. `StudySync.bas`

В `RunStudySync`, в строку исследований по умолчанию, добавить:

```vb
studies = _
    "OP AP GC SSF GOR BP REC EMV OPOH CVD CCE JOIN DLST STABLE"
```

Полностью обновить `DefaultStudies`:

```vb
Private Function DefaultStudies() As Variant

    DefaultStudies = Array( _
        "OP", "OPOH", "GC", "BP", "AP", "GOR", "SSF", _
        "REC", "EMV", "JOIN", "CCE", "CVD", "DLST", _
        "STABLE" _
    )

End Function
```

В `LoadStudyForms`, перед `Case Else`, добавить:

```vb
        Case "DLST"
            LoadStudyForms = _
                Load_DLST_results_to_forms(formFillMode)

        Case "STABLE"
            LoadStudyForms = _
                Load_Stable_results_to_forms(formFillMode)
```

В `IsKnownStudyCode` добавить `"DLST", "STABLE"` в общий `Case`.

В `Save_All`, перед итоговой строкой `Save_All = totalResult`, добавить два
блока:

```vb
    currentResult = Save_DLST()

    If currentResult = MACRO_FAILED Then
        Save_All = MACRO_FAILED
        Exit Function
    End If

    If currentResult = MACRO_CHANGED Then
        totalResult = MACRO_CHANGED
    End If

    currentResult = Save_Stable()

    If currentResult = MACRO_FAILED Then
        Save_All = MACRO_FAILED
        Exit Function
    End If

    If currentResult = MACRO_CHANGED Then
        totalResult = MACRO_CHANGED
    End If
```

В `Validate_All`, перед итоговой строкой `Validate_All = totalResult`, добавить:

```vb
    currentResult = Validate_DLST()

    If currentResult = MACRO_FAILED Then
        Validate_All = MACRO_FAILED
        Exit Function
    End If

    If currentResult = MACRO_CHANGED Then
        totalResult = MACRO_CHANGED
    End If

    currentResult = Validate_Stable()

    If currentResult = MACRO_FAILED Then
        Validate_All = MACRO_FAILED
        Exit Function
    End If

    If currentResult = MACRO_CHANGED Then
        totalResult = MACRO_CHANGED
    End If
```

## 11. `modMetadata.bas`

Добавить константы к трём соответствующим группам:

```vb
Public Const PROP_DLST_DATA_HASH As String = "DLSTDataHash"
Public Const PROP_STABLE_DATA_HASH As String = "StableDataHash"

Public Const PROP_DLST_DATA_STATE As String = "DLSTDataState"
Public Const PROP_STABLE_DATA_STATE As String = "StableDataState"

Public Const PROP_DLST_LAST_WORKFLOW_STAGE As String = _
    "DLSTLastWorkflowStage"

Public Const PROP_STABLE_LAST_WORKFLOW_STAGE As String = _
    "StableLastWorkflowStage"
```

В `RefreshAllStudySyncStatuses` добавить:

```vb
UpdateStudySyncStatus _
    wb, "DLST", PROP_DLST_DATA_HASH, PROP_DLST_DATA_STATE

UpdateStudySyncStatus _
    wb, "STABLE", PROP_STABLE_DATA_HASH, PROP_STABLE_DATA_STATE
```

В `RefreshChangedStudySyncStatuses` добавить:

```vb
RefreshStudyIfNeeded _
    wb, "DLST", PROP_DLST_DATA_HASH, PROP_DLST_DATA_STATE

RefreshStudyIfNeeded _
    wb, "STABLE", PROP_STABLE_DATA_HASH, PROP_STABLE_DATA_STATE
```

В `GetCurrentStudyHash`, перед `Case Else`, добавить:

```vb
        Case "DLST"
            GetCurrentStudyHash = CalculateCurrentDLSTHash()

        Case "STABLE"
            GetCurrentStudyHash = CalculateCurrentStableHash()
```

В `GetUnsyncedStudiesText` конец двух массивов должен выглядеть так:

```vb
studyNames = Array( _
    "OP", "OPOH", "AP", "BP", "GC", "EMV", _
    "GOR", "SSF", "CCE", "CVD", "REC", "JOIN", _
    "DLST", "STABLE" _
)

stateProperties = Array( _
    PROP_OP_DATA_STATE, PROP_OPOH_DATA_STATE, _
    PROP_AP_DATA_STATE, PROP_BP_DATA_STATE, _
    PROP_GC_DATA_STATE, PROP_EMV_DATA_STATE, _
    PROP_GOR_DATA_STATE, PROP_SSF_DATA_STATE, _
    PROP_CCE_DATA_STATE, PROP_CVD_DATA_STATE, _
    PROP_REC_DATA_STATE, PROP_JOIN_DATA_STATE, _
    PROP_DLST_DATA_STATE, PROP_STABLE_DATA_STATE _
)
```

В `SetStudyBaselineAfterLoad`, перед `Case Else`, добавить:

```vb
        Case "DLST"
            hashProperty = PROP_DLST_DATA_HASH
            stateProperty = PROP_DLST_DATA_STATE

        Case "STABLE"
            hashProperty = PROP_STABLE_DATA_HASH
            stateProperty = PROP_STABLE_DATA_STATE
```

В `SetAllStudyBaselinesAfterLoad` добавить:

```vb
SetStudyBaselineAfterLoad wb, "DLST"
SetStudyBaselineAfterLoad wb, "STABLE"
```

## 12. `PowerQuery.bas`

### Объявления в `Refresh_Project`

К списку листов добавить:

```vb
DLST_list As Worksheet, Stable_list As Worksheet
```

К назначениям листов добавить:

```vb
Set DLST_list = ActiveWorkbook.Worksheets("DLSepTest")
Set Stable_list = ActiveWorkbook.Worksheets("Stable")
```

К счётчикам и номерам свободных форм добавить:

```vb
Dim countDLST As Integer, countStable As Integer
Dim emptyDLSTPage As Integer, emptyStablePage As Integer
```

### В цепочку обработки `taskType`

Перед веткой OPOH или после CVD вставить:

```vb
        ElseIf UCase$(Trim$(taskType)) = "DLST" Then

            DLST_list.Visible = xlSheetVisible
            countDLST = countDLST + 1

            Dim existingDLSTForm As Long
            existingDLSTForm = FindDLSTFormByTaskId(cellX.Value)

            If existingDLSTForm = 0 Then

                emptyDLSTPage = FindEmptyFormDLST()

                If emptyDLSTPage = 0 Then
                    Err.Raise _
                        vbObjectError + 4301, _
                        "Refresh_Project", _
                        "Для задания DLST №" & cellX.Value & _
                        " нет свободной формы DLSepTest."
                End If

                With GetDLSTFormRange( _
                    emptyDLSTPage, "TaskId" _
                ).MergeArea

                    .Value = cellX.Value
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark2
                    .Interior.TintAndShade = _
                        -9.99481185338908E-02
                    .Interior.PatternTintAndShade = 0
                    .Locked = True

                End With

                With GetDLSTFormRange( _
                    emptyDLSTPage, "sampleCode" _
                ).MergeArea

                    .Value = _
                        cellX.Offset(0, offsetToSampleCode).Value
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark2
                    .Interior.TintAndShade = _
                        -9.99481185338908E-02
                    .Interior.PatternTintAndShade = 0
                    .Locked = True

                End With

            End If

        ElseIf UCase$(Trim$(taskType)) = "STABLE" Then

            Stable_list.Visible = xlSheetVisible
            countStable = countStable + 1

            Dim existingStableForm As Long
            existingStableForm = FindStableFormByTaskId(cellX.Value)

            If existingStableForm = 0 Then

                emptyStablePage = FindEmptyFormStable()

                With GetStableFormRange( _
                    emptyStablePage, "TaskId" _
                ).MergeArea

                    .Value = cellX.Value
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark2
                    .Interior.TintAndShade = _
                        -9.99481185338908E-02
                    .Interior.PatternTintAndShade = 0
                    .Locked = True

                End With

                With GetStableFormRange( _
                    emptyStablePage, "sampleCode" _
                ).MergeArea

                    .Value = _
                        cellX.Offset(0, offsetToSampleCode).Value
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark2
                    .Interior.TintAndShade = _
                        -9.99481185338908E-02
                    .Interior.PatternTintAndShade = 0
                    .Locked = True

                End With

            End If
```

Для DLST нельзя обращаться к `DLST_2_TaskId` напрямую: таких имён нет.
Использовать только `GetDLSTFormRange`.

## 13. Python-схема

Для DLST должно быть:

```python
_col(
    "density20",
    "density20",
    typ="REAL",
    value_type="real",
),
```

Не должно остаться Excel-имени `"20density"`.

Для Stable должно присутствовать:

```python
_col(
    "analyseNumber",
    "analyseNumber",
    typ="INTEGER",
    value_type="integer",
    required=True,
),
```

В `STUDIES`:

```python
"STABLE": StudySchema("STABLE", STABLE_RESULTS, None),

"DLST": StudySchema(
    code="DLST",
    result_table=DLST_RESULTS,
    source_table=DLST_SOURCE_DATA,
    progress_step_field="step",
    progress_total_steps=10,
    completion_step=11,
),
```

Проверить, что все варианты пустого необязательного `TEXT` преобразуются в
`None`, а не в `""`. Это важно для уникальных `sampleCodeLPh` и
`sampleCodeGPh`.

## 14. Миграции SQLite

Если миграционный runner не регистрирует миграции самостоятельно, перед
`COMMIT` в `002_add_dlst_tables.sql` вставить:

```sql
INSERT INTO schema_migrations (
    migration_id,
    schema_version,
    migration_name
)
VALUES (
    2,
    '1.1.0',
    '002_add_dlst_tables'
);
```

Перед `COMMIT` в `003_add_stable_table.sql` вставить:

```sql
INSERT INTO schema_migrations (
    migration_id,
    schema_version,
    migration_name
)
VALUES (
    3,
    '1.2.0',
    '003_add_stable_table'
);
```

Если в проекте выбраны другие номера версий, заменить `1.1.0` и `1.2.0`.
Последняя версия должна совпасть с `database_schema` в `project_version.json`.

Рекомендуемая дополнительная защита от двух одинаковых step одного результата:

```sql
CREATE UNIQUE INDEX UX_DLST_sourceData_resultIdDLSP_step
    ON DLST_sourceData (resultIdDLSP, step)
    WHERE step IS NOT NULL;
```

Если добавляется этот уникальный индекс, существующий обычный индекс
`IX_DLST_sourceData_resultIdDLSP_step` можно не создавать.

## 15. `Stable.bas`: показывать ошибку сохранения

В конце `Save_Stable` заменить обработчик:

```vb
ErrHandler:
    Debug.Print "Save_Stable: " & Err.Number & ": " & Err.Description
    Save_Stable = MACRO_FAILED
```

на:

```vb
ErrHandler:

    Debug.Print _
        "Save_Stable: " & Err.Number & ": " & Err.Description

    MsgBox _
        "Ошибка Save_Stable: " & _
        Err.Number & ": " & Err.Description, _
        vbExclamation

    Save_Stable = MACRO_FAILED
```

Существующий `End Function`, находящийся сразу после обработчика, оставить на
месте. Итоговая процедура должна содержать только один `End Function`.

## 16. `task_table.bas`

Этот пункт нужен, только если именно `task_table.bas` формирует таблицу `Task`.

В `AllowedTasks` добавить официальное значение задания DLST:

```vb
AllowedTasks = Array( _
    "OP", "BP", "GC", "AP", "GOR", "Stable", "MIX", _
    "DLST" _
)
```

Если DLST переносится также в сводную таблицу `Table`, добавить `"DLST"` в
`AllowedTasksTarget`. В самой таблице при этом должны существовать требуемые
столбцы с суффиксом `_DLST`. Иначе строка:

```vb
Set tmp_map = target_map_col(cur_task)
```

получит отсутствующий ключ и завершится ошибкой.

## 17. Формы и таблицы

Итоговые имена должны быть точными:

| Назначение | DLST | Stable |
| --- | --- | --- |
| Код исследования | `DLST` | `STABLE` |
| Лист формы | `DLSepTest` | `Stable` |
| Лист результатов | `DLSepTest_results` | `Stable_results` |
| Лист sourceData | `DLSepTest_sourceData` | отсутствует |
| Таблица результатов | `DLST_results` | `Stable_results` |
| Таблица sourceData | `DLST_sourceData` | отсутствует |
| Базовый префикс имён | `DLST_1_` | `Stable_1_` |

В имени `volumetricCoefficient` должна использоваться латинская буква `C`.

## 18. Версии и финальная проверка

Обновить:

- общую версию проекта;
- `PVTFormVersion` в форме;
- `PVTAddinVersion` в надстройке;
- версию Python-компонента;
- `database_schema`;
- ожидаемое количество экспортируемых VBA-компонентов;
- список обязательных миграций.

После сборки:

1. Выполнить `Debug -> Compile VBAProject`.
2. Проверить DLST на формах 1, 2 и последней форме.
3. Проверить одинаковый `sampleCode` с анализами 1 и 2.
4. Проверить запрет одинаковой пары `sampleCode + analyseNumber`.
5. Проверить DLST `Step1`, `Step10` и `Last Step`.
6. Сохранить, загрузить, изменить и повторно сохранить существующие записи.
7. Очистить ранее сохранённый DLST-step и убедиться, что он удалился из БД.
8. Проверить `Save all`, `Load all`, `Validate all`.
9. Проверить обновление проекта и отмену задания.
10. Проверить предупреждение о несинхронизированных DLST/Stable при закрытии.
11. На копии БД выполнить обе миграции и затем:

```sql
PRAGMA integrity_check;
PRAGMA foreign_key_check;
```

Порядок выкладки: резервная копия БД, миграции, Python, форма, надстройка.
