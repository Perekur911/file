# CVD: переход на имена только первой формы

Документ составлен по последней присланной версии `CVD.bas`. Сам модуль и надстройка не изменялись.

После переделки в книге должны существовать только базовые имена:

- родительское поле: `CVD_1_<поле>`;
- поле дочерней таблицы: `CVD_1_st1_<поле>` ... `CVD_1_st10_<поле>` и `CVD_1_last_<поле>`.

Для форм 2+ VBA получает диапазон смещением базового диапазона `CVD_1_*` вправо на:

```text
(номер формы - 1) * CVD_FORM_COLUMN_COUNT
```

## 1. Константы в начале `CVD.bas`

Существующие константы геометрии заменить следующим блоком. Значение `16` взято из `CVD_dev.CreateShiftedNamedRanges`, где сейчас указан `COLUMN_SHIFT = 16`. Перед внедрением его всё равно нужно сверить с фактической формой.

```vb
Private Const CVD_STUDY_CODE As String = "CVD"

Private Const CVD_FORM_SHEET As String = "CVD"
Private Const CVD_RESULTS_SHEET As String = "CVD_results"
Private Const CVD_SOURCE_SHEET As String = "CVD_sourceData"

Private Const CVD_RESULTS_TABLE As String = "CVD_results"
Private Const CVD_SOURCE_TABLE As String = "CVD_sourceData"

' Проверить по фактическому листу перед релизом.
Private Const CVD_FORM_COLUMN_COUNT As Long = 16
Private Const CVD_FIRST_FORM_FIRST_COL As Long = 1

' Номер страницы первой формы: G1.
Private Const CVD_PAGE_ROW As Long = 1
Private Const CVD_PAGE_FIRST_COL As Long = 7

' st1 = F, st10 = O, last = P.
' Если фактическое расположение другое — поменять только эти константы.
Private Const CVD_FIRST_STEP_COL As Long = 6
Private Const CVD_LAST_REGULAR_STEP_COL As Long = 15
Private Const CVD_FINAL_STEP_COL As Long = 16

Private Const CVD_MIN_FORM_COUNT As Long = 1
Private Const CVD_MAX_FORMS As Long = 999
```

Константы полей БД `COL_CVD_*` оставить как есть.

`CVD_MIN_FORM_COUNT = 1` означает, что удалить базовую первую форму нельзя, а вторую и последующие — можно. Если в шаблоне обязательно должны постоянно оставаться две формы, поставить `2`.

## 2. Добавить общий доступ к диапазонам и страницам

Этот блок вставить после констант и до `cancelCurrentTaskCVD`.

```vb
Public Function GetCVDFormRange( _
    ByVal formIndex As Long, _
    ByVal fieldKey As String _
) As Range

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(CVD_FORM_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    Set GetCVDFormRange = CVD_FormRange( _
        ws, _
        formIndex, _
        fieldKey _
    )

End Function


Private Function CVD_FormRange( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long, _
    ByVal fieldKey As String _
) As Range

    If ws Is Nothing Then Exit Function
    If formIndex <= 0 Then Exit Function
    If CVD_FORM_COLUMN_COUNT <= 0 Then Exit Function
    If Len(Trim$(fieldKey)) = 0 Then Exit Function

    Dim baseRange As Range
    Dim targetRange As Range
    Dim columnShift As Long

    On Error Resume Next
    Set baseRange = ws.Parent.Names( _
        "CVD_1_" & fieldKey _
    ).RefersToRange
    On Error GoTo 0

    If baseRange Is Nothing Then Exit Function
    If Not baseRange.Worksheet Is ws Then Exit Function

    columnShift = _
        (formIndex - 1) * CVD_FORM_COLUMN_COUNT

    If baseRange.Column + columnShift + _
       baseRange.Columns.count - 1 > ws.Columns.count Then
        Exit Function
    End If

    On Error Resume Next
    Set targetRange = baseRange.Offset(0, columnShift)
    On Error GoTo 0

    Set CVD_FormRange = targetRange

End Function


Private Function CVD_PageCell( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long _
) As Range

    If ws Is Nothing Then Exit Function
    If formIndex <= 0 Then Exit Function
    If CVD_FORM_COLUMN_COUNT <= 0 Then Exit Function

    Dim targetColumn As Long

    targetColumn = CVD_PAGE_FIRST_COL + _
                   (formIndex - 1) * CVD_FORM_COLUMN_COUNT

    If targetColumn > ws.Columns.count Then Exit Function

    Set CVD_PageCell = _
        ws.Cells(CVD_PAGE_ROW, targetColumn)

End Function


Public Function CVD_GetLastFormIndex() As Long

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(CVD_FORM_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    CVD_GetLastFormIndex = _
        CVD_LastFormIndexOnSheet(ws)

End Function


Private Function CVD_LastFormIndexOnSheet( _
    ByVal ws As Worksheet _
) As Long

    Dim formIndex As Long
    Dim pageCell As Range
    Dim pageValue As Variant

    For formIndex = 1 To CVD_MAX_FORMS

        Set pageCell = Nothing
        Set pageCell = CVD_PageCell(ws, formIndex)

        If pageCell Is Nothing Then Exit For

        pageValue = pageCell.Value2

        If IsError(pageValue) Then Exit For
        If Len(Trim$(CStr(pageValue))) = 0 Then Exit For

        CVD_LastFormIndexOnSheet = formIndex

    Next formIndex

End Function


Public Function CVD_FormExists( _
    ByVal formIndex As Long _
) As Boolean

    If formIndex <= 0 Then Exit Function

    CVD_FormExists = _
        (formIndex <= CVD_GetLastFormIndex())

End Function


Private Function CVD_RangeLabel( _
    ByVal formIndex As Long, _
    ByVal fieldKey As String _
) As String

    CVD_RangeLabel = _
        "CVD_" & formIndex & "_" & fieldKey

End Function
```

`CVD_RangeLabel` формирует только понятное имя для сообщений об ошибках. Обращаться через это имя к `Worksheet.Range(...)` нельзя: физического имени формы 2+ больше нет.

## 3. Полностью заменить определение текущей формы и step

Старый `CVD_CurrentFormIndex` вычисляет длину по `CVD_2_sampleCode`, а старый `CVD_CurrentStepKey` читает имя активной ячейки. После удаления имён форм 2+ оба метода перестанут работать.

Обе функции заменить целиком:

```vb
Private Function CVD_CurrentFormIndex() As Long

    CVD_CurrentFormIndex = 0

    If ActiveWorkbook Is Nothing Then Exit Function
    If ActiveSheet Is Nothing Then Exit Function
    If CVD_FORM_COLUMN_COUNT <= 0 Then Exit Function

    If StrComp( _
        ActiveSheet.name, _
        CVD_FORM_SHEET, _
        vbTextCompare _
    ) <> 0 Then Exit Function

    Dim formIndex As Long

    formIndex = _
        (ActiveCell.Column - CVD_FIRST_FORM_FIRST_COL) \ _
        CVD_FORM_COLUMN_COUNT + 1

    If formIndex < 1 Then Exit Function
    If formIndex > CVD_GetLastFormIndex() Then Exit Function

    CVD_CurrentFormIndex = formIndex

End Function


Private Function CVD_CurrentStepKey() As String

    CVD_CurrentStepKey = vbNullString

    Dim formIndex As Long
    formIndex = CVD_CurrentFormIndex()

    If formIndex <= 0 Then Exit Function

    Dim localColumn As Long

    localColumn = ActiveCell.Column - _
                  (formIndex - 1) * CVD_FORM_COLUMN_COUNT

    If localColumn >= CVD_FIRST_STEP_COL _
       And localColumn <= CVD_LAST_REGULAR_STEP_COL Then

        CVD_CurrentStepKey = _
            "st" & CStr( _
                localColumn - CVD_FIRST_STEP_COL + 1 _
            )

    ElseIf localColumn = CVD_FINAL_STEP_COL Then

        CVD_CurrentStepKey = "last"

    End If

End Function
```

## 4. Заменить заглушки добавления и удаления формы

Старые `AddFormCVD` и `DeleteLastListCVD` удалить целиком. На их место вставить следующий блок.

```vb
Public Sub AddFormCVD()

    On Error GoTo CleanFail

    If CVD_FORM_COLUMN_COUNT <= 0 Then
        Err.Raise _
            vbObjectError + 4200, _
            "AddFormCVD", _
            "Не задана длина формы CVD."
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        Err.Raise _
            vbObjectError + 4201, _
            "AddFormCVD", _
            "Нет активной книги с формой CVD."
    End If

    Dim ws As Worksheet
    Set ws = wb.Worksheets(CVD_FORM_SHEET)

    Dim oldFormIndex As Long
    Dim newFormIndex As Long

    oldFormIndex = CVD_LastFormIndexOnSheet(ws)

    If oldFormIndex < CVD_MIN_FORM_COUNT Then
        Err.Raise _
            vbObjectError + 4202, _
            "AddFormCVD", _
            "Не найдена первая форма CVD."
    End If

    If oldFormIndex >= CVD_MAX_FORMS Then
        Err.Raise _
            vbObjectError + 4203, _
            "AddFormCVD", _
            "Достигнуто максимальное число форм CVD."
    End If

    newFormIndex = oldFormIndex + 1

    Dim sourceFirstColumn As Long
    Dim sourceLastColumn As Long
    Dim targetFirstColumn As Long
    Dim targetLastColumn As Long

    sourceFirstColumn = _
        CVD_FIRST_FORM_FIRST_COL + _
        (oldFormIndex - 1) * CVD_FORM_COLUMN_COUNT

    sourceLastColumn = _
        sourceFirstColumn + CVD_FORM_COLUMN_COUNT - 1

    targetFirstColumn = _
        sourceFirstColumn + CVD_FORM_COLUMN_COUNT

    targetLastColumn = _
        targetFirstColumn + CVD_FORM_COLUMN_COUNT - 1

    If targetLastColumn > ws.Columns.count Then
        Err.Raise _
            vbObjectError + 4204, _
            "AddFormCVD", _
            "Для новой формы CVD недостаточно столбцов листа."
    End If

    Dim sourceColumns As Range

    Set sourceColumns = ws.Range( _
        ws.Columns(sourceFirstColumn), _
        ws.Columns(sourceLastColumn) _
    )

    Dim oldCopyObjectsWithCells As Boolean
    Dim copySettingChanged As Boolean

    oldCopyObjectsWithCells = _
        Application.CopyObjectsWithCells

    Application.CopyObjectsWithCells = False
    copySettingChanged = True

    sourceColumns.Copy

    ws.Columns(targetFirstColumn).PasteSpecial _
        Paste:=xlPasteAll

    ws.Columns(targetFirstColumn).PasteSpecial _
        Paste:=xlPasteColumnWidths

    Application.CutCopyMode = False
    Application.CopyObjectsWithCells = _
        oldCopyObjectsWithCells

    copySettingChanged = False

    ' Новые имена CVD_N_* не создаются.
    CVD_ClearCopiedFormFields _
        wb, _
        ws, _
        newFormIndex

    Dim pageCell As Range
    Set pageCell = CVD_PageCell(ws, newFormIndex)

    If pageCell Is Nothing Then
        Err.Raise _
            vbObjectError + 4205, _
            "AddFormCVD", _
            "Не удалось определить ячейку страницы CVD."
    End If

    pageCell.Value2 = newFormIndex

    Exit Sub

CleanFail:

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    Application.CutCopyMode = False

    If copySettingChanged Then
        Application.CopyObjectsWithCells = _
            oldCopyObjectsWithCells
    End If

    On Error GoTo 0

    Debug.Print _
        "AddFormCVD: " & errNumber & ": " & errDescription

    Err.Raise errNumber, errSource, errDescription

End Sub


Public Sub DeleteLastListCVD()

    On Error GoTo CleanFail

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = wb.Worksheets(CVD_FORM_SHEET)

    Dim lastFormIndex As Long
    lastFormIndex = CVD_LastFormIndexOnSheet(ws)

    ' Первую форму удалять нельзя: от неё считаются смещения.
    If lastFormIndex <= CVD_MIN_FORM_COUNT Then
        Debug.Print _
            "DeleteLastListCVD: первая форма не удаляется."
        Exit Sub
    End If

    Dim firstColumn As Long
    Dim lastColumn As Long

    firstColumn = _
        CVD_FIRST_FORM_FIRST_COL + _
        (lastFormIndex - 1) * CVD_FORM_COLUMN_COUNT

    lastColumn = _
        firstColumn + CVD_FORM_COLUMN_COUNT - 1

    ws.Range( _
        ws.Columns(firstColumn), _
        ws.Columns(lastColumn) _
    ).EntireColumn.Delete

    Exit Sub

CleanFail:

    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    Debug.Print _
        "DeleteLastListCVD: " & _
        errNumber & ": " & errDescription

    Err.Raise errNumber, errSource, errDescription

End Sub


Private Sub CVD_ClearCopiedFormFields( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long _
)

    Dim resultTable As ListObject
    Dim sourceTable As ListObject

    Set resultTable = _
        wb.Worksheets(CVD_RESULTS_SHEET) _
          .ListObjects(CVD_RESULTS_TABLE)

    Set sourceTable = _
        wb.Worksheets(CVD_SOURCE_SHEET) _
          .ListObjects(CVD_SOURCE_TABLE)

    Dim tableColumn As ListColumn
    Dim fieldName As String
    Dim targetRange As Range

    For Each tableColumn In resultTable.ListColumns

        fieldName = Trim$(CStr(tableColumn.name))

        If StrComp( _
            fieldName, _
            "dateTimeSync", _
            vbTextCompare _
        ) <> 0 Then

            Set targetRange = Nothing
            Set targetRange = CVD_FormRange( _
                ws, _
                formIndex, _
                fieldName _
            )

            CVD_ClearNonFormulaCells targetRange

        End If

    Next tableColumn

    Dim steps As Variant
    Dim stepIndex As Long
    Dim stepKey As String

    steps = CVD_StepKeys()

    For stepIndex = LBound(steps) To UBound(steps)

        stepKey = CStr(steps(stepIndex))

        For Each tableColumn In sourceTable.ListColumns

            fieldName = Trim$(CStr(tableColumn.name))

            Select Case LCase$(fieldName)

                Case LCase$(COL_CVD_RESULT_ID_DB), _
                     LCase$(COL_CVD_STEP_DB)

                    ' Эти поля вычисляются при сохранении.

                Case Else

                    Set targetRange = Nothing
                    Set targetRange = CVD_FormRange( _
                        ws, _
                        formIndex, _
                        stepKey & "_" & fieldName _
                    )

                    CVD_ClearNonFormulaCells targetRange

            End Select

        Next tableColumn

    Next stepIndex

End Sub


Private Sub CVD_ClearNonFormulaCells( _
    ByVal targetRange As Range _
)

    If targetRange Is Nothing Then Exit Sub

    If targetRange.MergeCells Then

        Dim mergedRange As Range
        Set mergedRange = targetRange.MergeArea

        If Not RangeHasFormula(mergedRange) Then
            mergedRange.ClearContents
        End If

        Exit Sub

    End If

    Dim cell As Range

    For Each cell In targetRange.Cells
        If Not cell.HasFormula Then
            cell.ClearContents
        End If
    Next cell

End Sub
```

Метод очищает только поля, соответствующие колонкам `CVD_results` и `CVD_sourceData`. Подписи, рамки и постоянные элементы формы не очищаются. Формулы сохраняются.

`Application.CopyObjectsWithCells = False` означает, что рисунки, кнопки и другие объекты листа вместе с формой не копируются. Для текущей схемы это соответствует обычному копированию без создания дополнительных элементов. Если внутри каждой формы CVD всё-таки есть объект, который должен повторяться, для него потребуется отдельное создание по аналогии с checkbox других форм.

## 5. Заменить `cancelCurrentTaskCVD`

Вместо прямого обращения:

```vb
Set taskCell = sheetCVD.Range("CVD_" & formIndex & "_TaskId")
```

поставить:

```vb
Set taskCell = CVD_FormRange( _
    sheetCVD, _
    formIndex, _
    "TaskId" _
)

If taskCell Is Nothing Then
    cancelCurrentTaskCVD = MACRO_NO_CHANGES
    Exit Function
End If
```

Остальную процедуру менять не требуется.

## 6. Полностью заменить `CreateSampleCodeCVD3`

```vb
Public Function CreateSampleCodeCVD3( _
    Optional ByVal formIndex As Long = 0, _
    Optional ByVal stepKey As String = "" _
) As Boolean

    CreateSampleCodeCVD3 = False

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    Set ws = wb.Worksheets(CVD_FORM_SHEET)

    If formIndex <= 0 Then Exit Function

    If Not CVD_FormExists(formIndex) Then
        MsgBox _
            "Форма CVD_" & formIndex & " не найдена.", _
            vbExclamation
        Exit Function
    End If

    stepKey = LCase$(Trim$(stepKey))

    If Len(stepKey) = 0 Then
        MsgBox _
            "Не указан step для создания шифров CVD.", _
            vbExclamation
        Exit Function
    End If

    Dim sampleRange As Range
    Dim lphRange As Range
    Dim gphRange As Range

    Set sampleRange = CVD_FormRange( _
        ws, formIndex, "sampleCode" _
    )

    Set lphRange = CVD_FormRange( _
        ws, _
        formIndex, _
        stepKey & "_sampleCodeLPh" _
    )

    Set gphRange = CVD_FormRange( _
        ws, _
        formIndex, _
        stepKey & "_sampleCodeGPh" _
    )

    If sampleRange Is Nothing Then
        MsgBox _
            "Не найден базовый диапазон CVD_1_sampleCode.", _
            vbExclamation
        Exit Function
    End If

    If lphRange Is Nothing Then
        MsgBox _
            "Не найден базовый диапазон CVD_1_" & _
            stepKey & "_sampleCodeLPh.", _
            vbExclamation
        Exit Function
    End If

    If gphRange Is Nothing Then
        MsgBox _
            "Не найден базовый диапазон CVD_1_" & _
            stepKey & "_sampleCodeGPh.", _
            vbExclamation
        Exit Function
    End If

    Dim baseSampleCode As String
    baseSampleCode = Trim$(CStr(sampleRange.Value))

    If Len(baseSampleCode) = 0 Then
        MsgBox _
            "Не заполнен шифр исследуемой пробы.", _
            vbExclamation
        Exit Function
    End If

    Dim sampleCodeLPh As New Collection
    Dim sampleCodeGPh As New Collection

    Dim existingLPh As String
    Dim existingGPh As String

    existingLPh = Trim$(CStr(lphRange.Value))
    existingGPh = Trim$(CStr(gphRange.Value))

    If Len(existingLPh) > 0 Then
        sampleCodeLPh.Add existingLPh
    End If

    If Len(existingGPh) > 0 Then
        sampleCodeGPh.Add existingGPh
    End If

    If Len(existingLPh) = 0 _
       And Len(existingGPh) = 0 Then

        Dim stepPostfix As String

        If stepKey = "last" Then
            stepPostfix = "laststep"
        Else
            stepPostfix = stepKey
        End If

        sampleCodeLPh.Add _
            baseSampleCode & _
            "-CVD-" & stepPostfix & "-Rinse"

        sampleCodeGPh.Add _
            baseSampleCode & _
            "-CVD-" & stepPostfix & "-G"

    End If

    Dim frm As New frmCheckCodes

    frm.InitWithCollections _
        sampleCodeLPh, _
        sampleCodeGPh, _
        CVD_STUDY_CODE, _
        sampleRange, _
        formIndex

    frm.show

    If Not frm.result Then
        MsgBox _
            "Ввод шифров для CVD_" & formIndex & _
            "_" & stepKey & " отменен.", _
            vbInformation
        Exit Function
    End If

    lphRange.NumberFormat = "@"
    gphRange.NumberFormat = "@"

    If sampleCodeLPh.count > 0 Then
        lphRange.Value = sampleCodeLPh(1)
    Else
        lphRange.Value = vbNullString
    End If

    If sampleCodeGPh.count > 0 Then
        gphRange.Value = sampleCodeGPh(1)
    Else
        gphRange.Value = vbNullString
    End If

    If sampleCodeLPh.count = 0 _
       And sampleCodeGPh.count = 0 Then

        MsgBox _
            "Для CVD_" & formIndex & "_" & stepKey & _
            " не указан ни один шифр полученной пробы.", _
            vbExclamation

        Exit Function
    End If

    CreateSampleCodeCVD3 = True

End Function
```

## 7. Добавить собственную нумерацию анализов CVD

`checkNumberAnalys2` использовать больше нельзя: он ищет `CVD_2_*`, читает `c.Name.Name` и обращается к именам каждой формы.

Перед `Save_CVD` добавить:

```vb
Private Sub EnsureCVDAnalyseNumber( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long _
)

    Dim sampleRange As Range
    Dim analyseRange As Range

    Set sampleRange = CVD_FormRange( _
        ws, formIndex, "sampleCode" _
    )

    Set analyseRange = CVD_FormRange( _
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

    lastFormIndex = CVD_LastFormIndexOnSheet(ws)

    For otherFormIndex = 1 To lastFormIndex

        If otherFormIndex <> formIndex Then

            Set otherSampleRange = CVD_FormRange( _
                ws, otherFormIndex, "sampleCode" _
            )

            Set otherAnalyseRange = CVD_FormRange( _
                ws, otherFormIndex, "analyseNumber" _
            )

            If Not otherSampleRange Is Nothing _
               And Not otherAnalyseRange Is Nothing Then

                If StrComp( _
                    Trim$(CStr(otherSampleRange.Value)), _
                    sampleCode, _
                    vbTextCompare _
                ) = 0 Then

                    otherAnalyseValue = _
                        otherAnalyseRange.Value

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

## 8. Точечные изменения в `Save_CVD`

### 8.1. Удалить повторное объявление

В текущем файле два раза подряд объявлен `oldHashResult`. Оставить одну строку:

```vb
Dim oldHashResult As String
```

### 8.2. Заменить цикл по именованным формам

Перед первым проходом добавить:

```vb
Dim lastFormIndex As Long
lastFormIndex = CVD_LastFormIndexOnSheet(sheetCVD)
```

Заменить:

```vb
i = 1
Do While NamedRangeExists(wb, "CVD_" & i & "_sampleCode")
```

на:

```vb
For i = 1 To lastFormIndex
```

В конце первого прохода заменить:

```vb
NextFormFirstPass:
    i = i + 1
Loop
```

на:

```vb
NextFormFirstPass:
Next i
```

### 8.3. Проверка отменённого задания

Заменить блок с `NamedRangeExists(...TaskId)` на:

```vb
Dim taskRange As Range

Set taskRange = CVD_FormRange( _
    sheetCVD, i, "TaskId" _
)

If Not taskRange Is Nothing Then
    If taskRange.Interior.ColorIndex = 3 Then
        GoTo NextFormFirstPass
    End If
End If
```

### 8.4. Номер анализа

Полностью удалить:

```vb
If sheetCVD.Range("CVD_" & i & "_analyseNumber").value = "" Then
    Call checkNumberAnalys2(sheetCVD.Range("CVD_" & i & "_sampleCode").value, "CVD")
End If
```

Вместо него вызвать:

```vb
EnsureCVDAnalyseNumber sheetCVD, i
```

### 8.5. Проверка шифров полученных проб

В цикле step вместо строк `lphRangeName` и `gphRangeName` использовать диапазоны:

```vb
Dim lphRange As Range
Dim gphRange As Range

Set lphRange = CVD_FormRange( _
    sheetCVD, _
    i, _
    stepKey & "_sampleCodeLPh" _
)

Set gphRange = CVD_FormRange( _
    sheetCVD, _
    i, _
    stepKey & "_sampleCodeGPh" _
)

If Not lphRange Is Nothing _
   And Not gphRange Is Nothing Then

    If Len(Trim$(CStr(lphRange.Value))) = 0 _
       And Len(Trim$(CStr(gphRange.Value))) = 0 Then

        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False

        If Not CreateSampleCodeCVD3(i, stepKey) Then
            MsgBox _
                "Невозможно синхронизировать данные " & _
                "без шифров полученных проб.", _
                vbExclamation

            Save_CVD = MACRO_FAILED
            Exit Function
        End If

    End If

End If
```

### 8.6. Чтение `resultIdCVD`

Заменить прямой `Range(...)`:

```vb
Dim resultIdRange As Range

Set resultIdRange = CVD_FormRange( _
    sheetCVD, _
    i, _
    COL_CVD_RESULT_ID_DB _
)

If resultIdRange Is Nothing Then
    resultIdValue = vbNullString
Else
    resultIdValue = resultIdRange.Value
End If
```

### 8.7. Формирование родительской строки

В цикле по заголовкам `CVD_results` больше не создавать `formRangeName` для доступа к листу. Ветка `Case Else` должна выглядеть так:

```vb
Dim formFieldRange As Range

' ...внутри цикла по headerCell:
Set formFieldRange = Nothing
Set formFieldRange = CVD_FormRange( _
    sheetCVD, _
    i, _
    headerName _
)

If headerName = "dateTimeSync" Then

    parentRow(parentColIndex) = _
        Format$(Now, "yyyy-mm-dd hh:nn:ss")

ElseIf headerName = COL_CVD_RESULT_ID_DB Then

    parentRow(parentColIndex) = parentKey

ElseIf Not formFieldRange Is Nothing Then

    parentRow(parentColIndex) = _
        FormatValueForSyncText( _
            headerName, _
            formFieldRange.Value _
        )

Else

    parentRow(parentColIndex) = vbNullString

End If
```

### 8.8. Чтение `rowIdCVD`

Заменить `rowIdRangeName` и `NamedRangeExists` на:

```vb
Dim rowIdRange As Range

Set rowIdRange = CVD_FormRange( _
    sheetCVD, _
    i, _
    stepKey & "_" & COL_CVD_ROW_ID_DB _
)

If Not rowIdRange Is Nothing Then
    oldRowIdCVD = _
        NormalizeKeyValue(rowIdRange.Value)
End If
```

### 8.9. Формирование дочерней строки

В `Case Else` вместо `childRangeName` и `NamedRangeExists`:

```vb
Dim childFieldRange As Range

If hasStepData Then

    Set childFieldRange = Nothing
    Set childFieldRange = CVD_FormRange( _
        sheetCVD, _
        i, _
        stepKey & "_" & childHeaderName _
    )

    If Not childFieldRange Is Nothing Then
        childRow(childColIndex) = _
            FormatValueForSyncText( _
                childHeaderName, _
                childFieldRange.Value _
            )
    Else
        childRow(childColIndex) = vbNullString
    End If

Else
    childRow(childColIndex) = vbNullString
End If
```

## 9. Изменения в `Build_CVD_HashData`

### 9.1. Цикл форм

Заменить цикл `Do While NamedRangeExists(...)` на:

```vb
Dim lastFormIndex As Long
lastFormIndex = CVD_LastFormIndexOnSheet(sheetCVD)

For i = 1 To lastFormIndex
```

В конце заменить:

```vb
NextForm:
    i = i + 1
Loop
```

на:

```vb
NextForm:
Next i
```

### 9.2. Отменённое задание

Использовать:

```vb
Dim taskRange As Range

Set taskRange = CVD_FormRange( _
    sheetCVD, i, "TaskId" _
)

If Not taskRange Is Nothing Then
    If taskRange.Interior.ColorIndex = 3 Then
        GoTo NextForm
    End If
End If
```

### 9.3. Родительские поля

Вместо `rangeName`, `NamedRangeExists` и `sheetCVD.Range(rangeName)`:

```vb
Dim fieldRange As Range

Set fieldRange = Nothing
Set fieldRange = CVD_FormRange( _
    sheetCVD, i, headerName _
)

If Not fieldRange Is Nothing Then
    formatValue = FormatValueForSyncText( _
        headerName, _
        fieldRange.Value _
    )
Else
    formatValue = vbNullString
End If
```

### 9.4. Дочерние поля

Вместо `childRangeName`, `NamedRangeExists` и `sheetCVD.Range(childRangeName)`:

```vb
Dim childFieldRange As Range

Set childFieldRange = Nothing
Set childFieldRange = CVD_FormRange( _
    sheetCVD, _
    i, _
    stepKey & "_" & childHeaderName _
)

If Not childFieldRange Is Nothing Then
    formatValue = FormatValueForSyncText( _
        childHeaderName, _
        childFieldRange.Value _
    )
Else
    formatValue = vbNullString
End If
```

## 10. Переделать загрузку таблиц в формы

Общие функции `FindFormIndexForResultRow` и `CopyTableRowToNamedRanges` использовать для CVD больше нельзя: они рассчитывают на физические имена каждой формы.

### 10.1. Изменения в `Load_CVD_results_to_forms`

После создания `fieldFilter` добавить:

```vb
Dim lastFormIndex As Long
lastFormIndex = CVD_LastFormIndexOnSheet(sheetCVD)
```

В режиме `sequential` оставить создание недостающих форм:

```vb
If mode = "sequential" Then

    EnsureCVDFormExists wb, resultRowIndex
    lastFormIndex = CVD_LastFormIndexOnSheet(sheetCVD)

    If resultRowIndex <= lastFormIndex Then
        formIndex = resultRowIndex
    Else
        formIndex = 0
    End If

Else

    formIndex = CVD_FindFormIndexForResultRow( _
        sheetCVD, _
        resultTable, _
        resultRowIndex, _
        Array("sampleCode", "analyseNumber") _
    )

End If
```

Проверку:

```vb
If formIndex = 0 Or _
   Not NamedRangeExists(...) Then
```

заменить на:

```vb
If formIndex <= 0 _
   Or Not CVD_FormExists(formIndex) Then
```

### 10.2. Полностью заменить `EnsureCVDFormExists`

```vb
Private Sub EnsureCVDFormExists( _
    ByVal wb As Workbook, _
    ByVal formIndex As Long _
)

    Dim previousLastForm As Long

    Do While formIndex > CVD_GetLastFormIndex()

        previousLastForm = CVD_GetLastFormIndex()
        AddFormCVD

        If CVD_GetLastFormIndex() <= previousLastForm Then
            Err.Raise _
                vbObjectError + 4206, _
                "EnsureCVDFormExists", _
                "Не удалось добавить форму CVD."
        End If

    Loop

End Sub
```

Параметр `wb` оставлен, чтобы не менять существующий вызов. Внутри функции он не требуется.

### 10.3. Добавить поиск формы по ключам

Добавить перед `FillOneCVDFormFromResultRow`:

```vb
Private Function CVD_FindFormIndexForResultRow( _
    ByVal ws As Worksheet, _
    ByVal resultTable As ListObject, _
    ByVal resultRowIndex As Long, _
    ByVal keyFields As Variant _
) As Long

    Dim fieldIndex As Long
    Dim fieldName As String
    Dim resultValue As String

    For fieldIndex = LBound(keyFields) To UBound(keyFields)

        fieldName = CStr(keyFields(fieldIndex))

        resultValue = NormalizeKeyValue( _
            GetTableValue( _
                resultTable, _
                resultRowIndex, _
                fieldName _
            ) _
        )

        If Len(resultValue) = 0 Then Exit Function

    Next fieldIndex

    Dim formIndex As Long
    Dim lastFormIndex As Long
    Dim formValue As String
    Dim allMatched As Boolean
    Dim fieldRange As Range

    lastFormIndex = CVD_LastFormIndexOnSheet(ws)

    For formIndex = 1 To lastFormIndex

        allMatched = True

        For fieldIndex = LBound(keyFields) To UBound(keyFields)

            fieldName = CStr(keyFields(fieldIndex))

            Set fieldRange = Nothing
            Set fieldRange = CVD_FormRange( _
                ws, formIndex, fieldName _
            )

            If fieldRange Is Nothing Then
                allMatched = False
                Exit For
            End If

            resultValue = NormalizeKeyValue( _
                GetTableValue( _
                    resultTable, _
                    resultRowIndex, _
                    fieldName _
                ) _
            )

            formValue = _
                NormalizeKeyValue(fieldRange.Value)

            If formValue <> resultValue Then
                allMatched = False
                Exit For
            End If

        Next fieldIndex

        If allMatched Then
            CVD_FindFormIndexForResultRow = formIndex
            Exit Function
        End If

    Next formIndex

End Function
```

### 10.4. Изменить `FillOneCVDFormFromResultRow`

Вместо вызова `CopyTableRowToNamedRanges` вызвать:

```vb
CVD_CopyTableRowToForm _
    sheetCVD, _
    resultTable, _
    resultRowIndex, _
    formIndex, _
    fieldFilter
```

Добавить следующую функцию:

```vb
Private Sub CVD_CopyTableRowToForm( _
    ByVal ws As Worksheet, _
    ByVal sourceTable As ListObject, _
    ByVal sourceRowIndex As Long, _
    ByVal formIndex As Long, _
    ByVal fieldFilter As Object _
)

    If sourceTable.DataBodyRange Is Nothing Then Exit Sub

    Dim tableColumn As ListColumn
    Dim fieldName As String
    Dim targetRange As Range

    For Each tableColumn In sourceTable.ListColumns

        fieldName = CStr(tableColumn.name)

        If ShouldTransferField(fieldFilter, fieldName) Then

            Set targetRange = Nothing
            Set targetRange = CVD_FormRange( _
                ws, formIndex, fieldName _
            )

            If Not targetRange Is Nothing Then
                If Not RangeHasFormula(targetRange) Then
                    targetRange.Value = _
                        sourceTable.DataBodyRange.Cells( _
                            sourceRowIndex, _
                            tableColumn.index _
                        ).Value
                End If
            End If

        End If

    Next tableColumn

End Sub
```

### 10.5. Полностью заменить `FillCVDSourceRangesFromSourceData`

```vb
Private Sub FillCVDSourceRangesFromSourceData( _
    ByVal sheetCVD As Worksheet, _
    ByVal sourceTable As ListObject, _
    ByVal formIndex As Long, _
    ByVal resultId As Variant, _
    ByVal fieldFilter As Object _
)

    If sourceTable.DataBodyRange Is Nothing Then Exit Sub

    Dim steps As Variant
    steps = CVD_StepKeys()

    Dim stepIndex As Long
    Dim stepKey As String
    Dim headerCell As Range
    Dim headerName As String
    Dim targetRange As Range

    ' Сначала очищаем переносимые поля всех step.
    For stepIndex = LBound(steps) To UBound(steps)

        stepKey = CStr(steps(stepIndex))

        For Each headerCell In sourceTable.HeaderRowRange.Cells

            headerName = CStr(headerCell.Value)

            If headerName <> COL_CVD_RESULT_ID_DB _
               And headerName <> COL_CVD_STEP_DB Then

                If ShouldTransferField( _
                    fieldFilter, headerName _
                ) Then

                    Set targetRange = Nothing
                    Set targetRange = CVD_FormRange( _
                        sheetCVD, _
                        formIndex, _
                        stepKey & "_" & headerName _
                    )

                    If Not targetRange Is Nothing Then
                        If Not RangeHasFormula(targetRange) Then
                            targetRange.Value = vbNullString
                        End If
                    End If

                End If

            End If

        Next headerCell

    Next stepIndex

    Dim rowIndex As Long
    Dim stepValue As Variant

    For rowIndex = 1 To sourceTable.ListRows.count

        If NormalizeKeyValue( _
            GetTableValue( _
                sourceTable, _
                rowIndex, _
                COL_CVD_RESULT_ID_DB _
            ) _
        ) <> NormalizeKeyValue(resultId) Then
            GoTo NextSourceRow
        End If

        stepValue = GetTableValue( _
            sourceTable, _
            rowIndex, _
            COL_CVD_STEP_DB _
        )

        stepKey = CVD_StepKeyFromDbValue(stepValue)
        If Len(stepKey) = 0 Then GoTo NextSourceRow

        For Each headerCell In sourceTable.HeaderRowRange.Cells

            headerName = CStr(headerCell.Value)

            If ShouldTransferField( _
                fieldFilter, headerName _
            ) Then

                Set targetRange = Nothing
                Set targetRange = CVD_FormRange( _
                    sheetCVD, _
                    formIndex, _
                    stepKey & "_" & headerName _
                )

                If Not targetRange Is Nothing Then
                    If Not RangeHasFormula(targetRange) Then
                        targetRange.Value = GetTableValue( _
                            sourceTable, _
                            rowIndex, _
                            headerName _
                        )
                    End If
                End If

            End If

        Next headerCell

NextSourceRow:
    Next rowIndex

End Sub
```

## 11. Полностью заменить `Get_CVD_FormState`

```vb
Private Function Get_CVD_FormState( _
    ByVal formIndex As Long _
) As CVDFormState

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(CVD_FORM_SHEET)

    Dim sampleRange As Range
    Set sampleRange = CVD_FormRange( _
        ws, formIndex, "sampleCode" _
    )

    If sampleRange Is Nothing Then
        Get_CVD_FormState = CVD_FORM_EMPTY
        Exit Function
    End If

    Dim hasSampleCode As Boolean
    hasSampleCode = _
        (Len(Trim$(CStr(sampleRange.Value))) > 0)

    Dim hasAnyStepData As Boolean
    Dim steps As Variant
    Dim stepIndex As Long

    steps = CVD_StepKeys()

    For stepIndex = LBound(steps) To UBound(steps)

        If CVD_StepHasData( _
            formIndex, _
            CStr(steps(stepIndex)) _
        ) Then

            hasAnyStepData = True
            Exit For

        End If

    Next stepIndex

    If Not hasSampleCode And Not hasAnyStepData Then
        Get_CVD_FormState = CVD_FORM_EMPTY
    ElseIf hasSampleCode And hasAnyStepData Then
        Get_CVD_FormState = CVD_FORM_FILLED
    Else
        Get_CVD_FormState = CVD_FORM_RESERVED
    End If

End Function
```

## 12. Изменения в `Validate_CVD_Form`

Полностью переписывать валидацию не требуется, но все обращения к диапазонам нужно перевести на `CVD_FormRange`.

### 12.1. Проверка существования формы

Заменить:

```vb
If Not NamedRangeExists(wb, "CVD_" & formIndex & "_sampleCode") Then
```

на:

```vb
If Not CVD_FormExists(formIndex) Then
```

### 12.2. Родительские поля

В цикле по `CVDTable.HeaderRowRange.Cells` оставить формирование `rangeName` только для текста ошибки:

```vb
rangeName = CVD_RangeLabel(formIndex, fieldName)

Set rng = Nothing
Set rng = CVD_FormRange( _
    sheetCVD, formIndex, fieldName _
)

If rng Is Nothing Then
    errorCount = errorCount + 1
    errors = errors & errorCount & _
             ". Не найден базовый диапазон CVD_1_" & _
             fieldName & vbCrLf
    GoTo next_result_header
End If
```

Старые `NamedRangeExists(wb, rangeName)` и `Set rng = sheetCVD.Range(rangeName)` удалить.

### 12.3. Проверка `TaskId`

В текущем коде ссылка на температуру собрана с ошибкой:

```vb
"CVD_" & formIndex & "CVD_" & formIndex & "_Ttransfer"
```

Вместо прямых `Range(...)` использовать:

```vb
Dim sampleCodeRange As Range
Dim transferTemperatureRange As Range

Set sampleCodeRange = CVD_FormRange( _
    sheetCVD, formIndex, "sampleCode" _
)

Set transferTemperatureRange = CVD_FormRange( _
    sheetCVD, formIndex, "Ttransfer" _
)

If Not sampleCodeRange Is Nothing _
   And Not transferTemperatureRange Is Nothing Then

    taskCheckResult = CheckCVDTask( _
        CLng(rng.Value), _
        CStr(sampleCodeRange.Value), _
        transferTemperatureRange.Value, _
        formIndex _
    )

    If taskCheckResult <> "OK" Then
        AddValidationError _
            errors, _
            errorCount, _
            taskCheckResult, _
            rng
    End If

End If
```

### 12.4. Дочерние поля

В цикле по `CVDSourceDataTable.HeaderRowRange.Cells`:

```vb
rangeName = CVD_RangeLabel( _
    formIndex, _
    stepKey & "_" & fieldName _
)

Set rng = Nothing
Set rng = CVD_FormRange( _
    sheetCVD, _
    formIndex, _
    stepKey & "_" & fieldName _
)

If rng Is Nothing Then
    errorCount = errorCount + 1
    errors = errors & errorCount & _
             ". Не найден базовый диапазон CVD_1_" & _
             stepKey & "_" & fieldName & vbCrLf
    GoTo next_source_header
End If
```

Старые `NamedRangeExists` и `sheetCVD.Range(rangeName)` удалить.

### 12.5. Проверка пары `sampleCodeLPh/sampleCodeGPh`

Заменить блок прямых обращений на:

```vb
Dim lphRange As Range
Dim gphRange As Range

Set lphRange = CVD_FormRange( _
    sheetCVD, _
    formIndex, _
    stepKey & "_sampleCodeLPh" _
)

Set gphRange = CVD_FormRange( _
    sheetCVD, _
    formIndex, _
    stepKey & "_sampleCodeGPh" _
)

lphText = vbNullString
gphText = vbNullString

If Not lphRange Is Nothing Then
    lphText = Trim$(CStr(lphRange.Value))
End If

If Not gphRange Is Nothing Then
    gphText = Trim$(CStr(gphRange.Value))
End If
```

### 12.6. Связанная проверка `mgas/Vgas/VGTot`

Заменить весь блок с тремя `NamedRangeExists` на:

```vb
Dim mgasRange As Range
Dim vgasRange As Range
Dim vgtotRange As Range

Set mgasRange = CVD_FormRange( _
    sheetCVD, formIndex, stepKey & "_mgas" _
)

Set vgasRange = CVD_FormRange( _
    sheetCVD, formIndex, stepKey & "_Vgas" _
)

Set vgtotRange = CVD_FormRange( _
    sheetCVD, formIndex, stepKey & "_VGTot" _
)

If Not mgasRange Is Nothing _
   And Not vgasRange Is Nothing _
   And Not vgtotRange Is Nothing Then

    If Len(Trim$(CStr(mgasRange.Value))) > 0 Then

        If Len(Trim$(CStr(vgasRange.Value))) = 0 _
           And Len(Trim$(CStr(vgtotRange.Value))) = 0 Then

            errorCount = errorCount + 1
            errors = errors & errorCount & _
                     ". В CVD_" & formIndex & _
                     "_" & stepKey & _
                     " указана mgas, но не указан " & _
                     "Vgas/VGTot." & vbCrLf

        End If

    End If

End If
```

## 13. Полностью заменить `CVD_StepHasData`

```vb
Private Function CVD_StepHasData( _
    ByVal formIndex As Long, _
    ByVal stepKey As String _
) As Boolean

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(CVD_FORM_SHEET)

    Dim keyFields As Variant

    keyFields = Array( _
        "operator", _
        "date", _
        "Ppic", _
        "Vcell", _
        "VGTot", _
        "VLiq", _
        "sampleCodeLPh", _
        "sampleCodeGPh" _
    )

    Dim fieldIndex As Long
    Dim targetRange As Range
    Dim fieldValue As Variant

    For fieldIndex = LBound(keyFields) To UBound(keyFields)

        Set targetRange = Nothing
        Set targetRange = CVD_FormRange( _
            ws, _
            formIndex, _
            stepKey & "_" & CStr(keyFields(fieldIndex)) _
        )

        If Not targetRange Is Nothing Then

            fieldValue = targetRange.Value

            If IsError(fieldValue) Then
                CVD_StepHasData = True
                Exit Function
            End If

            If Len(Trim$(CStr(fieldValue))) > 0 Then
                CVD_StepHasData = True
                Exit Function
            End If

        End If

    Next fieldIndex

End Function
```

## 14. Полностью заменить `FindEmptyFormCVD`

```vb
Public Function FindEmptyFormCVD() As Long

    Dim lastFormIndex As Long
    lastFormIndex = CVD_GetLastFormIndex()

    Dim formIndex As Long

    For formIndex = 1 To lastFormIndex

        If Get_CVD_FormState(formIndex) = _
           CVD_FORM_EMPTY Then

            FindEmptyFormCVD = formIndex
            Exit Function

        End If

    Next formIndex

    AddFormCVD

    If CVD_GetLastFormIndex() > lastFormIndex Then
        FindEmptyFormCVD = lastFormIndex + 1
    Else
        FindEmptyFormCVD = 0
    End If

End Function
```

## 15. Исправить вызовы кнопок в `wrappers.bas`

В присланной версии обёртки CVD вызывают отсутствующие процедуры `InsertListCVD2` и `DeleteLastListCVD2`.

Заменить обе обёртки на следующий код:

```vb
Public Sub Add_CVD_Form_wrap()

    CallMacro "AddFormCVD", "CVD"

End Sub


Public Sub Delete_CVD_Form_wrap()

    CallMacro "DeleteLastListCVD", "CVD"

End Sub
```

## 16. Что делать с `CVD_dev.bas`

Методы `CreateShiftedNamedRanges` и `CreateStepNames` после перехода больше не использовать. Они снова создадут имена `CVD_2_*` и далее.

Сам `CVD_dev.bas` можно оставить как служебный модуль, но перед релизом лучше:

- удалить/закомментировать вызовы этих процедур, если они где-то есть;
- не размещать их на Ribbon;
- добавить комментарий `OBSOLETE: формы 2+ используют смещение`.

`CreateCVDTableColumns` можно оставить: он специально читает только `CVD_1_*`.

## 17. Процедуры, которые менять не требуется

После перечисленных исправлений без изменений остаются:

- `editCurrentSamplesFormCVD`;
- `CalculateCurrentCVDHash`;
- `CheckCVDTask`;
- `CVD_FieldCaption`;
- `CVD_StepKeys`;
- `CVD_StepKeyFromDbValue`;
- `CVD_CalculateCommonTaskStatus`;
- `Validate_CVD`;
- `Load_CVD`.

## 18. Финальная проверка перед удалением имён форм 2+

Сначала внести код, скомпилировать и проверить его на книге, в которой старые имена ещё существуют. Только после этого удалить `CVD_2_*`, `CVD_3_*` и далее.

### Поиск по исходнику

После переделки в `CVD.bas` не должно остаться рабочих обращений:

```text
NamedRangeExists(wb, "CVD_"
sheetCVD.Range("CVD_"
sheetCVD.Range(rangeName)
CVD_2_sampleCode
CopyTableRowToNamedRanges
FindFormIndexForResultRow
checkNumberAnalys2
ActiveCell.Name.Name
```

Допустимое прямое обращение к имени остаётся только внутри `CVD_FormRange`:

```vb
ws.Parent.Names("CVD_1_" & fieldKey).RefersToRange
```

### Проверка формы

1. Сверить `CVD_FORM_COLUMN_COUNT`, `CVD_PAGE_*` и `CVD_*_STEP_COL`.
2. Выполнить `Debug -> Compile VBAProject`.
3. На первой форме проверить сохранение, загрузку, валидацию и отмену.
4. На второй форме повторить те же операции после удаления её имён.
5. Добавить новую форму:
   - скопирован ровно один блок столбцов;
   - формулы сохранились;
   - введённые значения и старые ID очищены;
   - номер страницы увеличился на 1;
   - имена `CVD_2_*`/`CVD_3_*` не создались.
6. Удалить последнюю форму и убедиться, что удалился ровно один блок.
7. Проверить `FindEmptyFormCVD` через обновление проекта.
8. Проверить загрузку в режимах `by_key` и `sequential`.
9. Проверить повторный `sampleCode`: номера анализов должны стать `1`, `2`, ... без `checkNumberAnalys2`.
10. Проверить формулы новой формы. Относительные ссылки Excel сдвинет автоматически, но формула, явно содержащая имя `CVD_1_*`, продолжит ссылаться на первую форму. Такие формулы нужно переделать на относительные адреса или отдельную логику смещения.

После успешной проверки имена форм 2+ можно удалить. Имена `CVD_1_*` удалять нельзя.
