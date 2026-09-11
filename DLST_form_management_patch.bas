' ФРАГМЕНТ ДЛЯ ВСТАВКИ В СУЩЕСТВУЮЩИЙ МОДУЛЬ DLST.bas.
' Это не отдельный VBA-модуль: код использует Private-константы и функции DLST.
'
' 1. Рядом с DLST_MAX_FORMS добавить:
Private Const DLST_MIN_FORM_COUNT As Long = 1
'
' 2. Удалить существующие заглушки AddFormDLST и DeleteLastFormDLST.
' 3. Вставить вместо них весь код ниже.


Public Sub AddFormDLST()

    On Error GoTo CleanFail

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        Err.Raise _
            vbObjectError + 4300, _
            "AddFormDLST", _
            "Нет активной книги с формой DLST."
    End If

    Dim ws As Worksheet
    Set ws = wb.Worksheets(DLST_FORM_SHEET)

    Dim oldFormIndex As Long
    Dim newFormIndex As Long

    oldFormIndex = DLST_LastFormIndexOnSheet(ws)

    If oldFormIndex < DLST_MIN_FORM_COUNT Then
        Err.Raise _
            vbObjectError + 4301, _
            "AddFormDLST", _
            "Не найдена первая форма DLST."
    End If

    If oldFormIndex >= DLST_MAX_FORMS Then
        Err.Raise _
            vbObjectError + 4302, _
            "AddFormDLST", _
            "Достигнуто максимальное число форм DLST."
    End If

    newFormIndex = oldFormIndex + 1

    Dim sourceFirstColumn As Long
    Dim sourceLastColumn As Long
    Dim targetFirstColumn As Long
    Dim targetLastColumn As Long

    sourceFirstColumn = _
        DLST_FIRST_FORM_FIRST_COL + _
        (oldFormIndex - 1) * DLST_FORM_COLUMN_COUNT

    sourceLastColumn = _
        sourceFirstColumn + DLST_FORM_COLUMN_COUNT - 1

    targetFirstColumn = _
        sourceFirstColumn + DLST_FORM_COLUMN_COUNT

    targetLastColumn = _
        targetFirstColumn + DLST_FORM_COLUMN_COUNT - 1

    If targetLastColumn > ws.Columns.count Then
        Err.Raise _
            vbObjectError + 4303, _
            "AddFormDLST", _
            "Для новой формы DLST недостаточно столбцов листа."
    End If

    Dim sourceColumns As Range
    Set sourceColumns = ws.Range( _
        ws.Columns(sourceFirstColumn), _
        ws.Columns(sourceLastColumn) _
    )

    Dim oldCopyObjectsWithCells As Boolean
    Dim copySettingChanged As Boolean

    oldCopyObjectsWithCells = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = False
    copySettingChanged = True

    sourceColumns.Copy

    ws.Columns(targetFirstColumn).PasteSpecial Paste:=xlPasteAll
    ws.Columns(targetFirstColumn).PasteSpecial _
        Paste:=xlPasteColumnWidths

    Application.CutCopyMode = False

    Application.CopyObjectsWithCells = oldCopyObjectsWithCells
    copySettingChanged = False

    ' Именованные диапазоны для новой формы не создаются.
    ' DLST_FormRange вычисляет их позиции по DLST_1_* и смещению.
    DLST_ClearCopiedFormFields wb, ws, newFormIndex

    Dim pageCell As Range
    Set pageCell = DLST_PageCell(ws, newFormIndex)

    If pageCell Is Nothing Then
        Err.Raise _
            vbObjectError + 4304, _
            "AddFormDLST", _
            "Не удалось определить ячейку страницы новой формы DLST."
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
        Application.CopyObjectsWithCells = oldCopyObjectsWithCells
    End If

    On Error GoTo 0

    Debug.Print _
        "AddFormDLST: " & errNumber & ": " & errDescription

    Err.Raise errNumber, errSource, errDescription

End Sub


Public Sub DeleteLastFormDLST()

    On Error GoTo CleanFail

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = wb.Worksheets(DLST_FORM_SHEET)

    Dim lastFormIndex As Long
    lastFormIndex = DLST_LastFormIndexOnSheet(ws)

    ' Первую базовую форму удалять нельзя: от неё считаются все смещения.
    If lastFormIndex <= DLST_MIN_FORM_COUNT Then
        Debug.Print "DeleteLastFormDLST: первая форма не удаляется."
        Exit Sub
    End If

    Dim firstColumn As Long
    Dim lastColumn As Long

    firstColumn = _
        DLST_FIRST_FORM_FIRST_COL + _
        (lastFormIndex - 1) * DLST_FORM_COLUMN_COUNT

    lastColumn = firstColumn + DLST_FORM_COLUMN_COUNT - 1

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
        "DeleteLastFormDLST: " & _
        errNumber & ": " & errDescription

    Err.Raise errNumber, errSource, errDescription

End Sub


Private Sub DLST_ClearCopiedFormFields( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long _
)

    Dim resultTable As ListObject
    Dim sourceTable As ListObject

    Set resultTable = _
        wb.Worksheets(DLST_RESULTS_SHEET) _
          .ListObjects(DLST_RESULTS_TABLE)

    Set sourceTable = _
        wb.Worksheets(DLST_SOURCE_SHEET) _
          .ListObjects(DLST_SOURCE_TABLE)

    Dim tableColumn As ListColumn
    Dim fieldName As String
    Dim targetRange As Range

    ' Родительские поля DLST_results.
    For Each tableColumn In resultTable.ListColumns

        fieldName = Trim$(CStr(tableColumn.name))

        If StrComp( _
            fieldName, _
            "dateTimeSync", _
            vbTextCompare _
        ) <> 0 Then

            Set targetRange = Nothing
            Set targetRange = _
                DLST_FormRange(ws, formIndex, fieldName)

            DLST_ClearNonFormulaCells targetRange

        End If

    Next tableColumn

    ' Дочерние поля DLST_sourceData для st1..st10 и last.
    Dim steps As Variant
    Dim stepIndex As Long
    Dim stepKey As String

    steps = DLST_StepKeys()

    For stepIndex = LBound(steps) To UBound(steps)

        stepKey = CStr(steps(stepIndex))

        For Each tableColumn In sourceTable.ListColumns

            fieldName = Trim$(CStr(tableColumn.name))

            Select Case LCase$(fieldName)

                Case LCase$(COL_DLST_RESULT_ID_DB), _
                     LCase$(COL_DLST_STEP_DB)

                    ' Эти значения вычисляются при сохранении и
                    ' отдельного диапазона на форме не имеют.

                Case Else

                    Set targetRange = Nothing
                    Set targetRange = DLST_FormRange( _
                        ws, _
                        formIndex, _
                        stepKey & "_" & fieldName _
                    )

                    DLST_ClearNonFormulaCells targetRange

            End Select

        Next tableColumn

    Next stepIndex

End Sub


Private Sub DLST_ClearNonFormulaCells( _
    ByVal targetRange As Range _
)

    If targetRange Is Nothing Then Exit Sub

    ' Объединённую ячейку можно очищать только целиком.
    If targetRange.MergeCells Then

        Dim mergedRange As Range
        Set mergedRange = targetRange.MergeArea

        If Not RangeHasFormula(mergedRange) Then
            mergedRange.ClearContents
        End If

        Exit Sub

    End If

    ' Для обычного многоклеточного диапазона сохраняем каждую формулу,
    ' а константы и введённые пользователем значения очищаем.
    Dim cell As Range

    For Each cell In targetRange.Cells
        If Not cell.HasFormula Then
            cell.ClearContents
        End If
    Next cell

End Sub
