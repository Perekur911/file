Attribute VB_Name = "Stable"
Option Explicit

' Имена листов, таблицы и диапазонов:
'   лист формы       Stable
'   лист результатов Stable_results
'   таблица           Stable_results
'   базовые диапазоны Stable_1_<имя поля>

Private Enum StableFormState
    STABLE_FORM_EMPTY = 0
    STABLE_FORM_RESERVED = 1
    STABLE_FORM_FILLED = 2
End Enum

Private Const STABLE_STUDY_CODE As String = "STABLE"
Private Const STABLE_TASK_TYPE As String = "Stable"
Private Const STABLE_SHEET_NAME As String = "Stable"
Private Const STABLE_RESULTS_SHEET_NAME As String = "Stable_results"
Private Const STABLE_RESULTS_TABLE_NAME As String = "Stable_results"
Private Const STABLE_RANGE_PREFIX As String = "Stable_"

Private Const COL_RESULT_ID_DB As String = "resultIdStable"
Private Const COL_TASK_ID_DB As String = "TaskId"
Private Const COL_SAMPLE_CODE_DB As String = "sampleCode"
Private Const COL_DATE_TIME_SYNC_DB As String = "dateTimeSync"

' Геометрия формы вынесена в константы.
' Если Stable расположен иначе, достаточно исправить значения ниже.
Private Const STABLE_FORM_COLUMN_COUNT As Long = 3
Private Const STABLE_FIRST_FORM_FIRST_COL As Long = 3
Private Const STABLE_PAGE_ROW As Long = 3
Private Const STABLE_PAGE_FIRST_COL As Long = 3
Private Const STABLE_MIN_FORM_COUNT As Long = 2
Private Const STABLE_MAX_FORMS As Long = 999


' Именованные диапазоны существуют только для первой формы:
' Stable_1_<имя поля>. Остальные формы вычисляются сдвигом вправо.
Public Function GetStableFormRange( _
    ByVal formIndex As Long, _
    ByVal fieldKey As String _
) As Range

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(STABLE_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    Set GetStableFormRange = _
        Stable_FormRange(ws, formIndex, fieldKey)

End Function


Private Function Stable_FormRange( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long, _
    ByVal fieldKey As String _
) As Range

    If ws Is Nothing Then Exit Function
    If formIndex <= 0 Then Exit Function
    If Len(Trim$(fieldKey)) = 0 Then Exit Function

    Dim baseRange As Range
    Dim columnShift As Long

    On Error Resume Next
    Set baseRange = ws.Parent.Names( _
        STABLE_RANGE_PREFIX & "1_" & fieldKey _
    ).RefersToRange
    On Error GoTo 0

    If baseRange Is Nothing Then Exit Function
    If Not baseRange.Worksheet Is ws Then Exit Function

    columnShift = _
        (formIndex - 1) * STABLE_FORM_COLUMN_COUNT

    If baseRange.Column + columnShift > ws.Columns.count Then
        Exit Function
    End If

    On Error Resume Next
    Set Stable_FormRange = baseRange.Offset(0, columnShift)
    On Error GoTo 0

End Function


Private Function Stable_TryGetFormRange( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long, _
    ByVal fieldKey As String, _
    ByRef resultRange As Range _
) As Boolean

    Set resultRange = Nothing
    Set resultRange = _
        Stable_FormRange(ws, formIndex, fieldKey)

    Stable_TryGetFormRange = Not resultRange Is Nothing

End Function


Private Function Stable_PageCell( _
    ByVal ws As Worksheet, _
    ByVal formIndex As Long _
) As Range

    If ws Is Nothing Then Exit Function
    If formIndex <= 0 Then Exit Function

    Dim targetColumn As Long

    targetColumn = STABLE_PAGE_FIRST_COL + _
        (formIndex - 1) * STABLE_FORM_COLUMN_COUNT

    If targetColumn > ws.Columns.count Then Exit Function

    Set Stable_PageCell = _
        ws.Cells(STABLE_PAGE_ROW, targetColumn)

End Function


Public Function Stable_GetLastFormIndex() As Long

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(STABLE_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    Stable_GetLastFormIndex = _
        Stable_LastFormIndexOnSheet(ws)

End Function


Private Function Stable_LastFormIndexOnSheet( _
    ByVal ws As Worksheet _
) As Long

    Dim formIndex As Long
    Dim pageCell As Range
    Dim pageValue As Variant

    For formIndex = 1 To STABLE_MAX_FORMS

        Set pageCell = Stable_PageCell(ws, formIndex)

        If pageCell Is Nothing Then Exit For

        pageValue = pageCell.Value2

        If IsError(pageValue) Then Exit For
        If Len(Trim$(CStr(pageValue))) = 0 Then Exit For

        Stable_LastFormIndexOnSheet = formIndex

    Next formIndex

End Function


Public Function Stable_FormExists( _
    ByVal formIndex As Long _
) As Boolean

    If formIndex <= 0 Then Exit Function

    Stable_FormExists = _
        (formIndex <= Stable_GetLastFormIndex())

End Function


Public Function Validate_Stable() As macroResult
    Validate_Stable = Save_Stable("Валидировано")
End Function


Public Sub Load_Stable()
    LoadResearchesFromDB STABLE_STUDY_CODE
End Sub


Public Function cancelCurrentTaskStable() As macroResult

    cancelCurrentTaskStable = MACRO_FAILED

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Function

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    If Not ActiveSheet Is sheetStable Then
        MsgBox "Для отмены задания выберите ячейку на листе Stable.", vbExclamation
        Exit Function
    End If

    If ActiveCell.Column < STABLE_FIRST_FORM_FIRST_COL Then
        MsgBox "Не удалось определить форму Stable по выбранной ячейке.", vbExclamation
        Exit Function
    End If

    Dim formIndex As Long
    formIndex = _
        (ActiveCell.Column - STABLE_FIRST_FORM_FIRST_COL) \ _
        STABLE_FORM_COLUMN_COUNT + 1

    If formIndex > Stable_LastFormIndexOnSheet(sheetStable) Then
        MsgBox "Форма Stable_" & formIndex & " не найдена.", vbExclamation
        Exit Function
    End If

    Dim taskCell As Range
    Set taskCell = Stable_FormRange( _
        sheetStable, _
        formIndex, _
        COL_TASK_ID_DB _
    )

    If taskCell Is Nothing Then
        MsgBox _
            "Не найден базовый диапазон Stable_1_" & _
            COL_TASK_ID_DB & ".", _
            vbExclamation
        Exit Function
    End If

    Set taskCell = taskCell.Cells(1, 1)

    If Len(Trim$(CStr(taskCell.value))) = 0 _
       Or Not IsNumeric(taskCell.value) Then
        MsgBox "В выбранной форме не указан номер задания.", vbExclamation
        cancelCurrentTaskStable = MACRO_NO_CHANGES
        Exit Function
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "Вы действительно хотите отменить задание №" & _
        taskCell.value & "?", _
        vbQuestion + vbYesNo, _
        "Подтверждение" _
    )

    If answer <> vbYes Then
        cancelCurrentTaskStable = MACRO_NO_CHANGES
        Exit Function
    End If

    Dim comment As String
    comment = InputBox("Введите причину отмены:", "Окно ввода")

    Dim syncResult As String
    syncResult = SetTaskStatusOnly( _
        CLng(taskCell.value), _
        STABLE_TASK_TYPE, _
        "Отмена", _
        comment _
    )

    If InStr(1, syncResult, """ok"": true", vbTextCompare) > 0 Then

        With taskCell
            .Interior.ColorIndex = 3
            If Not .Locked Then .Locked = True
        End With

        cancelCurrentTaskStable = MACRO_CHANGED

    ElseIf syncResult = "NoComment" _
        Or syncResult = "NoTaskId" Then

        cancelCurrentTaskStable = MACRO_NO_CHANGES

    End If

End Function


Public Function Save_Stable( _
    Optional ByVal taskStatus As String = "" _
) As macroResult

    On Error GoTo ErrHandler

    Save_Stable = MACRO_FAILED

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim newHashResult As String
    Dim oldHashResult As String
    Dim lastWorkflowStage As String

    newHashResult = CalculateCurrentStableHash()
    oldHashResult = _
        GetWorkbookMetadataProperty(wb, PROP_STABLE_DATA_HASH)
    lastWorkflowStage = _
        GetWorkbookMetadataProperty( _
            wb, _
            PROP_STABLE_LAST_WORKFLOW_STAGE _
        )

    Dim needToSave As shouldSave
    needToSave = ShouldSaveStudy( _
        newHashResult, _
        oldHashResult, _
        lastWorkflowStage, _
        taskStatus _
    )

    If needToSave = DATA_NO_CHANGES Then
        Save_Stable = MACRO_NO_CHANGES
        Exit Function

    ElseIf needToSave = DATA_EMPTY Then

        SetWorkbookMetadataProperty _
            wb, _
            PROP_STABLE_DATA_HASH, _
            newHashResult, _
            False

        SetWorkbookMetadataProperty _
            wb, _
            PROP_STABLE_LAST_WORKFLOW_STAGE, _
            EffectiveWorkflowStage(taskStatus), _
            False

        SetStudySyncStatus _
            wb, _
            PROP_STABLE_DATA_STATE, _
            SYNC_STATUS_EMPTY, _
            False

        Save_Stable = MACRO_NO_CHANGES
        Exit Function

    ElseIf needToSave = DATA_DELETED Then

        SetStudySyncStatus _
            wb, _
            PROP_STABLE_DATA_STATE, _
            SYNC_STATUS_ERROR, _
            False

        Exit Function

    End If

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim resultTable As ListObject
    Set resultTable = _
        wb.Worksheets(STABLE_RESULTS_SHEET_NAME) _
          .ListObjects(STABLE_RESULTS_TABLE_NAME)

    Dim resultRows As Collection
    Set resultRows = New Collection

    Dim sampleCodes As Object
    Set sampleCodes = CreateObject("Scripting.Dictionary")
    sampleCodes.CompareMode = vbTextCompare

    Dim newResultIDs As Collection
    Set newResultIDs = New Collection

    Dim requireComplete As Boolean
    requireComplete = _
        (LCase$(Trim$(taskStatus)) = LCase$("Валидировано"))

    Dim formIndex As Long
    Dim formState As StableFormState
    Dim lastFormIndex As Long

    lastFormIndex = Stable_LastFormIndexOnSheet(sheetStable)

    For formIndex = 1 To lastFormIndex

        formState = Get_Stable_FormState(formIndex)

        Select Case formState
            Case STABLE_FORM_EMPTY, STABLE_FORM_RESERVED
                GoTo NextForm
            Case STABLE_FORM_FILLED
                ' Продолжаем сохранение.
        End Select

        Dim taskRange As Range
        Set taskRange = Stable_FormRange( _
            sheetStable, _
            formIndex, _
            COL_TASK_ID_DB _
        )

        If Not taskRange Is Nothing Then
            If taskRange.Interior.ColorIndex = 3 Then
                GoTo NextForm
            End If
        End If

        If Not Validate_Stable_Form(formIndex, requireComplete) Then
            Exit Function
        End If

        Dim sampleCode As String
        Dim sampleRange As Range

        Set sampleRange = Stable_FormRange( _
            sheetStable, _
            formIndex, _
            COL_SAMPLE_CODE_DB _
        )

        sampleCode = Trim$(CStr(sampleRange.value))

        If sampleCodes.Exists(sampleCode) Then

            MsgBox _
                "Stable_" & formIndex & " не сохранена." & _
                vbCrLf & vbCrLf & _
                "Шифр пробы " & sampleCode & _
                " уже указан на странице " & _
                sampleCodes(sampleCode) & ".", _
                vbExclamation

            Exit Function

        End If

        sampleCodes.Add sampleCode, formIndex

        Dim resultIdValue As Variant
        Dim parentKey As String
        Dim resultIdRange As Range

        Set resultIdRange = Stable_FormRange( _
            sheetStable, _
            formIndex, _
            COL_RESULT_ID_DB _
        )

        resultIdValue = resultIdRange.value

        If Len(NormalizeKeyValue(resultIdValue)) = 0 _
           Or NormalizeKeyValue(resultIdValue) = "0" Then
            parentKey = "temp_" & formIndex
        Else
            parentKey = NormalizeKeyValue(resultIdValue)
        End If

        If IsNumeric(parentKey) Then
            newResultIDs.Add parentKey
        End If

        Dim resultRow() As Variant
        ReDim resultRow(1 To resultTable.ListColumns.count)

        Dim headerCell As Range
        Dim headerName As String
        Dim columnIndex As Long
        Dim fieldValue As Variant
        Dim fieldRange As Range

        For Each headerCell In resultTable.HeaderRowRange.Cells

            headerName = Trim$(CStr(headerCell.value))
            columnIndex = resultTable.ListColumns(headerName).index

            Select Case LCase$(headerName)

                Case LCase$(COL_DATE_TIME_SYNC_DB)
                    resultRow(columnIndex) = _
                        Format$(Now, "yyyy-mm-dd hh:nn:ss")

                Case LCase$(COL_RESULT_ID_DB)
                    resultRow(columnIndex) = parentKey

                Case Else
                    Set fieldRange = Stable_FormRange( _
                        sheetStable, _
                        formIndex, _
                        headerName _
                    )

                    fieldValue = fieldRange.value

                    resultRow(columnIndex) = _
                        FormatValueForSyncText( _
                            headerName, _
                            fieldValue _
                        )

            End Select

        Next headerCell

        resultRows.Add resultRow

NextForm:
    Next formIndex

    Dim cellID As Range
    Dim oldResultID As String

    If Not resultTable.DataBodyRange Is Nothing Then

        For Each cellID In _
            resultTable.ListColumns(COL_RESULT_ID_DB) _
                       .DataBodyRange.Cells

            If Not IsError(cellID.value) Then

                oldResultID = NormalizeKeyValue(cellID.value)

                If IsNumeric(oldResultID) _
                   And oldResultID <> "0" Then

                    If Not CollectionContains( _
                        newResultIDs, _
                        oldResultID _
                    ) Then

                        MsgBox _
                            "Удалены данные синхронизированной " & _
                            "формы Stable. Синхронизация " & _
                            "невозможна. Вызовите загрузку " & _
                            "данных из БД.", _
                            vbExclamation

                        Exit Function

                    End If

                End If

            End If

        Next cellID

    End If

    WriteCollectionToListObject resultTable, resultRows

    Dim syncResult As String
    syncResult = RunStudySync( _
        "save", _
        studies:=STABLE_STUDY_CODE, _
        taskStatus:=taskStatus _
    )

    If InStr(1, syncResult, """ok"": true", vbTextCompare) > 0 Then

        Load_Stable_results_to_forms _
            "by_key", _
            COL_RESULT_ID_DB

        newHashResult = CalculateCurrentStableHash()

        SetWorkbookMetadataProperty _
            wb, _
            PROP_STABLE_DATA_HASH, _
            newHashResult, _
            False

        SetWorkbookMetadataProperty _
            wb, _
            PROP_STABLE_LAST_WORKFLOW_STAGE, _
            EffectiveWorkflowStage(taskStatus), _
            False

        SetStudySyncStatus _
            wb, _
            PROP_STABLE_DATA_STATE, _
            SYNC_STATUS_SYNCED, _
            False

        ClearStudyHashCheck wb, STABLE_STUDY_CODE

    Else

        MsgBox _
            "Python-синхронизация Stable не подтвердила " & _
            "успешное сохранение." & vbCrLf & vbCrLf & _
            syncResult, _
            vbExclamation

        Exit Function

    End If

    Save_Stable = MACRO_CHANGED
    Exit Function

ErrHandler:
    Debug.Print "Save_Stable: " & Err.Number & ": " & Err.Description
    Save_Stable = MACRO_FAILED

End Function


Public Function CalculateCurrentStableHash() As String

    Dim Data As Object
    Set Data = Build_Stable_HashData(ActiveWorkbook)

    CalculateCurrentStableHash = CalculateDataHash(Data)

End Function


Private Function Build_Stable_HashData( _
    ByVal wb As Workbook _
) As Object

    Dim Data As Object
    Set Data = CreateObject("Scripting.Dictionary")

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim resultTable As ListObject
    Set resultTable = _
        wb.Worksheets(STABLE_RESULTS_SHEET_NAME) _
          .ListObjects(STABLE_RESULTS_TABLE_NAME)

    Dim formIndex As Long
    Dim resultIndex As Long
    Dim lastFormIndex As Long

    lastFormIndex = Stable_LastFormIndexOnSheet(sheetStable)

    For formIndex = 1 To lastFormIndex

        If Get_Stable_FormState(formIndex) <> _
           STABLE_FORM_FILLED Then
            GoTo NextForm
        End If

        Dim taskRange As Range
        Set taskRange = Stable_FormRange( _
            sheetStable, _
            formIndex, _
            COL_TASK_ID_DB _
        )

        If Not taskRange Is Nothing Then
            If taskRange.Interior.ColorIndex = 3 Then
                GoTo NextForm
            End If
        End If

        resultIndex = resultIndex + 1

        Dim headerCell As Range
        Dim headerName As String
        Dim fieldValue As Variant
        Dim formatValue As Variant
        Dim fieldRange As Range

        For Each headerCell In resultTable.HeaderRowRange.Cells

            headerName = Trim$(CStr(headerCell.value))

            Select Case LCase$(headerName)

                Case LCase$(COL_DATE_TIME_SYNC_DB), _
                     LCase$(COL_RESULT_ID_DB)

                Case Else
                    Set fieldRange = Stable_FormRange( _
                        sheetStable, _
                        formIndex, _
                        headerName _
                    )

                    If fieldRange Is Nothing Then
                        Err.Raise _
                            vbObjectError + 4103, _
                            "Build_Stable_HashData", _
                            "Не найден базовый диапазон " & _
                            StableRangeLabel(formIndex, headerName)
                    End If

                    fieldValue = fieldRange.value

                    formatValue = FormatValueForSyncText( _
                        headerName, _
                        fieldValue _
                    )

                    Data.Add _
                        resultIndex & "." & headerName, _
                        formatValue

            End Select

        Next headerCell

NextForm:
    Next formIndex

    Set Build_Stable_HashData = Data

End Function


Private Function CheckStableTask( _
    ByVal taskId As Long, _
    ByVal sampleCode As String, _
    ByVal pageNumber As Long _
) As String

    Dim tasks As ListObject
    Set tasks = ActiveWorkbook.Worksheets("Task").ListObjects(1)

    If tasks.DataBodyRange Is Nothing Then
        CheckStableTask = _
            "Таблица заданий пуста. Обновите данные проекта."
        Exit Function
    End If

    Dim taskCheck As Range
    Set taskCheck = _
        tasks.ListColumns("Номер задания") _
             .DataBodyRange.Find( _
                 What:=taskId, _
                 LookAt:=xlWhole, _
                 SearchOrder:=xlByColumns _
             )

    If taskCheck Is Nothing Then

        CheckStableTask = _
            "Указан неизвестный номер задания " & taskId & _
            " на странице " & pageNumber & _
            ". Попробуйте обновить данные проекта " & _
            "или исправьте значение."

        Exit Function

    End If

    Dim taskRowIndex As Long
    taskRowIndex = _
        taskCheck.row - tasks.DataBodyRange.row + 1

    If CStr(tasks.ListColumns("Код проекта") _
                .DataBodyRange.Cells(taskRowIndex, 1).value) _
       <> sampleCode Then

        CheckStableTask = _
            "Указан неверный номер задания " & taskId & _
            " на странице " & pageNumber & _
            ". Шифр пробы в задании отличается " & _
            "от шифра пробы в исследовании."

        Exit Function

    End If

    CheckStableTask = "OK"

End Function


Public Function Load_Stable_results_to_forms( _
    Optional ByVal mode As String = "by_key", _
    Optional ByVal fields As Variant, _
    Optional ByVal silence As Boolean = False _
) As String

    On Error GoTo CleanFail

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim resultTable As ListObject
    Set resultTable = _
        wb.Worksheets(STABLE_RESULTS_SHEET_NAME) _
          .ListObjects(STABLE_RESULTS_TABLE_NAME)

    Dim fieldFilter As Object
    Set fieldFilter = BuildFieldFilter(fields)

    Dim text As String
    Dim warnText As String

    mode = LCase$(Trim$(mode))

    If mode <> "by_key" And mode <> "sequential" Then

        text = _
            "ERROR|mode должен быть 'by_key' или 'sequential'"
        Load_Stable_results_to_forms = text

        If Not silence Then MsgBox text, vbExclamation
        Exit Function

    End If

    If resultTable.DataBodyRange Is Nothing _
       Or Application.WorksheetFunction.CountA( _
            resultTable.DataBodyRange _
       ) = 0 Then

        text = "OK|Таблица Stable_results пустая."
        Load_Stable_results_to_forms = text
        Exit Function

    End If

    Dim resultRowIndex As Long
    Dim formIndex As Long
    Dim notFoundCount As Long

    For resultRowIndex = 1 To resultTable.ListRows.count

        If mode = "sequential" Then

            formIndex = resultRowIndex
            EnsureStableFormExists wb, formIndex

        Else

            formIndex = FindStableFormIndexForResultRow( _
                sheetStable, _
                resultTable, _
                resultRowIndex _
            )

        End If

        If formIndex = 0 _
           Or formIndex > _
                Stable_LastFormIndexOnSheet(sheetStable) Then

            notFoundCount = notFoundCount + 1

            warnText = warnText & _
                "Не найдена форма для строки " & _
                "Stable_results #" & resultRowIndex & vbCrLf & _
                "sampleCode = " & _
                GetTableValue( _
                    resultTable, _
                    resultRowIndex, _
                    COL_SAMPLE_CODE_DB _
                ) & vbCrLf & _
                "TaskId = " & _
                GetTableValue( _
                    resultTable, _
                    resultRowIndex, _
                    COL_TASK_ID_DB _
                ) & vbCrLf & vbCrLf

            GoTo NextResultRow

        End If

        CopyStableTableRowToForm _
            sheetStable, _
            resultTable, _
            resultRowIndex, _
            formIndex, _
            fieldFilter

NextResultRow:
    Next resultRowIndex

    If notFoundCount = 0 Then

        text = "OK|Данные Stable перенесены в формы."
        Load_Stable_results_to_forms = text

    Else

        text = _
            "WARN|Данные Stable перенесены частично. " & _
            "Не найдено форм: " & notFoundCount & _
            vbCrLf & warnText

        Load_Stable_results_to_forms = text

        If Not silence Then MsgBox text, vbExclamation

    End If

    Exit Function

CleanFail:
    text = _
        "ERROR|" & CStr(Err.Number) & "|" & Err.Description

    Load_Stable_results_to_forms = text

    If Not silence Then MsgBox text, vbCritical

End Function


Private Function FindStableFormIndexForResultRow( _
    ByVal ws As Worksheet, _
    ByVal resultTable As ListObject, _
    ByVal resultRowIndex As Long _
) As Long

    Dim resultValue As String
    resultValue = NormalizeKeyValue( _
        GetTableValue( _
            resultTable, _
            resultRowIndex, _
            COL_SAMPLE_CODE_DB _
        ) _
    )

    If Len(resultValue) = 0 Then Exit Function

    Dim formIndex As Long
    Dim lastFormIndex As Long
    Dim sampleRange As Range

    lastFormIndex = Stable_LastFormIndexOnSheet(ws)

    For formIndex = 1 To lastFormIndex

        Set sampleRange = Stable_FormRange( _
            ws, _
            formIndex, _
            COL_SAMPLE_CODE_DB _
        )

        If Not sampleRange Is Nothing Then
            If NormalizeKeyValue(sampleRange.value) = resultValue Then
                FindStableFormIndexForResultRow = formIndex
                Exit Function
            End If
        End If

    Next formIndex

End Function


Private Sub CopyStableTableRowToForm( _
    ByVal ws As Worksheet, _
    ByVal sourceTable As ListObject, _
    ByVal sourceRowIndex As Long, _
    ByVal formIndex As Long, _
    Optional ByVal fieldFilter As Object _
)

    If sourceTable.DataBodyRange Is Nothing Then Exit Sub

    Dim tableColumn As ListColumn
    Dim fieldName As String
    Dim targetRange As Range

    For Each tableColumn In sourceTable.ListColumns

        fieldName = CStr(tableColumn.name)

        If ShouldTransferField(fieldFilter, fieldName) Then

            Set targetRange = Stable_FormRange( _
                ws, _
                formIndex, _
                fieldName _
            )

            If Not targetRange Is Nothing Then
                If Not RangeHasFormula(targetRange) Then
                    targetRange.value = _
                        sourceTable.DataBodyRange.Cells( _
                            sourceRowIndex, _
                            tableColumn.index _
                        ).value
                End If
            End If

        End If

    Next tableColumn

End Sub


Private Function Validate_Stable_Form( _
    ByVal formIndex As Long, _
    Optional ByVal requireComplete As Boolean = False _
) As Boolean

    Validate_Stable_Form = False

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim resultTable As ListObject
    Set resultTable = _
        wb.Worksheets(STABLE_RESULTS_SHEET_NAME) _
          .ListObjects(STABLE_RESULTS_TABLE_NAME)

    Dim errors As String
    Dim errorCount As Long

    If formIndex <= 0 Then
        MsgBox _
            "Некорректный номер формы Stable: " & formIndex, _
            vbExclamation
        Exit Function
    End If

    If formIndex > Stable_LastFormIndexOnSheet(sheetStable) Then
        MsgBox _
            "Форма Stable_" & formIndex & " не найдена.", _
            vbExclamation
        Exit Function
    End If

    Dim headerCell As Range
    Dim rng As Range
    Dim fieldName As String
    Dim fieldNameLower As String
    Dim fieldCaption As String
    Dim rangeName As String
    Dim fieldValue As Variant
    Dim valueText As String
    Dim normalizedExecutor As String
    Dim sampleCodeRange As Range

    For Each headerCell In resultTable.HeaderRowRange.Cells

        fieldName = Trim$(CStr(headerCell.value))
        fieldNameLower = LCase$(fieldName)

        If fieldNameLower = LCase$(COL_DATE_TIME_SYNC_DB) Then
            GoTo NextHeader
        End If

        rangeName = StableRangeLabel(formIndex, fieldName)

        Set rng = Stable_FormRange( _
            sheetStable, _
            formIndex, _
            fieldName _
        )

        If rng Is Nothing Then

            AddValidationError _
                errors, _
                errorCount, _
                "Не найден базовый именованный диапазон: " & _
                STABLE_RANGE_PREFIX & "1_" & fieldName & _
                " (поле " & rangeName & ")"

            GoTo NextHeader

        End If

        Set rng = rng.Cells(1, 1)

        ClearValidationFill rng

        fieldValue = rng.value
        fieldCaption = Stable_FieldCaption(fieldName)

        If IsError(fieldValue) Then

            AddValidationError _
                errors, _
                errorCount, _
                rangeName & " содержит ошибку Excel: " & _
                rng.text & " (" & rng.Worksheet.name & _
                "!" & rng.Address(False, False) & ")", _
                rng

            GoTo NextHeader

        End If

        valueText = Trim$(CStr(fieldValue))

        ' Шифр обязателен для любой сохраняемой формы.
        If fieldNameLower = "samplecode" _
           And valueText = "" Then

            AddValidationError _
                errors, _
                errorCount, _
                "Не заполнено обязательное поле """ & _
                fieldCaption & """: " & rangeName, _
                rng

        End If

        ' При валидации требуется полностью заполненный результат.
        If requireComplete Then

            Select Case fieldNameLower

                Case "datestart", _
                     "timestart", _
                     "dateend", _
                     "timeend", _
                     "operator", _
                     "pvtcell", _
                     "temperature", _
                     "pressure", _
                     "pressureunit", _
                     "pressureabs", _
                     "pressurempaabs", _
                     "liquidvolume"

                    If valueText = "" Then

                        AddValidationError _
                            errors, _
                            errorCount, _
                            "Не заполнено обязательное поле """ & _
                            fieldCaption & """: " & rangeName, _
                            rng

                    End If

            End Select

        End If

        Select Case fieldNameLower

            Case "taskid", _
                 "temperature", _
                 "pressure", _
                 "pressurempaabs", _
                 "liquidvolume"

                If valueText <> "" _
                   And Not IsNumeric(fieldValue) Then

                    AddValidationError _
                        errors, _
                        errorCount, _
                        "Поле """ & fieldCaption & _
                        """ должно содержать число: " & _
                        rangeName, _
                        rng

                End If

        End Select

        Select Case fieldNameLower

            Case "taskid", _
                 "pressure", _
                 "pressurempaabs", _
                 "liquidvolume"

                If valueText <> "" _
                   And IsNumeric(fieldValue) Then

                    If CDbl(fieldValue) < 0 Then

                        AddValidationError _
                            errors, _
                            errorCount, _
                            "Поле """ & fieldCaption & _
                            """ не может быть отрицательным: " & _
                            rangeName, _
                            rng

                    End If

                End If

        End Select

        If fieldNameLower = "taskid" _
           And valueText <> "" _
           And IsNumeric(fieldValue) Then

            If CDbl(fieldValue) <> Fix(CDbl(fieldValue)) Then

                AddValidationError _
                    errors, _
                    errorCount, _
                    "Поле """ & fieldCaption & _
                    """ должно содержать целое число: " & _
                    rangeName, _
                    rng

            ElseIf CDbl(fieldValue) > 0 Then

                Dim taskCheckResult As String

                Set sampleCodeRange = Stable_FormRange( _
                    sheetStable, _
                    formIndex, _
                    COL_SAMPLE_CODE_DB _
                )

                taskCheckResult = CheckStableTask( _
                    CLng(fieldValue), _
                    CStr(sampleCodeRange.value), _
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

        End If

        Select Case fieldNameLower

            Case "datestart", "dateend"

                If valueText <> "" _
                   And Not IsDate(fieldValue) Then

                    AddValidationError _
                        errors, _
                        errorCount, _
                        "Значение поля """ & fieldCaption & _
                        """ должно быть датой: " & rangeName, _
                        rng

                End If

            Case "timestart", "timeend"

                If valueText <> "" Then

                    If Not IsDate(fieldValue) _
                       And Not IsNumeric(fieldValue) Then

                        AddValidationError _
                            errors, _
                            errorCount, _
                            "Значение поля """ & fieldCaption & _
                            """ должно быть временем: " & _
                            rangeName, _
                            rng

                    End If

                End If

        End Select

        If fieldNameLower = "operator" _
           And valueText <> "" Then

            If Not IsValidExecutor(fieldValue) Then

                AddValidationError _
                    errors, _
                    errorCount, _
                    "Неверно указан исполнитель: " & rangeName, _
                    rng

            Else

                normalizedExecutor = Replace(valueText, "/", ",")
                normalizedExecutor = _
                    Replace(normalizedExecutor, "\", ",")
                rng.value = normalizedExecutor

            End If

        End If

        If fieldNameLower = "pressureabs" _
           And valueText <> "" Then

            Select Case LCase$(valueText)
                Case "true", "false", "истина", "ложь", _
                     "1", "0", "yes", "no", "да", "нет", _
                     "+", "-"
                    ' Допустимое логическое значение.
                Case Else
                    AddValidationError _
                        errors, _
                        errorCount, _
                        "Поле """ & fieldCaption & _
                        """ должно содержать True/False: " & _
                        rangeName, _
                        rng
            End Select

        End If

        If fieldNameLower = "temperature" _
           And valueText <> "" _
           And IsNumeric(fieldValue) Then

            If CDbl(fieldValue) < -20 _
               Or CDbl(fieldValue) > 140 Then

                AddValidationError _
                    errors, _
                    errorCount, _
                    "Температура должна находиться в диапазоне " & _
                    "от -20 до 140 °C: " & rangeName, _
                    rng

            End If

        End If

NextHeader:
    Next headerCell

    ValidateStableDateTimePairs _
        wb, _
        formIndex, _
        errors, _
        errorCount

    If errorCount > 0 Then

        MsgBox _
            "Stable_" & formIndex & _
            " не сохранена. Исправьте ошибки:" & _
            vbCrLf & vbCrLf & errors, _
            vbExclamation

        Exit Function

    End If

    Validate_Stable_Form = True

End Function


Private Sub ValidateStableDateTimePairs( _
    ByVal wb As Workbook, _
    ByVal formIndex As Long, _
    ByRef errors As String, _
    ByRef errorCount As Long _
)

    Dim startDateRange As Range
    Dim startTimeRange As Range
    Dim endDateRange As Range
    Dim endTimeRange As Range
    Dim ws As Worksheet

    Set ws = wb.Worksheets(STABLE_SHEET_NAME)

    If Not Stable_TryGetFormRange( _
        ws, _
        formIndex, _
        "dateStart", _
        startDateRange _
    ) Then Exit Sub

    If Not Stable_TryGetFormRange( _
        ws, _
        formIndex, _
        "timeStart", _
        startTimeRange _
    ) Then Exit Sub

    If Not Stable_TryGetFormRange( _
        ws, _
        formIndex, _
        "dateEnd", _
        endDateRange _
    ) Then Exit Sub

    If Not Stable_TryGetFormRange( _
        ws, _
        formIndex, _
        "timeEnd", _
        endTimeRange _
    ) Then Exit Sub

    Set startDateRange = startDateRange.Cells(1, 1)
    Set startTimeRange = startTimeRange.Cells(1, 1)
    Set endDateRange = endDateRange.Cells(1, 1)
    Set endTimeRange = endTimeRange.Cells(1, 1)

    Dim startDateText As String
    Dim startTimeText As String
    Dim endDateText As String
    Dim endTimeText As String
    Dim validationTarget As Range

    startDateText = Trim$(CStr(startDateRange.value))
    startTimeText = Trim$(CStr(startTimeRange.value))
    endDateText = Trim$(CStr(endDateRange.value))
    endTimeText = Trim$(CStr(endTimeRange.value))

    If (startDateText = "") Xor (startTimeText = "") Then

        If startDateText = "" Then
            Set validationTarget = startDateRange
        Else
            Set validationTarget = startTimeRange
        End If

        AddValidationError _
            errors, _
            errorCount, _
            "Дата и время начала должны быть заполнены вместе.", _
            validationTarget

    End If

    If (endDateText = "") Xor (endTimeText = "") Then

        If endDateText = "" Then
            Set validationTarget = endDateRange
        Else
            Set validationTarget = endTimeRange
        End If

        AddValidationError _
            errors, _
            errorCount, _
            "Дата и время окончания должны быть заполнены вместе.", _
            validationTarget

    End If

    If startDateText = "" _
       Or startTimeText = "" _
       Or endDateText = "" _
       Or endTimeText = "" Then
        Exit Sub
    End If

    If Not IsDate(startDateRange.value) _
       Or (Not IsDate(startTimeRange.value) _
           And Not IsNumeric(startTimeRange.value)) _
       Or Not IsDate(endDateRange.value) _
       Or (Not IsDate(endTimeRange.value) _
           And Not IsNumeric(endTimeRange.value)) Then
        Exit Sub
    End If

    Dim startDateTime As Date
    Dim endDateTime As Date

    startDateTime = _
        DateValue(CDate(startDateRange.value)) + _
        TimeValue(CDate(startTimeRange.value))

    endDateTime = _
        DateValue(CDate(endDateRange.value)) + _
        TimeValue(CDate(endTimeRange.value))

    If endDateTime < startDateTime Then

        AddValidationError _
            errors, _
            errorCount, _
            "Дата и время окончания не могут быть раньше " & _
            "даты и времени начала.", _
            endDateRange

    End If

End Sub


Private Function Stable_FieldCaption( _
    ByVal fieldName As String _
) As String

    Select Case LCase$(fieldName)

        Case "resultidstable"
            Stable_FieldCaption = "Номер результата"

        Case "taskid"
            Stable_FieldCaption = "Номер задания"

        Case "samplecode"
            Stable_FieldCaption = "Шифр пробы"

        Case "datestart"
            Stable_FieldCaption = "Дата начала"

        Case "timestart"
            Stable_FieldCaption = "Время начала"

        Case "dateend"
            Stable_FieldCaption = "Дата окончания"

        Case "timeend"
            Stable_FieldCaption = "Время окончания"

        Case "operator"
            Stable_FieldCaption = "Исполнитель"

        Case "pvtcell"
            Stable_FieldCaption = "PVT-ячейка"

        Case "temperature"
            Stable_FieldCaption = "Температура"

        Case "pressure"
            Stable_FieldCaption = "Давление"

        Case "pressureunit"
            Stable_FieldCaption = "Единица измерения давления"

        Case "pressureabs"
            Stable_FieldCaption = "Абсолютное давление"

        Case "pressurempaabs"
            Stable_FieldCaption = "Давление, МПа абс."

        Case "liquidvolume"
            Stable_FieldCaption = "Объём жидкости"

        Case "datetimesync"
            Stable_FieldCaption = "Дата синхронизации"

        Case Else
            Stable_FieldCaption = fieldName

    End Select

End Function


Private Function Get_Stable_FormState( _
    ByVal formIndex As Long _
) As StableFormState

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim sampleCode As String
    Dim sampleRange As Range

    Set sampleRange = Stable_FormRange( _
        sheetStable, _
        formIndex, _
        COL_SAMPLE_CODE_DB _
    )

    If sampleRange Is Nothing Then
        Get_Stable_FormState = STABLE_FORM_RESERVED
        Exit Function
    End If

    sampleCode = Trim$(CStr(sampleRange.value))

    Dim hasMeaningfulData As Boolean
    Dim resultTable As ListObject
    Set resultTable = _
        wb.Worksheets(STABLE_RESULTS_SHEET_NAME) _
          .ListObjects(STABLE_RESULTS_TABLE_NAME)

    Dim headerCell As Range
    Dim fieldName As String
    Dim fieldNameLower As String
    Dim rng As Range
    Dim fieldValue As Variant
    Dim valueText As String

    For Each headerCell In resultTable.HeaderRowRange.Cells

        fieldName = Trim$(CStr(headerCell.value))
        fieldNameLower = LCase$(fieldName)

        Select Case fieldNameLower
            Case LCase$(COL_RESULT_ID_DB), _
                 LCase$(COL_TASK_ID_DB), _
                 LCase$(COL_SAMPLE_CODE_DB), _
                 LCase$(COL_DATE_TIME_SYNC_DB), _
                 "pressureabs"
                GoTo NextStateField
        End Select

        Set rng = Nothing

        If Stable_TryGetFormRange( _
            sheetStable, _
            formIndex, _
            fieldName, _
            rng _
        ) Then

            If Not RangeHasFormula(rng) Then

                fieldValue = rng.Cells(1, 1).value

                If IsError(fieldValue) Then
                    hasMeaningfulData = True
                    Exit For
                End If

                valueText = Trim$(CStr(fieldValue))

                If valueText <> "" Then

                    If VarType(fieldValue) = vbBoolean Then
                        If CBool(fieldValue) Then
                            hasMeaningfulData = True
                            Exit For
                        End If
                    ElseIf IsNumeric(fieldValue) Then
                        If CDbl(fieldValue) <> 0 Then
                            hasMeaningfulData = True
                            Exit For
                        End If
                    Else
                        hasMeaningfulData = True
                        Exit For
                    End If

                End If

            End If

        End If

NextStateField:
    Next headerCell

    If sampleCode = "" And Not hasMeaningfulData Then
        Get_Stable_FormState = STABLE_FORM_EMPTY
        Exit Function
    End If

    If sampleCode <> "" And hasMeaningfulData Then
        Get_Stable_FormState = STABLE_FORM_FILLED
        Exit Function
    End If

    Get_Stable_FormState = STABLE_FORM_RESERVED

End Function


Private Function StableRangeLabel( _
    ByVal formIndex As Long, _
    ByVal fieldName As String _
) As String

    StableRangeLabel = _
        STABLE_RANGE_PREFIX & formIndex & "_" & fieldName

End Function


Private Function getLastFormNumberStable() As Long

    getLastFormNumberStable = Stable_GetLastFormIndex()

End Function


Public Sub AddFormStable()

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim formLength As Long
    formLength = STABLE_FORM_COLUMN_COUNT

    Dim newFormNumber As Long
    newFormNumber = getLastFormNumberStable() + 1

    Dim lastColumn As Long
    Dim firstColumn As Long

    lastColumn = _
        newFormNumber * formLength + _
        STABLE_FIRST_FORM_FIRST_COL - 1

    firstColumn = lastColumn - formLength + 1

    sheetStable.Activate

    Dim oldCopyObjectsWithCells As Boolean
    oldCopyObjectsWithCells = Application.CopyObjectsWithCells

    Application.CopyObjectsWithCells = False

    sheetStable.Range( _
        sheetStable.Columns(firstColumn - formLength), _
        sheetStable.Columns(lastColumn - formLength) _
    ).EntireColumn.Copy

    sheetStable.Columns(firstColumn).PasteSpecial Paste:=xlPasteAll
    sheetStable.Columns(firstColumn).PasteSpecial _
        Paste:=xlPasteColumnWidths

    Application.CutCopyMode = False
    Application.CopyObjectsWithCells = oldCopyObjectsWithCells

    Dim pageCell As Range
    Set pageCell = _
        Stable_PageCell(sheetStable, newFormNumber)

    If pageCell Is Nothing Then
        Err.Raise _
            vbObjectError + 4104, _
            "AddFormStable", _
            "Не удалось определить ячейку номера страницы Stable."
    End If

    pageCell.value = newFormNumber

    ResetStableCopiedForm _
        wb, _
        sheetStable, _
        newFormNumber - 1, _
        newFormNumber

    CreateStableCheckBoxes newFormNumber

End Sub


Public Sub DeleteLastListStable()

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim lastForm As Long
    lastForm = getLastFormNumberStable()

    If lastForm <= STABLE_MIN_FORM_COUNT Then
        MsgBox _
            "Нельзя удалить первые " & _
            STABLE_MIN_FORM_COUNT & " формы.", _
            vbExclamation
        Exit Sub
    End If

    Dim formLength As Long
    formLength = STABLE_FORM_COLUMN_COUNT

    Dim firstColLastForm As Long
    firstColLastForm = _
        STABLE_FIRST_FORM_FIRST_COL + _
        (lastForm - 1) * formLength

    Dim lastColLastForm As Long
    lastColLastForm = firstColLastForm + formLength - 1

    sheetStable.Range( _
        sheetStable.Columns(firstColLastForm), _
        sheetStable.Columns(lastColLastForm) _
    ).EntireColumn.Delete

    DeleteCheckBoxesFromColumnRight _
        sheetStable, _
        firstColLastForm

End Sub


Public Sub EnsureStableFormExists( _
    ByVal wb As Workbook, _
    ByVal formIndex As Long _
)

    Dim previousLastForm As Long

    Do While formIndex > Stable_GetLastFormIndex()
        previousLastForm = Stable_GetLastFormIndex()
        AddFormStable

        If Stable_GetLastFormIndex() <= previousLastForm Then
            Err.Raise _
                vbObjectError + 4105, _
                "EnsureStableFormExists", _
                "Не удалось добавить форму Stable."
        End If
    Loop

End Sub


Private Sub CreateStableCheckBoxes( _
    ByVal formIndex As Long _
)

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim rangeName As String
    rangeName = StableRangeLabel(formIndex, "pressureAbs")

    Dim linkedCell As Range
    Set linkedCell = Stable_FormRange( _
        sheetStable, _
        formIndex, _
        "pressureAbs" _
    )

    If linkedCell Is Nothing Then
        Err.Raise _
            vbObjectError + 4102, _
            "CreateStableCheckBoxes", _
            "Не удалось вычислить диапазон " & rangeName & _
            ". Проверьте базовое имя Stable_1_pressureAbs."
    End If

    Dim area As Range
    If linkedCell.MergeCells Then
        Set area = linkedCell.MergeArea
    Else
        Set area = linkedCell
    End If

    Dim cb As CheckBox
    Set cb = sheetStable.CheckBoxes.Add( _
        area.Left, _
        area.Top, _
        area.Width, _
        area.Height _
    )

    With cb
        .name = "CheckBox_Stable_" & _
                formIndex & "_pressureAbs"
        .Caption = vbNullString
        .linkedCell = _
            "'" & sheetStable.name & "'!" & _
            linkedCell.Address(False, False)
        .value = xlOff
        .Placement = xlMove
    End With

    linkedCell.value = False

End Sub


Private Sub ResetStableCopiedForm( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal oldFormNumber As Long, _
    ByVal newFormNumber As Long _
)

    Dim resultTable As ListObject
    Set resultTable = _
        wb.Worksheets(STABLE_RESULTS_SHEET_NAME) _
          .ListObjects(STABLE_RESULTS_TABLE_NAME)

    Dim headerCell As Range
    Dim fieldName As String
    Dim newRange As Range

    For Each headerCell In resultTable.HeaderRowRange.Cells

        fieldName = Trim$(CStr(headerCell.value))

        If fieldName = vbNullString Then GoTo ContinueLoop

        If LCase$(fieldName) = _
           LCase$(COL_DATE_TIME_SYNC_DB) Then
            GoTo ContinueLoop
        End If

        Set newRange = Stable_FormRange( _
            ws, _
            newFormNumber, _
            fieldName _
        )

        If newRange Is Nothing Then
            Err.Raise _
                vbObjectError + 4106, _
                "ResetStableCopiedForm", _
                "Не найден базовый диапазон Stable_1_" & _
                fieldName
        End If

        ResetCopiedCell _
            newRange, _
            oldFormNumber, _
            newFormNumber

ContinueLoop:
    Next headerCell

End Sub


Public Function FindStableFormByTaskId( _
    ByVal taskId As Variant _
) As Long

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(STABLE_SHEET_NAME)

    Dim searchValue As String
    searchValue = NormalizeKeyValue(taskId)

    If Len(searchValue) = 0 Then Exit Function

    Dim formIndex As Long
    Dim lastFormIndex As Long
    Dim taskRange As Range

    lastFormIndex = Stable_LastFormIndexOnSheet(ws)

    For formIndex = 1 To lastFormIndex

        Set taskRange = Stable_FormRange( _
            ws, _
            formIndex, _
            COL_TASK_ID_DB _
        )

        If Not taskRange Is Nothing Then
            If NormalizeKeyValue(taskRange.value) = searchValue Then
                FindStableFormByTaskId = formIndex
                Exit Function
            End If
        End If

    Next formIndex

End Function


Public Function FindEmptyFormStable() As Long

    Dim formIndex As Long
    Dim lastFormIndex As Long

    lastFormIndex = Stable_GetLastFormIndex()

    For formIndex = 1 To lastFormIndex

        If Get_Stable_FormState(formIndex) = _
           STABLE_FORM_EMPTY Then

            FindEmptyFormStable = formIndex
            Exit Function

        End If

    Next formIndex

    AddFormStable
    FindEmptyFormStable = Stable_GetLastFormIndex()

End Function


Public Sub equipmentModelsRefreshStable()

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetStable As Worksheet
    Set sheetStable = wb.Worksheets(STABLE_SHEET_NAME)

    Dim equipmentModels() As String
    equipmentModels = getEquipmentModels("", "PVTCell")

    Dim validationList As String
    Dim i As Long

    validationList = equipmentModels(1)

    For i = 2 To UBound(equipmentModels)
        validationList = _
            validationList & "," & equipmentModels(i)
    Next i

    Dim targetCell As Range
    Dim formIndex As Long
    Dim lastFormIndex As Long

    lastFormIndex = Stable_LastFormIndexOnSheet(sheetStable)

    On Error Resume Next

    For formIndex = 1 To lastFormIndex

        Set targetCell = Nothing
        Set targetCell = Stable_FormRange( _
            sheetStable, _
            formIndex, _
            "PVTcell" _
        )

        If targetCell Is Nothing Then Exit For

        With targetCell.Validation
            .Delete
            .Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:=validationList
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With

    Next formIndex

    On Error GoTo 0

End Sub
