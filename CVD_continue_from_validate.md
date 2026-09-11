# CVD: продолжение изменений от `Validate_CVD_Form`

Проверен файл `Pasted text(20260911-202723).txt` от 11.09.2026.

До `Validate_CVD_Form` переход на диапазоны первой формы уже в основном выполнен. Повторно вставлять константы, `CVD_FormRange`, добавление/удаление формы, сохранение, хеширование и загрузку не нужно.

В текущем `CVD.bas` остались старые обращения к именам форм 2+ только в следующих местах:

1. `Validate_CVD_Form` — прямые ссылки на `sampleCode`, `Ttransfer` и все дочерние поля.
2. `CVD_StepHasData` — вся функция ещё использует `NamedRangeExists`.
3. `FindEmptyFormCVD` — форма определяется по существованию `CVD_N_sampleCode`.
4. В `wrappers.bas` из ранее присланной версии могут оставаться вызовы несуществующих `InsertListCVD2` и `DeleteLastListCVD2`.

Ниже дан порядок изменений именно для текущего состояния файла.

## Шаг 1. Полностью заменить `Validate_CVD_Form`

В `CVD.bas` найти:

```vb
Private Function Validate_CVD_Form(ByVal formIndex As Long) As Boolean
```

Удалить всю функцию целиком до соответствующего `End Function`, расположенного непосредственно перед:

```vb
Private Function CVD_FieldCaption( _
```

Вставить вместо неё:

```vb
Private Function Validate_CVD_Form( _
    ByVal formIndex As Long _
) As Boolean

    Validate_CVD_Form = False

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim sheetCVD As Worksheet
    Set sheetCVD = wb.Worksheets(CVD_FORM_SHEET)

    Dim CVDTable As ListObject
    Set CVDTable = _
        wb.Worksheets(CVD_RESULTS_SHEET) _
          .ListObjects(CVD_RESULTS_TABLE)

    Dim CVDSourceDataTable As ListObject
    Set CVDSourceDataTable = _
        wb.Worksheets(CVD_SOURCE_SHEET) _
          .ListObjects(CVD_SOURCE_TABLE)

    Dim errors As String
    Dim errorCount As Long

    If formIndex <= 0 Then
        MsgBox _
            "Некорректный номер формы CVD: " & formIndex, _
            vbExclamation
        Exit Function
    End If

    If Not CVD_FormExists(formIndex) Then
        MsgBox _
            "Форма CVD_" & formIndex & " не найдена.", _
            vbExclamation
        Exit Function
    End If

    Dim headerCell As Range
    Dim fieldName As String
    Dim rangeName As String
    Dim rng As Range
    Dim valueText As String

    Dim taskCheckResult As String
    Dim sampleCodeRange As Range
    Dim transferTemperatureRange As Range

    ' ============================================================
    ' 1. ПОЛЯ CVD_RESULTS
    ' ============================================================

    For Each headerCell In CVDTable.HeaderRowRange.Cells

        fieldName = CStr(headerCell.Value)

        If StrComp( _
            fieldName, _
            "dateTimeSync", _
            vbTextCompare _
        ) = 0 Then
            GoTo NextResultHeader
        End If

        ' Это только подпись для сообщений.
        rangeName = CVD_RangeLabel(formIndex, fieldName)

        Set rng = Nothing
        Set rng = CVD_FormRange( _
            sheetCVD, _
            formIndex, _
            fieldName _
        )

        If rng Is Nothing Then
            errorCount = errorCount + 1
            errors = errors & errorCount & _
                     ". Не найден базовый диапазон CVD_1_" & _
                     fieldName & "." & vbCrLf
            GoTo NextResultHeader
        End If

        If IsError(rng.Value) Then
            errorCount = errorCount + 1
            errors = errors & errorCount & ". " & _
                     rangeName & _
                     " содержит ошибку Excel: " & _
                     rng.Text & " (" & _
                     rng.Worksheet.name & "!" & _
                     rng.Address(False, False) & ")" & _
                     vbCrLf
            GoTo NextResultHeader
        End If

        valueText = Trim$(CStr(rng.Value))

        Select Case fieldName

            Case COL_CVD_RESULT_ID_DB

                ' Может быть пустым до первого сохранения.

            Case "sampleCode"

                If Len(valueText) = 0 Then
                    errorCount = errorCount + 1
                    errors = errors & errorCount & _
                             ". Не заполнен шифр пробы: " & _
                             rangeName & vbCrLf
                End If

            Case "TaskId"

                If Len(valueText) > 0 Then

                    If Not IsNumeric(rng.Value) Then

                        AddValidationError _
                            errors, _
                            errorCount, _
                            "Поле """ & _
                            CVD_FieldCaption(fieldName) & _
                            """ должно содержать число: " & _
                            rangeName, _
                            rng

                    ElseIf CDbl(rng.Value) <> _
                           Fix(CDbl(rng.Value)) Then

                        AddValidationError _
                            errors, _
                            errorCount, _
                            "Поле """ & _
                            CVD_FieldCaption(fieldName) & _
                            """ должно содержать целое число: " & _
                            rangeName, _
                            rng

                    ElseIf CDbl(rng.Value) < 0 Then

                        AddValidationError _
                            errors, _
                            errorCount, _
                            "Поле """ & _
                            CVD_FieldCaption(fieldName) & _
                            """ не может быть отрицательным: " & _
                            rangeName, _
                            rng

                    ElseIf CDbl(rng.Value) > 0 Then

                        Set sampleCodeRange = Nothing
                        Set transferTemperatureRange = Nothing

                        Set sampleCodeRange = CVD_FormRange( _
                            sheetCVD, _
                            formIndex, _
                            "sampleCode" _
                        )

                        Set transferTemperatureRange = _
                            CVD_FormRange( _
                                sheetCVD, _
                                formIndex, _
                                "Ttransfer" _
                            )

                        If Not sampleCodeRange Is Nothing _
                           And Not transferTemperatureRange Is Nothing Then

                            If Not IsError(sampleCodeRange.Value) _
                               And Not IsError( _
                                   transferTemperatureRange.Value _
                               ) Then

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

                        End If

                    End If

                End If

            Case "analyseNumber"

                If Len(valueText) = 0 Then
                    errorCount = errorCount + 1
                    errors = errors & errorCount & _
                             ". Не указан номер анализа: " & _
                             rangeName & vbCrLf

                ElseIf Not IsNumeric(rng.Value) Then
                    errorCount = errorCount + 1
                    errors = errors & errorCount & _
                             ". Номер анализа должен быть числом: " & _
                             rangeName & vbCrLf
                End If

            Case "dateStart", "dateEnd"

                If Len(valueText) > 0 Then
                    If Not IsDate(rng.Value) Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Значение должно быть датой: " & _
                                 rangeName & vbCrLf
                    End If
                End If

            Case "timeStart", "timeEnd"

                If Len(valueText) > 0 Then
                    If Not rng.Text Like "##:##" _
                       And Not rng.Text Like "#:##" Then

                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Значение должно быть временем: " & _
                                 rangeName & vbCrLf
                    End If
                End If

            Case "Pabs"

                ' Пока без дополнительной проверки.

            Case "deltaMatBalance", _
                 "Pstart", _
                 "Ptransfer", _
                 "Tstart", _
                 "Ttransfer", _
                 "Vstart", _
                 "Vtransfer"

                If Len(valueText) > 0 Then
                    If Not IsNumeric(rng.Value) Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Значение должно быть числом: " & _
                                 rangeName & vbCrLf
                    End If
                End If

            Case "Punit", "pvtCell"

                ' Пока могут быть пустыми.

        End Select

NextResultHeader:
    Next headerCell

    ' ============================================================
    ' 2. ПОЛЯ CVD_SOURCEDATA ПО STEP
    ' ============================================================

    Dim steps As Variant
    Dim stepIndex As Long
    Dim stepKey As String

    steps = CVD_StepKeys()

    Dim lphText As String
    Dim gphText As String
    Dim lphRange As Range
    Dim gphRange As Range

    Dim mgasRange As Range
    Dim vgasRange As Range
    Dim vgtotRange As Range

    For stepIndex = LBound(steps) To UBound(steps)

        stepKey = CStr(steps(stepIndex))

        If Not CVD_StepHasData(formIndex, stepKey) Then
            GoTo NextStep
        End If

        For Each headerCell _
            In CVDSourceDataTable.HeaderRowRange.Cells

            fieldName = CStr(headerCell.Value)

            Select Case fieldName

                Case COL_CVD_ROW_ID_DB, _
                     COL_CVD_RESULT_ID_DB, _
                     COL_CVD_STEP_DB

                    GoTo NextSourceHeader

            End Select

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
                         stepKey & "_" & fieldName & "." & _
                         vbCrLf
                GoTo NextSourceHeader
            End If

            If IsError(rng.Value) Then
                errorCount = errorCount + 1
                errors = errors & errorCount & ". " & _
                         rangeName & _
                         " содержит ошибку Excel: " & _
                         rng.Text & " (" & _
                         rng.Worksheet.name & "!" & _
                         rng.Address(False, False) & ")" & _
                         vbCrLf
                GoTo NextSourceHeader
            End If

            valueText = Trim$(CStr(rng.Value))

            Select Case fieldName

                Case "operator"

                    If Len(valueText) = 0 Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Не указан исполнитель: " & _
                                 rangeName & vbCrLf

                    ElseIf Not IsValidExecutor(rng.Value) Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Неверно указан исполнитель: " & _
                                 rangeName & vbCrLf
                    End If

                Case "date"

                    If Len(valueText) = 0 Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Не указана дата step: " & _
                                 rangeName & vbCrLf

                    ElseIf Not IsDate(rng.Value) Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Значение должно быть датой: " & _
                                 rangeName & vbCrLf
                    End If

                Case "sampleCodeLPh", "sampleCodeGPh"

                    ' Проверяются общей парой после цикла полей.

                Case "Ppic"

                    If Len(valueText) > 0 Then
                        If Not valueText Like "20##-P##" Then
                            errorCount = errorCount + 1
                            errors = errors & errorCount & _
                                     ". Формат Ppic должен быть " & _
                                     "20##-P##: " & _
                                     rangeName & vbCrLf
                        End If
                    End If

                Case "Upic"

                    ' Пока без строгой проверки.

                Case "gorEquipment"

                    If Len(valueText) = 0 Then
                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". Не указано оборудование: " & _
                                 rangeName & vbCrLf
                    End If

                Case "cellFluidMass", "d20", "dGas", _
                     "evacuatedFluidMass", "GF", "initPatm", _
                     "m0", "m0Trap", "m1", "m1Trap", "m2", _
                     "mgas", "mLiq", "Patm", "picDen", _
                     "stepPcell", "stepTcell", "T", _
                     "transferPcell", "V1", "V2", "Vcell", _
                     "Vcellend", "VCellLiq", "Vgas", "VgasSt", _
                     "VGCyl", "VGTot", "VHeCyl", "VHeTot", _
                     "VLiq", "VLiqRelative", "Vtransfer"

                    If Len(valueText) > 0 Then
                        If Not IsNumeric(rng.Value) Then
                            errorCount = errorCount + 1
                            errors = errors & errorCount & _
                                     ". Значение должно быть числом: " & _
                                     rangeName & vbCrLf
                        End If
                    End If

            End Select

NextSourceHeader:
        Next headerCell

        ' --------------------------------------------------------
        ' Хотя бы один шифр полученной пробы должен быть указан.
        ' --------------------------------------------------------

        Set lphRange = Nothing
        Set gphRange = Nothing

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
            If Not IsError(lphRange.Value) Then
                lphText = Trim$(CStr(lphRange.Value))
            End If
        End If

        If Not gphRange Is Nothing Then
            If Not IsError(gphRange.Value) Then
                gphText = Trim$(CStr(gphRange.Value))
            End If
        End If

        If Not lphRange Is Nothing _
           And Not gphRange Is Nothing Then

            If Len(lphText) = 0 And Len(gphText) = 0 Then
                errorCount = errorCount + 1
                errors = errors & errorCount & _
                         ". В CVD_" & formIndex & _
                         "_" & stepKey & _
                         " не указан ни один шифр " & _
                         "полученной пробы: " & _
                         "sampleCodeLPh/sampleCodeGPh." & _
                         vbCrLf
            End If

        End If

        ' --------------------------------------------------------
        ' Если заполнена mgas, требуется Vgas или VGTot.
        ' --------------------------------------------------------

        Set mgasRange = Nothing
        Set vgasRange = Nothing
        Set vgtotRange = Nothing

        Set mgasRange = CVD_FormRange( _
            sheetCVD, _
            formIndex, _
            stepKey & "_mgas" _
        )

        Set vgasRange = CVD_FormRange( _
            sheetCVD, _
            formIndex, _
            stepKey & "_Vgas" _
        )

        Set vgtotRange = CVD_FormRange( _
            sheetCVD, _
            formIndex, _
            stepKey & "_VGTot" _
        )

        If Not mgasRange Is Nothing _
           And Not vgasRange Is Nothing _
           And Not vgtotRange Is Nothing Then

            If Not IsError(mgasRange.Value) _
               And Not IsError(vgasRange.Value) _
               And Not IsError(vgtotRange.Value) Then

                If Len(Trim$(CStr(mgasRange.Value))) > 0 Then

                    If Len(Trim$(CStr(vgasRange.Value))) = 0 _
                       And Len( _
                           Trim$(CStr(vgtotRange.Value)) _
                       ) = 0 Then

                        errorCount = errorCount + 1
                        errors = errors & errorCount & _
                                 ". В CVD_" & formIndex & _
                                 "_" & stepKey & _
                                 " указана mgas, но не указан " & _
                                 "Vgas/VGTot." & vbCrLf

                    End If

                End If

            End If

        End If

NextStep:
    Next stepIndex

    If errorCount > 0 Then
        MsgBox _
            "CVD_" & formIndex & _
            " не сохранена. Исправьте ошибки:" & _
            vbCrLf & vbCrLf & errors, _
            vbExclamation
        Exit Function
    End If

    Validate_CVD_Form = True

End Function
```

### Что здесь принципиально изменено

- `rangeName` теперь применяется только как подпись вида `CVD_2_st1_Ppic` в сообщении.
- Сам диапазон всегда получается через `CVD_FormRange`.
- Для `TaskId` значения `sampleCode` и `Ttransfer` также берутся через `CVD_FormRange`.
- Больше нет `NamedRangeExists` для форм 2+.
- Отсутствие поля означает отсутствие базового имени `CVD_1_<поле>`, а не отсутствие отдельного имени текущей формы.

## Шаг 2. Полностью заменить `CVD_StepHasData`

Найти функцию:

```vb
Private Function CVD_StepHasData( _
```

Удалить её целиком до `End Function` и вставить:

```vb
Private Function CVD_StepHasData( _
    ByVal formIndex As Long, _
    ByVal stepKey As String _
) As Boolean

    CVD_StepHasData = False

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

## Шаг 3. Полностью заменить `FindEmptyFormCVD`

Найти:

```vb
Public Function FindEmptyFormCVD() As Long
```

Удалить функцию целиком и вставить:

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

## Шаг 4. Проверить `wrappers.bas`

Это изменение в другом модуле.

Найти:

```vb
Public Sub Add_CVD_Form_wrap()
```

Внутри должно быть:

```vb
CallMacro "AddFormCVD", "CVD"
```

Затем найти:

```vb
Public Sub Delete_CVD_Form_wrap()
```

Внутри должно быть:

```vb
CallMacro "DeleteLastListCVD", "CVD"
```

Старых имён:

```text
InsertListCVD2
DeleteLastListCVD2
```

остаться не должно.

## Шаг 5. Контрольный поиск по `CVD.bas`

После трёх замен выполнить поиск в VBE по следующим строкам:

```text
NamedRangeExists
sheetCVD.Range("CVD_
sheetCVD.Range(rangeName)
CVD_2_sampleCode
checkNumberAnalys2
CopyTableRowToNamedRanges
FindFormIndexForResultRow
ActiveCell.Name.Name
```

В текущем `CVD.bas` после исправления совпадений быть не должно.

Исключение: внутри `CVD_FormRange` должна остаться эта строка:

```vb
Set baseRange = ws.Parent.Names( _
    "CVD_1_" & fieldKey _
).RefersToRange
```

Это единственное место, которое физически обращается к именованному диапазону.

## Шаг 6. Компиляция и короткий тест

1. Выполнить `Debug -> Compile VBAProject`.
2. На форме 2, пока её старые имена ещё существуют, проверить `Validate_CVD`.
3. Удалить имена `CVD_2_*` и снова проверить форму 2.
4. Проверить определение заполненного step кликом по форме 2.
5. Проверить `FindEmptyFormCVD` через обновление проекта.
6. Добавить форму кнопкой Ribbon и убедиться, что новые имена не создаются.
7. Удалить последнюю форму кнопкой Ribbon.

Текущие значения геометрии в присланном файле:

```vb
Private Const CVD_FORM_COLUMN_COUNT As Long = 16
Private Const CVD_FIRST_STEP_COL As Long = 5
Private Const CVD_LAST_REGULAR_STEP_COL As Long = 14
Private Const CVD_FINAL_STEP_COL As Long = 15
Private Const CVD_MAX_FORMS As Long = 6
```

Они выглядят как уже намеренно исправленные тобой значения, поэтому в данном продолжении их менять не требуется.
