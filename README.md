Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim lastCol As Long
    Dim targetLastCol As Long
    Dim cell As Range
    Dim sourceRowRange As Range
    Dim targetRowRange As Range
    Dim response As VbMsgBoxResult
    Dim matchFound As Boolean
    Dim i As Long

    ' Установите ссылку на лист назначения
    Set wsTarget = ThisWorkbook.Sheets("данные по оплате")

    ' Проверяем, изменился ли столбец A
    If Not Intersect(Target, Me.Columns("A")) Is Nothing Then
        Application.EnableEvents = False ' Отключаем события для предотвращения зацикливания
        On Error GoTo CleanUp ' Устанавливаем обработчик ошибок

        For Each cell In Target
            If cell.Value = 1 Then
                ' Добавление признака 1: копируем строку на лист "данные по оплате"
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row + 1
                lastCol = Me.Cells(cell.Row, Me.Columns.Count).End(xlToLeft).Column

                ' Проверяем, есть ли данные для копирования
                If lastCol >= 2 Then
                    Set sourceRowRange = Me.Range(Me.Cells(cell.Row, 2), Me.Cells(cell.Row, lastCol))
                    Set targetRowRange = wsTarget.Cells(lastRowTarget, "D").Resize(1, sourceRowRange.Columns.Count)

                    ' Копируем данные и формат
                    sourceRowRange.Copy
                    targetRowRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    targetRowRange.PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
                End If

            ElseIf cell.Value = "" Then
                ' Удаление признака 1: поиск совпадений на листе "данные по оплате"
                matchFound = False
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row
                lastCol = Me.Cells(cell.Row, Me.Columns.Count).End(xlToLeft).Column
                
                ' Сохраняем данные текущей строки для сравнения
                If lastCol >= 2 Then
                    Set sourceRowRange = Me.Range(Me.Cells(cell.Row, 2), Me.Cells(cell.Row, lastCol))
                Else
                    GoTo NextCell ' Пропускаем, если нет данных
                End If

                ' Ищем совпадающую строку на листе "данные по оплате"
                For i = 1 To lastRowTarget
                    targetLastCol = wsTarget.Cells(i, wsTarget.Columns.Count).End(xlToLeft).Column
                    If targetLastCol >= 4 Then
                        Set targetRowRange = wsTarget.Range(wsTarget.Cells(i, 4), wsTarget.Cells(i, 3 + sourceRowRange.Columns.Count))

                        ' Сравниваем данные
                        If CompareRanges(sourceRowRange, targetRowRange) Then
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next i

                ' Если совпадение найдено, запрос подтверждения на удаление
                If matchFound Then
                    response = MsgBox("Вы действительно хотите удалить строку с данными: " & targetRowRange.Address & "?", vbYesNo + vbQuestion, "Подтверждение удаления")
                    
                    If response = vbYes Then
                        wsTarget.Rows(i).Delete
                    ElseIf response = vbNo Then
                        MsgBox "Удаление отменено. Макрос остановлен.", vbInformation, "Операция отменена"
                        GoTo CleanUp ' Завершаем выполнение макроса
                    End If
                End If
            End If
NextCell:
        Next cell

CleanUp:
        Application.EnableEvents = True ' Включаем события обратно
    End If
End Sub

' Функция для сравнения двух диапазонов
Private Function CompareRanges(rng1 As Range, rng2 As Range) As Boolean
    Dim i As Long
    Dim isMatch As Boolean

    ' Убедимся, что диапазоны одинаковой ширины
    If rng1.Columns.Count <> rng2.Columns.Count Then
        CompareRanges = False
        Exit Function
    End If

    isMatch = True
    For i = 1 To rng1.Columns.Count
        If rng1.Cells(1, i).Value <> rng2.Cells(1, i).Value Then
            isMatch = False
            Exit For
        End If
    Next i

    CompareRanges = isMatch
End Function