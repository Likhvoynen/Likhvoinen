Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim cell As Range
    Dim sourceRowRange As Range
    Dim targetRowRange As Range
    Dim matchFound As Boolean
    Dim response As VbMsgBoxResult
    Dim i As Long
    Dim lastCol As Long
    Dim targetLastCol As Long

    ' Установите ссылку на лист назначения
    Set wsTarget = ThisWorkbook.Sheets("данные по оплате")

    ' Проверяем, изменился ли столбец A
    If Not Intersect(Target, Me.Columns("A")) Is Nothing Then
        Application.EnableEvents = False ' Отключаем события для предотвращения зацикливания
        On Error GoTo CleanUp ' Устанавливаем обработчик ошибок

        For Each cell In Target
            If cell.Value = "" Then ' Если значение ячейки в столбце A удалено
                ' Запоминаем данные строки с текущего листа, начиная со столбца B
                lastCol = Me.Cells(cell.Row, Me.Columns.Count).End(xlToLeft).Column
                If lastCol >= 2 Then
                    Set sourceRowRange = Me.Range(Me.Cells(cell.Row, 2), Me.Cells(cell.Row, lastCol))
                Else
                    GoTo NextCell ' Пропустить, если нет данных для сравнения
                End If

                ' Поиск совпадающей строки на листе "данные по оплате"
                matchFound = False
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row
                
                For i = 1 To lastRowTarget
                    ' Определяем диапазон текущей строки на листе "данные по оплате"
                    targetLastCol = wsTarget.Cells(i, wsTarget.Columns.Count).End(xlToLeft).Column
                    If targetLastCol >= 4 Then ' Убедиться, что есть данные для сравнения
                        Set targetRowRange = wsTarget.Range(wsTarget.Cells(i, 4), wsTarget.Cells(i, 3 + sourceRowRange.Columns.Count))

                        ' Сравниваем данные строки
                        If CompareRanges(sourceRowRange, targetRowRange) Then
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next i

                ' Если совпадение найдено, запросить подтверждение удаления
                If matchFound Then
                    response = MsgBox("Вы действительно хотите удалить строку с данными: " & targetRowRange.Address & "?", vbYesNo + vbQuestion, "Подтверждение удаления")
                    
                    If response = vbYes Then
                        wsTarget.Rows(i).Delete
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
    Dim cell1 As Range, cell2 As Range
    Dim isMatch As Boolean
    Dim i As Long

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