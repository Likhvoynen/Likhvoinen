Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim lastCol As Long
    Dim targetLastCol As Long
    Dim cell As Range
    Dim response As VbMsgBoxResult
    Dim matchFound As Boolean
    Dim sourceRowRange As Range
    Dim targetRowRange As Range
    Dim i As Long

    ' Установите ссылку на лист назначения
    Set wsTarget = ThisWorkbook.Sheets("данные по оплате")

    ' Проверяем, изменился ли столбец A
    If Not Intersect(Target, Me.Columns("A")) Is Nothing Then
        Application.EnableEvents = False ' Отключаем события для предотвращения зацикливания
        On Error GoTo CleanUp ' Устанавливаем обработчик ошибок

        For Each cell In Target
            If cell.Value = 1 Then
                ' Найдите первую пустую строку на листе "данные по оплате"
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row + 1
                
                ' Определите последний столбец с данными в текущей строке
                lastCol = Me.Cells(cell.Row, Me.Columns.Count).End(xlToLeft).Column
                
                ' Копируйте данные начиная со столбца B до последнего столбца с данными
                If lastCol >= 2 Then
                    Set sourceRowRange = Me.Range(Me.Cells(cell.Row, 2), Me.Cells(cell.Row, lastCol))
                    Set targetRowRange = wsTarget.Cells(lastRowTarget, "D").Resize(1, lastCol - 1)
                    
                    ' Копируем данные и формат
                    sourceRowRange.Copy
                    targetRowRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    targetRowRange.PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
                End If
            ElseIf cell.Value = "" Then ' Если значение ячейки удалено
                ' Поиск строки для удаления
                matchFound = False
                lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "D").End(xlUp).Row
                lastCol = Me.Cells(cell.Row, Me.Columns.Count).End(xlToLeft).Column
                
                For i = 1 To lastRowTarget
                    ' Диапазон данных текущей строки на листе назначения
                    targetLastCol = wsTarget.Cells(i, wsTarget.Columns.Count).End(xlToLeft).Column
                    If targetLastCol >= 4 Then
                        Set targetRowRange = wsTarget.Range(wsTarget.Cells(i, 4), wsTarget.Cells(i, 3 + lastCol - 1))
                        
                        ' Сравниваем данные строки
                        If Application.WorksheetFunction.CountIfs(targetRowRange, Me.Range(Me.Cells(cell.Row, 2), Me.Cells(cell.Row, lastCol))) = targetRowRange.Columns.Count Then
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next i
                
                ' Если совпадение найдено, запросите подтверждение на удаление
                If matchFound Then
                    response = MsgBox("Вы действительно хотите удалить строку с данными: " & targetRowRange.Address & "?", vbYesNo + vbQuestion, "Подтверждение удаления")
                    
                    If response = vbYes Then
                        wsTarget.Rows(i).Delete
                    End If
                End If
            End If
        Next cell

CleanUp:
        Application.EnableEvents = True ' Включаем события обратно
    End If
End Sub