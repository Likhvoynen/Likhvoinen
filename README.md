Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsTarget As Worksheet
    Dim lastRowTarget As Long
    Dim lastCol As Long
    Dim cell As Range
    Dim foundCell As Range
    Dim response As VbMsgBoxResult

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
                    ' Вставляем только значения
                    wsTarget.Cells(lastRowTarget, "D").Resize(1, lastCol - 1).Value = Me.Range(Me.Cells(cell.Row, 2), Me.Cells(cell.Row, lastCol)).Value
                End If
            ElseIf cell.Value = "" Then ' Если значение ячейки удалено
                ' Найдите строку с данными на листе "данные по оплате"
                Set foundCell = wsTarget.Columns("D").Find(What:=Me.Cells(cell.Row, 2).Value, LookIn:=xlValues, LookAt:=xlWhole)

                If Not foundCell Is Nothing Then
                    ' Запросите подтверждение на удаление
                    response = MsgBox("Вы действительно хотите удалить строку с данными: " & foundCell.Address & "?", vbYesNo + vbQuestion, "Подтверждение удаления")

                    If response = vbYes Then
                        ' Удалите строку
                        wsTarget.Rows(foundCell.Row).Delete
                    End If
                End If
            End If
        Next cell

CleanUp:
        Application.EnableEvents = True ' Включаем события обратно
    End If
End Sub