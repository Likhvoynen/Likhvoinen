Sub TransferData()
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    Dim lastRowTarget As Long
    Dim sourceDate As Variant, cell As Range
    Dim wbTarget As Workbook
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim rowsCopied As Long
    rowsCopied = 0

    ' Проверка, что открыто как минимум два файла (исходный и целевой)
    If Workbooks.Count < 2 Then
        MsgBox "Необходимо открыть исходный и целевой файлы.", vbExclamation
        Exit Sub
    End If

    ' Установка активного файла как исходного
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.ActiveSheet
    
    ' Запрос даты у пользователя
    sourceDate = Application.InputBox("Введите дату для фильтрации (ДД.ММ.ГГГГ):", Type:=2)
    If sourceDate = False Then Exit Sub ' Пользователь нажал Отмена

    ' Проверка на правильный формат даты
    If Not IsDate(sourceDate) Then
        MsgBox "Неправильный формат даты. Пожалуйста, введите корректную дату."
        Exit Sub
    End If
    
    ' Установка второго открытого файла как целевого (не ThisWorkbook)
    For Each wbTarget In Workbooks
        If wbTarget.Name <> wbSource.Name Then
            Set wsTarget = wbTarget.Sheets("план на месяц")
            Exit For
        End If
    Next wbTarget
    
    If wsTarget Is Nothing Then
        MsgBox "Лист 'план на месяц' не найден в целевом файле.", vbExclamation
        Exit Sub
    End If

    ' Нахождение последней заполненной строки в столбце AV на целевом листе
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "AV").End(xlUp).Row + 1

    ' Перебор каждой ячейки в столбце B, начиная с B2
    For Each cell In wsSource.Range("B2:B" & wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row)
        If cell.Value = CDate(sourceDate) Then
            With wsTarget
                .Cells(lastRowTarget, "C").Value = cell.Offset(0, 19).Value ' Столбец U -> C
                .Cells(lastRowTarget, "D").Value = cell.Offset(0, 20).Value ' Столбец V -> D
                .Cells(lastRowTarget, "H").Value = cell.Offset(0, 21).Value ' Столбец W -> H
                .Cells(lastRowTarget, "M").Value = cell.Offset(0, 23).Value ' Столбец Y -> M
                .Cells(lastRowTarget, "O").Value = cell.Offset(0, 25).Value ' Столбец AA -> O
                .Cells(lastRowTarget, "R").Value = cell.Offset(0, 3).Value  ' Столбец D -> R
                .Cells(lastRowTarget, "S").Value = cell.Offset(0, 8).Value  ' Столбец I -> S
                .Cells(lastRowTarget, "T").Value = cell.Offset(0, 6).Value  ' Столбец G -> T
                .Cells(lastRowTarget, "U").Value = cell.Offset(0, 7).Value  ' Столбец H -> U
                .Cells(lastRowTarget, "V").Value = cell.Offset(0, 13).Value ' Столбец N -> V
                .Cells(lastRowTarget, "W").Value = cell.Offset(0, 16).Value ' Столбец Q -> W
                .Cells(lastRowTarget, "AE").Value = cell.Offset(0, 19).Value ' Столбец T -> AE
                .Cells(lastRowTarget, "AI").Value = cell.Offset(0, 23).Value ' Столбец X -> AI
                .Cells(lastRowTarget, "AK").Value = cell.Offset(0, 14).Value ' Столбец O -> AK
                .Cells(lastRowTarget, "AN").Value = cell.Offset(0, 9).Value  ' Столбец J -> AN
                .Cells(lastRowTarget, "AB").Value = cell.Offset(0, 1).Value  ' Столбец B -> AB
            End With
            lastRowTarget = lastRowTarget + 1
            rowsCopied = rowsCopied + 1
        End If
    Next cell
    
    ' Сообщение об успехе
    MsgBox "Копирование завершено. Количество перенесенных строк: " & rowsCopied, vbInformation
End Sub








' Определение последней строки в столбце B
lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
MsgBox "Последняя строка в столбце B: " & lastRowSource
MsgBox "Диапазон данных в B: " & wsSource.Range("B2:B" & lastRowSource).Address





Sub TransferData()
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    Dim lastRowTarget As Long, lastRowSource As Long
    Dim sourceDate As Variant, dateColumn As Range
    Dim cell As Range
    Dim wbTarget As Workbook
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet

    ' Проверка, что открыто как минимум два файла (исходный и целевой)
    If Workbooks.Count < 2 Then
        MsgBox "Необходимо открыть исходный и целевой файлы.", vbExclamation
        Exit Sub
    End If

    ' Установка активного файла как исходного
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.ActiveSheet
    
    ' Запрос даты у пользователя
    sourceDate = Application.InputBox("Введите дату для фильтрации (ДД.ММ.ГГГГ):", Type:=2)
    If sourceDate = False Then Exit Sub ' Пользователь нажал Отмена

    ' Проверка на правильный формат даты
    If Not IsDate(sourceDate) Then
        MsgBox "Неправильный формат даты. Пожалуйста, введите корректную дату."
        Exit Sub
    End If

    ' Установка второго открытого файла как целевого (не ThisWorkbook)
    For Each wbTarget In Workbooks
        If wbTarget.Name <> wbSource.Name Then
            Set wsTarget = wbTarget.Sheets("план на месяц")
            Exit For
        End If
    Next wbTarget
    
    If wsTarget Is Nothing Then
        MsgBox "Лист 'план на месяц' не найден в целевом файле.", vbExclamation
        Exit Sub
    End If

    ' Нахождение последней заполненной строки в столбце AV на целевом листе
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "AV").End(xlUp).Row + 1
    MsgBox "Начальная строка для вставки в целевом файле: " & lastRowTarget

    ' Нахождение последней строки на исходном листе
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    MsgBox "Последняя строка в исходном файле: " & lastRowSource

    ' Определение диапазона столбца B с данными
    Set dateColumn = wsSource.Range("B2:B" & lastRowSource)
    
    Dim rowsCopied As Long
    rowsCopied = 0

    ' Перебор каждой строки с проверкой на соответствие дате
    For Each cell In dateColumn
        If cell.Value = CDate(sourceDate) Then
            With wsTarget
                .Cells(lastRowTarget, "C").Value = cell.Offset(0, 19).Value ' Столбец U -> C
                .Cells(lastRowTarget, "D").Value = cell.Offset(0, 20).Value ' Столбец V -> D
                .Cells(lastRowTarget, "H").Value = cell.Offset(0, 21).Value ' Столбец W -> H
                .Cells(lastRowTarget, "M").Value = cell.Offset(0, 23).Value ' Столбец Y -> M
                .Cells(lastRowTarget, "O").Value = cell.Offset(0, 25).Value ' Столбец AA -> O
                .Cells(lastRowTarget, "R").Value = cell.Offset(0, 3).Value  ' Столбец D -> R
                .Cells(lastRowTarget, "S").Value = cell.Offset(0, 8).Value  ' Столбец I -> S
                .Cells(lastRowTarget, "T").Value = cell.Offset(0, 6).Value  ' Столбец G -> T
                .Cells(lastRowTarget, "U").Value = cell.Offset(0, 7).Value  ' Столбец H -> U
                .Cells(lastRowTarget, "V").Value = cell.Offset(0, 13).Value ' Столбец N -> V
                .Cells(lastRowTarget, "W").Value = cell.Offset(0, 16).Value ' Столбец Q -> W
                .Cells(lastRowTarget, "AE").Value = cell.Offset(0, 19).Value ' Столбец T -> AE
                .Cells(lastRowTarget, "AI").Value = cell.Offset(0, 23).Value ' Столбец X -> AI
                .Cells(lastRowTarget, "AK").Value = cell.Offset(0, 14).Value ' Столбец O -> AK
                .Cells(lastRowTarget, "AN").Value = cell.Offset(0, 9).Value  ' Столбец J -> AN
                .Cells(lastRowTarget, "AB").Value = cell.Offset(0, 1).Value  ' Столбец B -> AB
            End With
            lastRowTarget = lastRowTarget + 1
            rowsCopied = rowsCopied + 1
        End If
    Next cell
    
    MsgBox "Копирование завершено. Количество перенесенных строк: " & rowsCopied, vbInformation
End Sub









 TransferData()
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    Dim lastRowTarget As Long, lastRowSource As Long
    Dim sourceDate As Variant, dateColumn As Range
    Dim cell As Range
    Dim wbTarget As Workbook
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet

    ' Проверка, что открыто как минимум два файла (исходный и целевой)
    If Workbooks.Count < 2 Then
        MsgBox "Необходимо открыть исходный и целевой файлы.", vbExclamation
        Exit Sub
    End If

    ' Установка активного файла как исходного
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.ActiveSheet
    
    ' Запрос даты у пользователя
    sourceDate = Application.InputBox("Введите дату для фильтрации (ДД.ММ.ГГГГ):", Type:=2)
    If sourceDate = False Then Exit Sub ' Пользователь нажал Отмена

    ' Проверка на правильный формат даты
    If Not IsDate(sourceDate) Then
        MsgBox "Неправильный формат даты. Пожалуйста, введите корректную дату."
        Exit Sub
    End If
    
    ' Установка второго открытого файла как целевого (не ThisWorkbook)
    For Each wbTarget In Workbooks
        If wbTarget.Name <> wbSource.Name Then
            Set wsTarget = wbTarget.Sheets("план на месяц")
            Exit For
        End If
    Next wbTarget
    
    ' Нахождение последней заполненной строки в столбце AV на целевом листе
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "AV").End(xlUp).Row + 1
    
    ' Нахождение последней строки на исходном листе
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    ' Определение диапазона столбца B с данными
    Set dateColumn = wsSource.Range("B2:B" & lastRowSource)
    
    ' Перебор каждой строки с проверкой на соответствие дате
    For Each cell In dateColumn
        If cell.Value = CDate(sourceDate) Then
            With wsTarget
                .Cells(lastRowTarget, "C").Value = cell.Offset(0, 19).Value ' Столбец U -> C
                .Cells(lastRowTarget, "D").Value = cell.Offset(0, 20).Value ' Столбец V -> D
                .Cells(lastRowTarget, "H").Value = cell.Offset(0, 21).Value ' Столбец W -> H
                .Cells(lastRowTarget, "M").Value = cell.Offset(0, 23).Value ' Столбец Y -> M
                .Cells(lastRowTarget, "O").Value = cell.Offset(0, 25).Value ' Столбец AA -> O
                .Cells(lastRowTarget, "R").Value = cell.Offset(0, 3).Value  ' Столбец D -> R
                .Cells(lastRowTarget, "S").Value = cell.Offset(0, 8).Value  ' Столбец I -> S
                .Cells(lastRowTarget, "T").Value = cell.Offset(0, 6).Value  ' Столбец G -> T
                .Cells(lastRowTarget, "U").Value = cell.Offset(0, 7).Value  ' Столбец H -> U
                .Cells(lastRowTarget, "V").Value = cell.Offset(0, 13).Value ' Столбец N -> V
                .Cells(lastRowTarget, "W").Value = cell.Offset(0, 16).Value ' Столбец Q -> W
                .Cells(lastRowTarget, "AE").Value = cell.Offset(0, 19).Value ' Столбец T -> AE
                .Cells(lastRowTarget, "AI").Value = cell.Offset(0, 23).Value ' Столбец X -> AI
                .Cells(lastRowTarget, "AK").Value = cell.Offset(0, 14).Value ' Столбец O -> AK
                .Cells(lastRowTarget, "AN").Value = cell.Offset(0, 9).Value  ' Столбец J -> AN
                .Cells(lastRowTarget, "AB").Value = cell.Offset(0, 1).Value  ' Столбец B -> AB
            End With
            lastRowTarget = lastRowTarget + 1
        End If
    Next cell
    
    ' Сообщение о завершении
    MsgBox "Данные успешно перенесены.", vbInformation
End Sub









Sub TransferData()
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    Dim lastRowTarget As Long, lastRowSource As Long
    Dim sourceDate As Variant, dateColumn As Range
    Dim cell As Range
    Dim wbTarget As Workbook
    
    ' Запрос даты у пользователя
    sourceDate = Application.InputBox("Введите дату для фильтрации (ДД.ММ.ГГГГ):", Type:=2)
    If sourceDate = False Then Exit Sub ' Пользователь нажал Отмена

    ' Проверка на правильный формат даты
    If Not IsDate(sourceDate) Then
        MsgBox "Неправильный формат даты. Пожалуйста, введите корректную дату."
        Exit Sub
    End If
    
    ' Установка активного листа как исходного листа
    Set sourceSheet = ThisWorkbook.ActiveSheet

    ' Открытие целевого файла
    Set wbTarget = Application.Workbooks.Open("путь_к_целевому_файлу.xlsx")
    Set targetSheet = wbTarget.Sheets("план на месяц")
    
    ' Нахождение последней заполненной строки в столбце AV на целевом листе
    lastRowTarget = targetSheet.Cells(targetSheet.Rows.Count, "AV").End(xlUp).Row + 1
    
    ' Нахождение последней строки на исходном листе
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Определение диапазона столбца B с данными
    Set dateColumn = sourceSheet.Range("B2:B" & lastRowSource)
    
    ' Перебор каждой строки с проверкой на соответствие дате
    For Each cell In dateColumn
        If cell.Value = CDate(sourceDate) Then
            With targetSheet
                .Cells(lastRowTarget, "C").Value = cell.Offset(0, 19).Value ' Столбец U -> C
                .Cells(lastRowTarget, "D").Value = cell.Offset(0, 20).Value ' Столбец V -> D
                .Cells(lastRowTarget, "H").Value = cell.Offset(0, 21).Value ' Столбец W -> H
                .Cells(lastRowTarget, "M").Value = cell.Offset(0, 23).Value ' Столбец Y -> M
                .Cells(lastRowTarget, "O").Value = cell.Offset(0, 25).Value ' Столбец AA -> O
                .Cells(lastRowTarget, "R").Value = cell.Offset(0, 3).Value  ' Столбец D -> R
                .Cells(lastRowTarget, "S").Value = cell.Offset(0, 8).Value  ' Столбец I -> S
                .Cells(lastRowTarget, "T").Value = cell.Offset(0, 6).Value  ' Столбец G -> T
                .Cells(lastRowTarget, "U").Value = cell.Offset(0, 7).Value  ' Столбец H -> U
                .Cells(lastRowTarget, "V").Value = cell.Offset(0, 13).Value ' Столбец N -> V
                .Cells(lastRowTarget, "W").Value = cell.Offset(0, 16).Value ' Столбец Q -> W
                .Cells(lastRowTarget, "AE").Value = cell.Offset(0, 19).Value ' Столбец T -> AE
                .Cells(lastRowTarget, "AI").Value = cell.Offset(0, 23).Value ' Столбец X -> AI
                .Cells(lastRowTarget, "AK").Value = cell.Offset(0, 14).Value ' Столбец O -> AK
                .Cells(lastRowTarget, "AN").Value = cell.Offset(0, 9).Value  ' Столбец J -> AN
                .Cells(lastRowTarget, "AB").Value = cell.Offset(0, 1).Value  ' Столбец B -> AB
            End With
            lastRowTarget = lastRowTarget + 1
        End If
    Next cell
    
    ' Закрытие целевого файла с сохранением
    wbTarget.Close SaveChanges:=True
    
    MsgBox "Данные успешно перенесены.", vbInformation
End Sub












Sub ДЗ_часть_3()
    Dim PivotTable As PivotTable
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim pivotTableRange As Range
    Dim sumRowsD As String, sumRowsE As String, sumRowsF As String, sumRowsG As String

    ' Сохраняем состояние сводной таблицы до изменений
    Dim pivotFieldState As Collection
    Set pivotFieldState = New Collection
    
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")
    
    ' Сохраняем фильтры полей сводной таблицы
    For Each pf In PivotTable.PivotFields
        Dim fieldState As Dictionary
        Set fieldState = New Dictionary
        fieldState.Add "Orientation", pf.orientation
        fieldState.Add "Position", pf.position
        
        Dim visibleItems As Collection
        Set visibleItems = New Collection
        For Each pi In pf.PivotItems
            If pi.Visible Then visibleItems.Add pi.Name
        Next pi
        fieldState.Add "VisibleItems", visibleItems
        
        pivotFieldState.Add fieldState, pf.Name
    Next pf
    
    ' (Далее код с изменениями сводной таблицы, созданием нового листа и форматированием...)
    
    ' Восстанавливаем сводную таблицу к исходному состоянию после всех операций
    For Each pf In PivotTable.PivotFields
        If pivotFieldState.Exists(pf.Name) Then
            Dim savedState As Dictionary
            Set savedState = pivotFieldState(pf.Name)
            
            ' Восстанавливаем ориентацию и позицию полей
            pf.orientation = savedState("Orientation")
            pf.position = savedState("Position")
            
            ' Восстанавливаем видимые элементы
            Set visibleItems = savedState("VisibleItems")
            For Each pi In pf.PivotItems
                If visibleItems.Contains(pi.Name) Then
                    pi.Visible = True
                Else
                    pi.Visible = False
                End If
            Next pi
        End If
    Next pf
End Sub






Sub ДЗ_часть_3()
    Dim PivotTable As PivotTable
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim pivotTableRange As Range
    Dim sumRowsD As String, sumRowsE As String, sumRowsF As String, sumRowsG As String

    ' Определяем лист и сводную таблицу
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")
    
    ' Добавляем новое поле "Тип СФ" в строки
    With PivotTable
        .PivotFields("Тип СФ").orientation = xlRowField
        .PivotFields("Тип СФ").position = 4 ' Устанавливаем в четвёртую позицию в строках
    End With
    
    ' Обновляем сводную таблицу
    PivotTable.RefreshTable
    
    ' Фильтруем поле "Тип СФ", убирая значения, содержащие "КА" и "Кредитовое авизо"
    Set pf = PivotTable.PivotFields("Тип СФ")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "КА") > 0 Or InStr(pi.Name, "Кредитовое авизо") > 0 Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Фильтруем поле "Категория просрочки", оставляя только нужные значения
    Set pf = PivotTable.PivotFields("Категория просрочки")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If pi.Name <> "просрочка более 60 дней" And _
           pi.Name <> "просрочка от 30 до 60 дней" And _
           pi.Name <> "просрочка от 15 до 30 дней" And _
           pi.Name <> "просрочка до 15 дней" Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
     ' Фильтруем поле "Тип СФ", убирая значения, содержащие "Факторинг" и "Суды и прочие"
    Set pf = PivotTable.PivotFields("Сегмент")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "Факторинг") > 0 Or InStr(pi.Name, "Суды и прочие") > 0 Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Сворачиваем все элементы по столбцу "Заказчик"
    Set pf = PivotTable.PivotFields("Заказчик")
    pf.ShowDetail = False
    
    ' Создаем новый лист с именем "Анализ ПДЗ"
    On Error Resume Next ' Игнорируем ошибку, если лист уже существует
    Set NewSheet = ActiveWorkbook.Sheets("Анализ ПДЗ")
    On Error GoTo 0

    If NewSheet Is Nothing Then
        Set NewSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        NewSheet.Name = "Анализ ПДЗ"
    Else
        ' Если лист уже существует, очищаем его
        NewSheet.Cells.Clear
    End If

' Окрашиваем лист "Анализ ПДЗ" в оранжевый цвет
NewSheet.Tab.Color = RGB(255, 204, 153) ' Светло-оранжевый цвет

    ' Определяем диапазон сводной таблицы
    Set pivotTableRange = PivotTable.TableRange2
    
     ' Копируем сводную таблицу с листа "Свод" на лист "Анализ ПДЗ" как значения и приводим к общему формату
    pivotTableRange.Copy
    NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues ' Вставляем как значения
    NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats ' Вставляем форматирование
    Application.CutCopyMode = False ' Очищаем буфер обмена

    With Worksheets("Анализ ПДЗ")
    .Rows(1).Delete
End With

With Worksheets("Анализ ПДЗ")
    .Range("A1:C1").Value = .Range("A2:C2").Value ' Перенос данных из A2:C2 в A1:C1
    .Rows(2).Delete ' Удаление второй строки
End With

With Worksheets("Анализ ПДЗ").Rows(1)
    .WrapText = True ' Перенос текста
    .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
    .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
    .Font.Bold = True ' Жирный шрифт
    .Font.Color = RGB(0, 0, 0) ' Черный цвет шрифта
End With
    
  ' Удаляем столбец D
    NewSheet.Columns("D").Delete

     ' Устанавливаем ширину всех столбцов на активном листе в соответствии с длиной текста
    ActiveSheet.Columns.AutoFit
    
    ' Удаляем строки, где значения в столбце H меньше 10
    lastRow = NewSheet.Cells(NewSheet.Rows.Count, "H").End(xlUp).Row
    For i = lastRow To 1 Step -1
        If IsNumeric(NewSheet.Cells(i, "H").Value) Then
            If NewSheet.Cells(i, "H").Value < 10 Then
                NewSheet.Rows(i).Delete
            End If
        End If
    Next i
    
      ' Удаляем строки, где значения в столбце С равно "Вертекс АО"
    lastRow = NewSheet.Cells(NewSheet.Rows.Count, "C").End(xlUp).Row
For i = lastRow To 1 Step -1
    If NewSheet.Cells(i, "C").Value = "ВЕРТЕКС АО" Then
        NewSheet.Rows(i).Delete
    End If
Next i

' Удаляем строки с итогом, если в них нет данных
Dim rowBelow As Long

    lastRow = Cells(Rows.Count, "A").End(xlUp).Row ' Определяем последнюю строку по столбцу A

    For currentRow = lastRow To 2 Step -1 ' Проходим с конца к началу
        ' Проверяем, содержит ли текущая строка "итог", но не "общий итог"
        If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 And _
           InStr(1, Cells(currentRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
            
            ' Ищем строку ниже
            rowBelow = currentRow + 1
            
            ' Проверяем, что следующая строка существует и содержит "итог", но не "общий итог"
            If rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) > 0 And _
               InStr(1, Cells(rowBelow, "A").Value, "общий итог", vbTextCompare) = 0 Then
                
                ' Проверяем, что между текущей и следующей строками нет других строк с данными
                Dim emptyRows As Boolean
                emptyRows = True
                
                Dim checkRow As Long
                For checkRow = currentRow + 1 To rowBelow - 1
                    If Not IsEmpty(Cells(checkRow, "A").Value) Then
                        emptyRows = False
                        Exit For
                    End If
                Next checkRow
                
                ' Если между строками только пустые строки, удаляем нижнюю строку
                If emptyRows Then
                    Rows(rowBelow).Delete
                    lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
                End If
            End If
        End If
    Next currentRow

    
  ' Найдите последнюю заполненную строку в столбцах D:G
    lastRow = Cells(Rows.Count, "D").End(xlUp).Row
    Dim lastRowH As Long
    lastRowH = Cells(Rows.Count, "G").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRow, lastRowH)

    ' Вставьте формулу суммирования в столбец H, начиная со второй строки
    For i = 2 To lastRow
        Cells(i, "H").formula = "=SUM(D" & i & ":G" & i & ")"
    Next i
    
' Вставьте формулу суммирования в столбцы D:G, начиная со второй строки
    Dim nextRow As Long
    Dim currentValue As String
    Dim totalRow As Long
    
    lastRow = Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки в столбце A
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки
    
    Do While currentRow <= lastRow
        currentValue = Cells(currentRow, 1).Value
        
        ' Поиск строк с одинаковыми значениями в столбце A до строки с "Итог"
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Определение начала диапазона для суммирования
            nextRow = totalRow - 1
            Do While nextRow >= 2 And Cells(nextRow, 1).Value = Cells(nextRow - 1, 1).Value
                nextRow = nextRow - 1
            Loop
            
            ' Вставка формулы для суммирования в строку с "Итог"
            Cells(totalRow, 4).formula = "=SUM(D" & nextRow & ":D" & totalRow - 1 & ")"
            Cells(totalRow, 5).formula = "=SUM(E" & nextRow & ":E" & totalRow - 1 & ")"
            Cells(totalRow, 6).formula = "=SUM(F" & nextRow & ":F" & totalRow - 1 & ")"
            Cells(totalRow, 7).formula = "=SUM(G" & nextRow & ":G" & totalRow - 1 & ")"
        End If
        
        currentRow = currentRow + 1
    Loop
    

' Теперь добавим формулу суммирования для строки с "Общий итог"
sumRowsD = ""
sumRowsE = ""
sumRowsF = ""
sumRowsG = ""

For r = 2 To lastRow
    If InStr(1, Cells(r, "A").Value, "итог", vbTextCompare) > 0 Then
        If InStr(1, Cells(r, "A").Value, "общий итог", vbTextCompare) = 0 Then
            If sumRowsD = "" Then
                sumRowsD = "D" & r
                sumRowsE = "E" & r
                sumRowsF = "F" & r
                sumRowsG = "G" & r
            Else
                sumRowsD = sumRowsD & ",D" & r
                sumRowsE = sumRowsE & ",E" & r
                sumRowsF = sumRowsF & ",F" & r
                sumRowsG = sumRowsG & ",G" & r
            End If
        End If
    End If
Next r

' Вставьте формулы суммирования в строку, где есть "Общий итог"
For r = 2 To lastRow
    If InStr(1, Cells(r, "A").Value, "Общий итог", vbTextCompare) > 0 Then
        If sumRowsD <> "" Then
            Cells(r, "D").formula = "=SUM(" & sumRowsD & ")"
            Cells(r, "E").formula = "=SUM(" & sumRowsE & ")"
            Cells(r, "F").formula = "=SUM(" & sumRowsF & ")"
            Cells(r, "G").formula = "=SUM(" & sumRowsG & ")"
        End If
        Exit For ' Выход из цикла после нахождения первой строки с "Общий итог"
    End If
Next r

 ' Определяем позицию для вставки новых столбцов (после столбца H)
    Set ws = ActiveWorkbook.Sheets("Анализ ПДЗ")
    Dim startColumn As Integer
    startColumn = 9 ' Столбец I (H = 8, I = 9)

    ' Вставляем новые столбцы по одному
    For i = 0 To 8
    Columns(startColumn + i).Insert Shift:=xlToRight
    Next i

    ' Добавляем заголовки столбцов
    Cells(1, startColumn).Value = "Техническая ПДЗ"
    Cells(1, startColumn + 1).Value = "Действительная ПДЗ"
    Cells(1, startColumn + 2).Value = "Дата возникновения действительной ПДЗ"
    Cells(1, startColumn + 3).Value = "КЛ СК"
    Cells(1, startColumn + 4).Value = "Крайняя дата Уведомления СК"
    Cells(1, startColumn + 5).Value = "Отсрочка СК"
    Cells(1, startColumn + 6).Value = "Отсрочка ВЕРТЕКС"
    Cells(1, startColumn + 7).Value = "Комментарий"
    Cells(1, startColumn + 8).Value = "Комментарий" ' Второй комментарий

   ' Задаем цвет заливки для заголовков новых столбцов (L1, M1, N1, O1, P1)
    Cells(1, startColumn + 3).Interior.Color = RGB(255, 204, 0) ' L1
    Cells(1, startColumn + 4).Interior.Color = RGB(255, 204, 0) ' M1
    Cells(1, startColumn + 5).Interior.Color = RGB(255, 204, 0) ' N1
    Cells(1, startColumn + 6).Interior.Color = RGB(255, 204, 0) ' O1
    Cells(1, startColumn + 7).Interior.Color = RGB(255, 204, 0) ' P1
    
    
   ' Суммирование столбцов I:J
   lastRow = Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки в столбце A
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки
    
    Do While currentRow <= lastRow
        currentValue = Cells(currentRow, 1).Value
        
        ' Поиск строк с одинаковыми значениями в столбце A до строки с "Итог"
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Определение начала диапазона для суммирования
            nextRow = totalRow - 1
            Do While nextRow >= 2 And Cells(nextRow, 1).Value = Cells(nextRow - 1, 1).Value
                nextRow = nextRow - 1
            Loop
            
            ' Вставка формулы для суммирования в строку с "Итог" для столбцов I:J
            Cells(totalRow, 9).formula = "=SUM(I" & nextRow & ":I" & totalRow - 1 & ")"
            Cells(totalRow, 10).formula = "=SUM(J" & nextRow & ":J" & totalRow - 1 & ")"
        End If
        
        currentRow = currentRow + 1
    Loop
    
    ' Теперь добавим формулу суммирования для строки с "Общий итог"
    Dim sumRowsI As String
    Dim sumRowsJ As String
    
    For r = 2 To lastRow
        If InStr(1, Cells(r, "A").Value, "итог", vbTextCompare) > 0 Then
            If InStr(1, Cells(r, "A").Value, "общий итог", vbTextCompare) = 0 Then
                If sumRowsI = "" Then
                    sumRowsI = "I" & r
                    sumRowsJ = "J" & r
                Else
                    sumRowsI = sumRowsI & ",I" & r
                    sumRowsJ = sumRowsJ & ",J" & r
                End If
            End If
        End If
    Next r

    ' Вставьте формулы суммирования в строку, где есть "Общий итог"
    For r = 2 To lastRow
        If InStr(1, ws.Cells(r, "A").Value, "Общий итог", vbTextCompare) > 0 Then
            If sumRowsI <> "" Then
                Cells(r, "I").formula = "=SUM(" & sumRowsI & ")"
                Cells(r, "J").formula = "=SUM(" & sumRowsJ & ")"
            End If
            Exit For ' Выход из цикла после нахождения первой строки с "Общий итог"
        End If
    Next r
    
    With Worksheets("Анализ ПДЗ").usedRange
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
End With

' Установка ширины столбцов D:H на 16
    ws.Columns("D:H").ColumnWidth = 16
    
    ' Автоматическая настройка ширины для остальных столбцов
    Dim col As Range
    For Each col In ws.Columns("A:C") ' Столбцы A:C
        col.AutoFit
    Next col
    For Each col In ws.Columns("I:Z") ' Столбцы I и далее (можно изменить на нужные)
        col.AutoFit
    Next col
    
    ' Установка масштаба листа на 80%
    ws.Parent.Windows(1).Zoom = 80
    
    ' Очищаем буфер обмена
    Application.CutCopyMode = False
    
End Sub



















' Удаляем строки между строками с "Итог", если они пустые
Dim deleteRow As Long
Dim currentRow As Long
Dim rowBelow As Long

For currentRow = lastRow To 2 Step -1 ' Проходим с конца к началу
    If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 And _
       InStr(1, Cells(currentRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
        
        ' Ищем строку ниже
        rowBelow = currentRow + 1
        
        ' Если следующая строка - "итог", удаляем её
        If rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) > 0 And _
           InStr(1, Cells(rowBelow, "A").Value, "общий итог", vbTextCompare) = 0 Then
            Rows(rowBelow).Delete
            lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
        End If
        
        ' Так как мы движемся вверх, можем выйти из цикла
        Exit For 
    End If
Next currentRow






' Удаляем строки между строками с "Итог", если они пустые
Dim deleteRow As Long
Dim rowBelow As Long

For currentRow = 2 To lastRow
    If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 And _
       InStr(1, Cells(currentRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
        
        ' Ищем строку ниже
        rowBelow = currentRow + 1
        Do While rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) = 0
            rowBelow = rowBelow + 1
        Loop
        
        ' Если следующая строка - "итог", удаляем её
        If rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) > 0 And _
           InStr(1, Cells(rowBelow, "A").Value, "общий итог", vbTextCompare) = 0 Then
            Rows(rowBelow).Delete
            lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
        End If
        
        currentRow = currentRow + 1 ' Переходим к следующей строке
    End If
Next currentRow












' Удаляем строки между строками с "Итог", если они пустые
Dim currentRow As Long
Dim nextRow As Long

For currentRow = 2 To lastRow
    If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 And _
       InStr(1, Cells(currentRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
        
        ' Ищем строку ниже с "итог"
        nextRow = currentRow + 1
        Do While nextRow <= lastRow And InStr(1, Cells(nextRow, "A").Value, "итог", vbTextCompare) = 0
            nextRow = nextRow + 1
        Loop
        
        ' Если найдена следующая строка с "итог", удаляем верхнюю строку
        If nextRow <= lastRow And InStr(1, Cells(nextRow, "A").Value, "итог", vbTextCompare) > 0 And _
           InStr(1, Cells(nextRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
            Rows(currentRow).Delete
            lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
            currentRow = currentRow - 1 ' Возвращаемся на одну строку назад, чтобы не пропустить следующую строку
        End If
    End If
Next currentRow





' Удаляем строки между строками с "Итог", если они пустые
Dim deleteRow As Long
Dim rowBelow As Long

For currentRow = 2 To lastRow
    If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 And _
       InStr(1, Cells(currentRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
        
        ' Ищем строку ниже
        rowBelow = currentRow + 1
        Do While rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) = 0
            rowBelow = rowBelow + 1
        Loop
        
        ' Если следующая строка - "итог", удаляем её
        If rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) > 0 And _
           InStr(1, Cells(rowBelow, "A").Value, "общий итог", vbTextCompare) = 0 Then
            Rows(rowBelow).Delete
            lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
        End If
        
        currentRow = currentRow + 1 ' Переходим к следующей строке
    End If
Next currentRow





' Удаляем строки между строками с "Итог", если они пустые
Dim deleteRow As Long
Dim rowAbove As Long
Dim rowBelow As Long

For currentRow = 2 To lastRow
    If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 And _
       InStr(1, Cells(currentRow, "A").Value, "общий итог", vbTextCompare) = 0 Then
        deleteRow = currentRow ' Запоминаем строку с "Итог"
        
        ' Ищем строку выше
        rowAbove = deleteRow - 1
        Do While rowAbove >= 2 And InStr(1, Cells(rowAbove, "A").Value, "итог", vbTextCompare) = 0
            rowAbove = rowAbove - 1
        Loop
        
        ' Ищем строку ниже
        rowBelow = deleteRow + 1
        Do While rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) = 0
            rowBelow = rowBelow + 1
        Loop
        
        ' Если обе строки (выше и ниже) не содержат "итог" и не равны "общий итог", удаляем строку с "итог"
        If rowAbove > 1 And (rowBelow > lastRow Or rowBelow - 1 = deleteRow) Then
            Rows(deleteRow).Delete
            lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
            currentRow = currentRow - 1 ' Сдвигаем текущую строку вверх
        End If
    End If
Next currentRow




' Удаляем строки между строками с "Итог", если они пустые
Dim deleteRow As Long
Dim rowAbove As Long
Dim rowBelow As Long

For currentRow = 2 To lastRow
    If InStr(1, Cells(currentRow, "A").Value, "итог", vbTextCompare) > 0 Then
        deleteRow = currentRow ' Запоминаем строку с "Итог"
        
        ' Ищем строку выше
        rowAbove = deleteRow - 1
        Do While rowAbove >= 2 And InStr(1, Cells(rowAbove, "A").Value, "итог", vbTextCompare) = 0
            rowAbove = rowAbove - 1
        Loop
        
        ' Ищем строку ниже
        rowBelow = deleteRow + 1
        Do While rowBelow <= lastRow And InStr(1, Cells(rowBelow, "A").Value, "итог", vbTextCompare) = 0
            rowBelow = rowBelow + 1
        Loop
        
        ' Если обе строки (выше и ниже) не содержат "итог", удаляем строку с "итог"
        If rowAbove > 1 And (rowBelow > lastRow Or rowBelow - 1 = deleteRow) Then
            Rows(deleteRow).Delete
            lastRow = lastRow - 1 ' Обновляем последнюю строку после удаления
            currentRow = currentRow - 1 ' Сдвигаем текущую строку вверх
        End If
    End If
Next currentRow





' Удаление значений #ССЫЛКА! из столбцов D:G
For Each cell In NewSheet.Range("D2:G" & lastRow)
    If IsError(cell.Value) And cell.Value = CVErr(xlErrRef) Then ' Проверка на #ССЫЛКА!
        cell.ClearContents ' Удаляем содержимое ячейки
    End If
Next cell




' Удаление значений #ССЫЛКА! из столбцов D:G
For Each cell In NewSheet.Range("D2:G" & lastRow)
    If IsError(cell.Value) And cell.Value = CVErr(xlErrRef) Then ' Проверка на #ССЫЛКА!
        cell.ClearContents ' Удаляем содержимое ячейки
    End If
Next cell



' Удаление значений #ССЫЛКА! из столбцов D:G
Dim cell As Range

For Each cell In NewSheet.Range("D2:G" & lastRow)
    If cell.Value = CVErr(xlErrRef) Then ' Проверка на #ССЫЛКА!
        cell.Value = "" ' Заменяем на пустую строку
    End If
Next cell





Sub ДЗ_часть_3()
    Dim PivotTable As PivotTable
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim pivotTableRange As Range
    Dim sumRowsD As String, sumRowsE As String, sumRowsF As String, sumRowsG As String
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextRow As Long
    Dim currentValue As String
    Dim totalRow As Long
    Dim i As Long

    ' Определяем лист и сводную таблицу
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")
    
    ' Добавляем новое поле "Тип СФ" в строки
    With PivotTable
        .PivotFields("Тип СФ").Orientation = xlRowField
        .PivotFields("Тип СФ").Position = 4 ' Устанавливаем в четвёртую позицию в строках
    End With
    
    ' Обновляем сводную таблицу
    PivotTable.RefreshTable
    
    ' Фильтруем поле "Тип СФ", убирая значения, содержащие "КА" и "Кредитовое авизо"
    Set pf = PivotTable.PivotFields("Тип СФ")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "КА") > 0 Or InStr(pi.Name, "Кредитовое авизо") > 0 Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Фильтруем поле "Категория просрочки", оставляя только нужные значения
    Set pf = PivotTable.PivotFields("Категория просрочки")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If pi.Name <> "просрочка более 60 дней" And _
           pi.Name <> "просрочка от 30 до 60 дней" And _
           pi.Name <> "просрочка от 15 до 30 дней" And _
           pi.Name <> "просрочка до 15 дней" Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Фильтруем поле "Тип СФ", убирая значения, содержащие "Факторинг" и "Суды и прочие"
    Set pf = PivotTable.PivotFields("Сегмент")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "Факторинг") > 0 Or InStr(pi.Name, "Суды и прочие") > 0 Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Сворачиваем все элементы по столбцу "Заказчик"
    Set pf = PivotTable.PivotFields("Заказчик")
    pf.ShowDetail = False
    
    ' Создаем новый лист с именем "Анализ ПДЗ"
    On Error Resume Next ' Игнорируем ошибку, если лист уже существует
    Set NewSheet = ActiveWorkbook.Sheets("Анализ ПДЗ")
    On Error GoTo 0

    If NewSheet Is Nothing Then
        Set NewSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        NewSheet.Name = "Анализ ПДЗ"
    Else
        ' Если лист уже существует, очищаем его
        NewSheet.Cells.Clear
    End If

    ' Определяем диапазон сводной таблицы
    Set pivotTableRange = PivotTable.TableRange2
    
    ' Копируем сводную таблицу с листа "Свод" на лист "Анализ ПДЗ" как значения и приводим к общему формату
    pivotTableRange.Copy
    NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues ' Вставляем как значения
    NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats ' Вставляем форматирование
    Application.CutCopyMode = False ' Очищаем буфер обмена

    With Worksheets("Анализ ПДЗ")
        .Rows(1).Delete
    End With

    With Worksheets("Анализ ПДЗ")
        .Range("A1:C1").Value = .Range("A2:C2").Value ' Перенос данных из A2:C2 в A1:C1
        .Rows(2).Delete ' Удаление второй строки
    End With

    With Worksheets("Анализ ПДЗ").Rows(1)
        .WrapText = True ' Перенос текста
        .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
        .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
        .Font.Bold = True ' Жирный шрифт
        .Font.Color = RGB(0, 0, 0) ' Черный цвет шрифта
    End With
    
    ' Удаляем столбец D
    NewSheet.Columns("D").Delete

    ' Устанавливаем ширину всех столбцов на активном листе в соответствии с длиной текста
    ActiveSheet.Columns.AutoFit
    
    ' Удаляем строки, где значения в столбце H меньше 10
    lastRow = NewSheet.Cells(NewSheet.Rows.Count, "H").End(xlUp).Row
    For i = lastRow To 1 Step -1
        If IsNumeric(NewSheet.Cells(i, "H").Value) Then
            If NewSheet.Cells(i, "H").Value < 10 Then
                NewSheet.Rows(i).Delete
            End If
        End If
    Next i
    
    ' Найдите последнюю заполненную строку в столбцах D:G
    lastRow = Cells(Rows.Count, "D").End(xlUp).Row
    Dim lastRowH As Long
    lastRowH = Cells(Rows.Count, "G").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRow, lastRowH)

    ' Вставьте формулу суммирования в столбец H, начиная со второй строки
    For i = 2 To lastRow
        Cells(i, "H").Formula = "=SUM(D" & i & ":G" & i & ")"
    Next i
    
    ' Вставьте формулы суммирования в столбцы D:G, начиная со второй строки
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки
    Do While currentRow <= lastRow
        currentValue = Cells(currentRow, 1).Value
        
        ' Поиск строк с одинаковыми значениями в столбце A до строки с "Итог"
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Определение начала диапазона для суммирования
            nextRow = totalRow - 1
            Do While nextRow >= 2 And Cells(nextRow, 1).Value = Cells(nextRow - 1, 1).Value
                nextRow = nextRow - 1
            Loop
            
            ' Вставка формулы для суммирования в строку с "Итог"
            Cells(totalRow, 4).Formula = "=SUM(D" & nextRow & ":D" & totalRow - 1 & ")"
            Cells(totalRow, 5).Formula = "=SUM(E" & nextRow & ":E" & totalRow - 1 & ")"
            Cells(totalRow, 6).Formula = "=SUM(F" & nextRow & ":F" & totalRow - 1 & ")"
            Cells(totalRow, 7).Formula = "=SUM(G" & nextRow & ":G" & totalRow - 1 & ")"
        End If
        
        currentRow = currentRow + 1
    Loop

    ' Теперь добавим формулу суммирования для строки с "Общий итог"
    sumRowsD = ""
    sumRowsE = ""
    sumRowsF = ""
    sumRowsG = ""

    For r = 2 To lastRow
        If InStr(1, Cells(r, "A").Value, "итог", vbTextCompare) > 0 Then
            If InStr(1, Cells(r, "A").Value, "общий итог", vbTextCompare) = 0 Then
                If sumRowsD = "" Then
                    sumRowsD = "D" & r
                    sumRowsE = "E" & r
                    sumRowsF = "F" & r
                    sumRowsG = "G" & r
                Else
                    sumRowsD = sumRowsD & ",D" & r
                    sumRowsE = sumRowsE & ",E" & r
                    sumRowsF = sumRowsF & ",F" & r
                    sumRowsG = sumRowsG & ",G" & r
                End If
            End If
        End If
    Next r

    ' Вставка итоговых формул в строку после последней найденной строки
    Cells(lastRow + 1, 1).Value = "Общий итог"
    If sumRowsD <> "" Then Cells(lastRow + 1, 4).Formula = "=SUM(" & sumRowsD & ")"
    If sumRowsE <> "" Then Cells(lastRow + 1, 5).Formula = "=SUM(" & sumRowsE & ")"
    If sumRowsF <> "" Then Cells(lastRow + 1, 6).Formula = "=SUM(" & sumRowsF & ")"
    If sumRowsG <> "" Then Cells(lastRow + 1, 7).Formula = "=SUM(" & sumRowsG & ")"

End Sub










    Dim PivotTable As PivotTable
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim pivotTableRange As Range
    Dim sumRowsD As String, sumRowsE As String, sumRowsF As String, sumRowsG As String

    ' Определяем лист и сводную таблицу
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")
    
    ' Добавляем новое поле "Тип СФ" в строки
    With PivotTable
        .PivotFields("Тип СФ").orientation = xlRowField
        .PivotFields("Тип СФ").position = 4 ' Устанавливаем в четвёртую позицию в строках
    End With
    
    ' Обновляем сводную таблицу
    PivotTable.RefreshTable
    
    ' Фильтруем поле "Тип СФ", убирая значения, содержащие "КА" и "Кредитовое авизо"
    Set pf = PivotTable.PivotFields("Тип СФ")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "КА") > 0 Or InStr(pi.Name, "Кредитовое авизо") > 0 Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Фильтруем поле "Категория просрочки", оставляя только нужные значения
    Set pf = PivotTable.PivotFields("Категория просрочки")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If pi.Name <> "просрочка более 60 дней" And _
           pi.Name <> "просрочка от 30 до 60 дней" And _
           pi.Name <> "просрочка от 15 до 30 дней" And _
           pi.Name <> "просрочка до 15 дней" Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
     ' Фильтруем поле "Тип СФ", убирая значения, содержащие "Факторинг" и "Суды и прочие"
    Set pf = PivotTable.PivotFields("Сегмент")
    On Error Resume Next ' Игнорируем ошибку, если элемент уже скрыт
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "Факторинг") > 0 Or InStr(pi.Name, "Суды и прочие") > 0 Then
            pi.Visible = False
        End If
    Next pi
    On Error GoTo 0 ' Включаем обработку ошибок обратно
    
    ' Сворачиваем все элементы по столбцу "Заказчик"
    Set pf = PivotTable.PivotFields("Заказчик")
    pf.ShowDetail = False
    
    ' Создаем новый лист с именем "Анализ ПДЗ"
    On Error Resume Next ' Игнорируем ошибку, если лист уже существует
    Set NewSheet = ActiveWorkbook.Sheets("Анализ ПДЗ")
    On Error GoTo 0

    If NewSheet Is Nothing Then
        Set NewSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        NewSheet.Name = "Анализ ПДЗ"
    Else
        ' Если лист уже существует, очищаем его
        NewSheet.Cells.Clear
    End If

    ' Определяем диапазон сводной таблицы
    Set pivotTableRange = PivotTable.TableRange2
    
     ' Копируем сводную таблицу с листа "Свод" на лист "Анализ ПДЗ" как значения и приводим к общему формату
    pivotTableRange.Copy
    NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues ' Вставляем как значения
    NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats ' Вставляем форматирование
    Application.CutCopyMode = False ' Очищаем буфер обмена

    With Worksheets("Анализ ПДЗ")
    .Rows(1).Delete
End With

With Worksheets("Анализ ПДЗ")
    .Range("A1:C1").Value = .Range("A2:C2").Value ' Перенос данных из A2:C2 в A1:C1
    .Rows(2).Delete ' Удаление второй строки
End With

With Worksheets("Анализ ПДЗ").Rows(1)
    .WrapText = True ' Перенос текста
    .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
    .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
    .Font.Bold = True ' Жирный шрифт
    .Font.Color = RGB(0, 0, 0) ' Черный цвет шрифта
End With
    
  ' Удаляем столбец D
    NewSheet.Columns("D").Delete

     ' Устанавливаем ширину всех столбцов на активном листе в соответствии с длиной текста
    ActiveSheet.Columns.AutoFit
    
    ' Удаляем строки, где значения в столбце H меньше 10
    lastRow = NewSheet.Cells(NewSheet.Rows.Count, "H").End(xlUp).Row
    For i = lastRow To 1 Step -1
        If IsNumeric(NewSheet.Cells(i, "H").Value) Then
            If NewSheet.Cells(i, "H").Value < 10 Then
                NewSheet.Rows(i).Delete
            End If
        End If
    Next i
    
  ' Найдите последнюю заполненную строку в столбцах D:G
    lastRow = Cells(Rows.Count, "D").End(xlUp).Row
    Dim lastRowH As Long
    lastRowH = Cells(Rows.Count, "G").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRow, lastRowH)

    ' Вставьте формулу суммирования в столбец H, начиная со второй строки
    For i = 2 To lastRow
        Cells(i, "H").formula = "=SUM(D" & i & ":G" & i & ")"
    Next i
    
' Вставьте формулу суммирования в столбцы D:G, начиная со второй строки
    Dim currentRow As Long
    Dim nextRow As Long
    Dim currentValue As String
    Dim totalRow As Long
    
    lastRow = Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки в столбце A
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки
    
    Do While currentRow <= lastRow
        currentValue = Cells(currentRow, 1).Value
        
        ' Поиск строк с одинаковыми значениями в столбце A до строки с "Итог"
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Определение начала диапазона для суммирования
            nextRow = totalRow - 1
            Do While nextRow >= 2 And Cells(nextRow, 1).Value = Cells(nextRow - 1, 1).Value
                nextRow = nextRow - 1
            Loop
            
            ' Вставка формулы для суммирования в строку с "Итог"
            Cells(totalRow, 4).formula = "=SUM(D" & nextRow & ":D" & totalRow - 1 & ")"
            Cells(totalRow, 5).formula = "=SUM(E" & nextRow & ":E" & totalRow - 1 & ")"
            Cells(totalRow, 6).formula = "=SUM(F" & nextRow & ":F" & totalRow - 1 & ")"
            Cells(totalRow, 7).formula = "=SUM(G" & nextRow & ":G" & totalRow - 1 & ")"
        End If
        
        currentRow = currentRow + 1
    Loop
    

' Теперь добавим формулу суммирования для строки с "Общий итог"
sumRowsD = ""
sumRowsE = ""
sumRowsF = ""
sumRowsG = ""

For r = 2 To lastRow
    If InStr(1, Cells(r, "A").Value, "итог", vbTextCompare) > 0 Then
        If InStr(1, Cells(r, "A").Value, "общий итог", vbTextCompare) = 0 Then
            If sumRowsD = "" Then
                sumRowsD = "D" & r
                sumRowsE = "E" & r
                sumRowsF = "F" & r
                sumRowsG = "G" & r
            Else
                sumRowsD = sumRowsD & ",D" & r
                sumRowsE = sumRowsE & ",E" & r
                sumRowsF = sumRowsF & ",F" & r
                sumRowsG = sumRowsG & ",G" & r
            End If
        End If
    End If
Next r

' Вставьте формулы суммирования в строку, где есть "Общий итог"
For r = 2 To lastRow
    If InStr(1, Cells(r, "A").Value, "Общий итог", vbTextCompare) > 0 Then
        If sumRowsD <> "" Then
            Cells(r, "D").formula = "=SUM(" & sumRowsD & ")"
            Cells(r, "E").formula = "=SUM(" & sumRowsE & ")"
            Cells(r, "F").formula = "=SUM(" & sumRowsF & ")"
            Cells(r, "G").formula = "=SUM(" & sumRowsG & ")"
        End If
        Exit For ' Выход из цикла после нахождения первой строки с "Общий итог"
    End If
Next r

    
    ' Очищаем буфер обмена
    Application.CutCopyMode = False
    
End Sub











' Теперь добавим формулу суммирования для строки с "Общий итог"
Dim sumRowsD As String, sumRowsE As String, sumRowsF As String, sumRowsG As String
sumRowsD = ""
sumRowsE = ""
sumRowsF = ""
sumRowsG = ""

For r = 2 To lastRow
    If InStr(1, Cells(r, "A").Value, "итог", vbTextCompare) > 0 Then
        If InStr(1, Cells(r, "A").Value, "общий итог", vbTextCompare) = 0 Then
            If sumRowsD = "" Then
                sumRowsD = "D" & r
                sumRowsE = "E" & r
                sumRowsF = "F" & r
                sumRowsG = "G" & r
            Else
                sumRowsD = sumRowsD & ",D" & r
                sumRowsE = sumRowsE & ",E" & r
                sumRowsF = sumRowsF & ",F" & r
                sumRowsG = sumRowsG & ",G" & r
            End If
        End If
    End If
Next r

' Вставьте формулы суммирования в строку, где есть "Общий итог"
For r = 2 To lastRow
    If InStr(1, Cells(r, "A").Value, "Общий итог", vbTextCompare) > 0 Then
        If sumRowsD <> "" Then
            Cells(r, "D").Formula = "=SUM(" & sumRowsD & ")"
            Cells(r, "E").Formula = "=SUM(" & sumRowsE & ")"
            Cells(r, "F").Formula = "=SUM(" & sumRowsF & ")"
            Cells(r, "G").Formula = "=SUM(" & sumRowsG & ")"
        End If
        Exit For ' Выход из цикла после нахождения первой строки с "Общий итог"
    End If
Next r




Sub SummarizeBetweenTotalsAndOverallTotal()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextTotalRow As Long
    Dim sumStartRow As Long
    Dim sumRowsD As String, sumRowsE As String, sumRowsF As String, sumRowsG As String
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажите название вашего листа
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Последняя заполненная строка
    
    currentRow = lastRow
    sumRowsD = ""
    sumRowsE = ""
    sumRowsF = ""
    sumRowsG = ""
    
    ' Идём от последней строки к первой
    Do While currentRow >= 2
        ' Если находим строку с "Итог" в столбце A
        If InStr(1, ws.Cells(currentRow, 1).Value, "Итог", vbTextCompare) > 0 Then
            ' Устанавливаем границы диапазона для суммирования
            sumStartRow = currentRow + 1 ' Начало диапазона после строки "Итог"
            nextTotalRow = currentRow - 1
            
            ' Ищем следующую строку с "Итог" выше
            Do While nextTotalRow >= 2
                If InStr(1, ws.Cells(nextTotalRow, 1).Value, "Итог", vbTextCompare) > 0 Then
                    Exit Do ' Выходим, когда находим следующий "Итог"
                End If
                nextTotalRow = nextTotalRow - 1
            Loop
            
            ' Вставляем формулу суммирования в строку с "Итог" для столбцов D:G
            ws.Cells(currentRow, 4).Formula = "=SUM(D" & nextTotalRow + 1 & ":D" & sumStartRow - 1 & ")"
            ws.Cells(currentRow, 5).Formula = "=SUM(E" & nextTotalRow + 1 & ":E" & sumStartRow - 1 & ")"
            ws.Cells(currentRow, 6).Formula = "=SUM(F" & nextTotalRow + 1 & ":F" & sumStartRow - 1 & ")"
            ws.Cells(currentRow, 7).Formula = "=SUM(G" & nextTotalRow + 1 & ":G" & sumStartRow - 1 & ")"
            
            ' Добавляем строки с "Итог" в список для общего суммирования
            If sumRowsD = "" Then
                sumRowsD = "D" & currentRow
                sumRowsE = "E" & currentRow
                sumRowsF = "F" & currentRow
                sumRowsG = "G" & currentRow
            Else
                sumRowsD = sumRowsD & ",D" & currentRow
                sumRowsE = sumRowsE & ",E" & currentRow
                sumRowsF = sumRowsF & ",F" & currentRow
                sumRowsG = sumRowsG & ",G" & currentRow
            End If
        End If
        
        currentRow = currentRow - 1 ' Переходим к следующей строке вверх
    Loop
    
    ' Теперь добавим формулу суммирования для строки с "Общий итог"
    For r = 2 To lastRow
        If InStr(1, ws.Cells(r, "A").Value, "Общий итог", vbTextCompare) > 0 Then
            If sumRowsD <> "" Then
                ws.Cells(r, "D").Formula = "=SUM(" & sumRowsD & ")"
                ws.Cells(r, "E").Formula = "=SUM(" & sumRowsE & ")"
                ws.Cells(r, "F").Formula = "=SUM(" & sumRowsF & ")"
                ws.Cells(r, "G").Formula = "=SUM(" & sumRowsG & ")"
            End If
            Exit For ' Выход из цикла после нахождения первой строки с "Общий итог"
        End If
    Next r
End Sub





' Теперь добавим формулу суммирования для строки с "Общий итог"
    Dim sumRows As String
    sumRows = ""

    For r = 2 To lastRow
        If InStr(1, Cells(r, "A").Value, "итог", vbTextCompare) > 0 Then
            If InStr(1, Cells(r, "A").Value, "общий итог", vbTextCompare) = 0 Then
                If sumRows = "" Then
                    sumRows = "D" & r
                Else
                    sumRows = sumRows & ",D" & r
                End If
            End If
        End If
    Next r
    
    ' Вставьте формулу суммирования в строку, где есть "Общий итог"
    For r = 2 To lastRow
        If InStr(1, Cells(r, "A").Value, "Общий итог", vbTextCompare) > 0 Then
            If sumRows <> "" Then
                Cells(r, "D").formula = "=SUM(" & sumRows & ")"
            End If
            Exit For ' Выход из цикла после нахождения первой строки с "Общий итог"
        End If
    Next r





Sub SummarizeBetweenTotals()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextTotalRow As Long
    Dim sumStartRow As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажите название вашего листа
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Последняя заполненная строка
    
    currentRow = lastRow
    
    ' Идём от последней строки к первой
    Do While currentRow >= 2
        ' Если находим строку с "Итог" в столбце A
        If InStr(1, ws.Cells(currentRow, 1).Value, "Итог", vbTextCompare) > 0 Then
            ' Устанавливаем границы диапазона для суммирования
            sumStartRow = currentRow + 1 ' Начало диапазона после строки "Итог"
            nextTotalRow = currentRow - 1
            
            ' Ищем следующую строку с "Итог" выше
            Do While nextTotalRow >= 2
                If InStr(1, ws.Cells(nextTotalRow, 1).Value, "Итог", vbTextCompare) > 0 Then
                    Exit Do ' Выходим, когда находим следующий "Итог"
                End If
                nextTotalRow = nextTotalRow - 1
            Loop
            
            ' Вставляем формулу суммирования в строку с "Итог"
            ws.Cells(currentRow, 4).Formula = "=SUM(D" & nextTotalRow + 1 & ":D" & sumStartRow - 1 & ")"
            ws.Cells(currentRow, 5).Formula = "=SUM(E" & nextTotalRow + 1 & ":E" & sumStartRow - 1 & ")"
            ws.Cells(currentRow, 6).Formula = "=SUM(F" & nextTotalRow + 1 & ":F" & sumStartRow - 1 & ")"
            ws.Cells(currentRow, 7).Formula = "=SUM(G" & nextTotalRow + 1 & ":G" & sumStartRow - 1 & ")"
        End If
        
        currentRow = currentRow - 1 ' Переходим к следующей строке вверх
    Loop
End Sub




Sub SummarizeByMatchingNamesWithoutFormulaFixed()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextRow As Long
    Dim currentValue As String
    Dim totalRow As Long
    Dim sumD As Double, sumE As Double, sumF As Double, sumG As Double
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Название листа с данными
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки данных
    currentRow = 2 ' Первая строка с данными (предполагаем, что первая строка - заголовок)
    
    Application.ScreenUpdating = False ' Отключение обновления экрана для повышения производительности
    
    Do While currentRow <= lastRow
        currentValue = ws.Cells(currentRow, 1).Value
        
        ' Поиск строки с "Итог" в столбце A
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Инициализация переменных для суммирования
            sumD = 0
            sumE = 0
            sumF = 0
            sumG = 0
            
            ' Поиск диапазона для суммирования (до строки с "Итог")
            nextRow = totalRow - 1
            Do While nextRow >= 2 And ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value
                sumD = sumD + ws.Cells(nextRow, 4).Value
                sumE = sumE + ws.Cells(nextRow, 5).Value
                sumF = sumF + ws.Cells(nextRow, 6).Value
                sumG = sumG + ws.Cells(nextRow, 7).Value
                nextRow = nextRow - 1
            Loop
            
            ' Добавление суммы в строку с "Итог" в столбцы D:G
            sumD = sumD + ws.Cells(nextRow, 4).Value
            sumE = sumE + ws.Cells(nextRow, 5).Value
            sumF = sumF + ws.Cells(nextRow, 6).Value
            sumG = sumG + ws.Cells(nextRow, 7).Value
            
            ws.Cells(totalRow, 4).Value = sumD
            ws.Cells(totalRow, 5).Value = sumE
            ws.Cells(totalRow, 6).Value = sumF
            ws.Cells(totalRow, 7).Value = sumG
        End If
        
        currentRow = currentRow + 1
    Loop
    
    Application.ScreenUpdating = True ' Включение обновления экрана
End Sub



Sub SummarizeByMatchingNamesWithoutFormula()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextRow As Long
    Dim currentValue As String
    Dim totalRow As Long
    Dim sumD As Double, sumE As Double, sumF As Double, sumG As Double
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажите название вашего листа
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки в столбце A
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки
    
    Do While currentRow <= lastRow
        currentValue = ws.Cells(currentRow, 1).Value
        
        ' Поиск строки с "Итог" в столбце A
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Инициализация переменных для суммирования
            sumD = 0
            sumE = 0
            sumF = 0
            sumG = 0
            
            ' Поиск диапазона для суммирования (до строки с "Итог")
            nextRow = totalRow - 1
            Do While nextRow >= 2 And ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value
                sumD = sumD + ws.Cells(nextRow, 4).Value
                sumE = sumE + ws.Cells(nextRow, 5).Value
                sumF = sumF + ws.Cells(nextRow, 6).Value
                sumG = sumG + ws.Cells(nextRow, 7).Value
                nextRow = nextRow - 1
            Loop
            
            ' Добавление суммы в строку с "Итог" в столбцы D:G
            sumD = sumD + ws.Cells(nextRow, 4).Value
            sumE = sumE + ws.Cells(nextRow, 5).Value
            sumF = sumF + ws.Cells(nextRow, 6).Value
            sumG = sumG + ws.Cells(nextRow, 7).Value
            
            ws.Cells(totalRow, 4).Value = sumD
            ws.Cells(totalRow, 5).Value = sumE
            ws.Cells(totalRow, 6).Value = sumF
            ws.Cells(totalRow, 7).Value = sumG
        End If
        
        currentRow = currentRow + 1
    Loop
End Sub


Sub SummarizeByMatchingNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextRow As Long
    Dim currentValue As String
    Dim totalRow As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажите название вашего листа
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки в столбце A
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки
    
    Do While currentRow <= lastRow
        currentValue = ws.Cells(currentRow, 1).Value
        
        ' Поиск строк с одинаковыми значениями в столбце A до строки с "Итог"
        If InStr(1, currentValue, "Итог", vbTextCompare) > 0 Then
            totalRow = currentRow ' Строка с "Итог"
            
            ' Определение начала диапазона для суммирования
            nextRow = totalRow - 1
            Do While nextRow >= 2 And ws.Cells(nextRow, 1).Value = ws.Cells(nextRow - 1, 1).Value
                nextRow = nextRow - 1
            Loop
            
            ' Вставка формулы для суммирования в строку с "Итог"
            ws.Cells(totalRow, 4).Formula = "=SUM(D" & nextRow & ":D" & totalRow - 1 & ")"
            ws.Cells(totalRow, 5).Formula = "=SUM(E" & nextRow & ":E" & totalRow - 1 & ")"
            ws.Cells(totalRow, 6).Formula = "=SUM(F" & nextRow & ":F" & totalRow - 1 & ")"
            ws.Cells(totalRow, 7).Formula = "=SUM(G" & nextRow & ":G" & totalRow - 1 & ")"
        End If
        
        currentRow = currentRow + 1
    Loop
End Sub




Sub SummarizeByColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim nextRow As Long
    Dim currentValue As String

    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажите название вашего листа
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Определение последней строки в столбце A
    currentRow = 2 ' Предполагается, что первая строка содержит заголовки

    Do While currentRow <= lastRow
        currentValue = ws.Cells(currentRow, 1).Value
        nextRow = currentRow + 1
        
        ' Поиск одинаковых значений в столбце A
        Do While nextRow <= lastRow And ws.Cells(nextRow, 1).Value = currentValue
            nextRow = nextRow + 1
        Loop
        
        ' Вставка формулы для суммирования диапазонов D:G
        ws.Cells(nextRow, 1).Value = currentValue
        ws.Cells(nextRow, 4).Formula = "=SUM(D" & currentRow & ":D" & nextRow - 1 & ")"
        ws.Cells(nextRow, 5).Formula = "=SUM(E" & currentRow & ":E" & nextRow - 1 & ")"
        ws.Cells(nextRow, 6).Formula = "=SUM(F" & currentRow & ":F" & nextRow - 1 & ")"
        ws.Cells(nextRow, 7).Formula = "=SUM(G" & currentRow & ":G" & nextRow - 1 & ")"
        
        currentRow = nextRow + 1
    Loop
End Sub




Sub СоздатьЛистИКопироватьСводную()
    Dim PivotSheet As Worksheet
    Dim AnalysisSheet As Worksheet
    Dim ParamSheet As Worksheet
    Dim LastRow As Long, LastCol As Long
    Dim PivotRange As Range

    ' Определяем лист "Свод" и лист "Параметры"
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    Set ParamSheet = ActiveWorkbook.Sheets("Параметры")
    
    ' Создаём новый лист "Анализ ПДЗ" и вставляем его после листа "Параметры"
    Set AnalysisSheet = ActiveWorkbook.Sheets.Add(After:=ParamSheet)
    AnalysisSheet.Name = "Анализ ПДЗ"
    
    ' Находим последнюю заполненную строку и столбец на листе "Свод"
    With PivotSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set PivotRange = .Range(.Cells(1, 1), .Cells(LastRow, LastCol))
    End With

    ' Копируем диапазон сводной таблицы на листе "Свод"
    PivotRange.Copy
    
    ' Вставляем значения и форматы на лист "Анализ ПДЗ"
    With AnalysisSheet.Cells(1, 1)
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
    End With
    
    ' Очищаем буфер обмена
    Application.CutCopyMode = False
    
    ' Автоширина для всех столбцов на листе "Анализ ПДЗ"
    AnalysisSheet.Columns.AutoFit
End Sub





Sub ДобавитьФильтрыИСворачивания()
    Dim PivotTable As PivotTable
    Dim ws As Worksheet
    Dim pf As PivotField
    Dim pi As PivotItem

    ' Определяем лист и сводную таблицу
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")
    
    ' Добавляем новое поле "Тип СФ" в строки
    With PivotTable
        .PivotFields("Тип СФ").Orientation = xlRowField
        .PivotFields("Тип СФ").Position = 1 ' Устанавливаем первым полем в строках
    End With
    
    ' Обновляем сводную таблицу
    PivotTable.RefreshTable
    
    ' Фильтруем поле "Тип СФ", убирая значения, содержащие "КА" и "Кредитовое авизо"
    Set pf = PivotTable.PivotFields("Тип СФ")
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "КА") > 0 Or InStr(pi.Name, "Кредитовое авизо") > 0 Then
            pi.Visible = False
        End If
    Next pi
    
    ' Фильтруем поле "Категория просрочки", оставляя только нужные значения
    Set pf = PivotTable.PivotFields("Категория просрочки")
    For Each pi In pf.PivotItems
        If pi.Name <> "просрочка более 60 дней" And pi.Name <> "просрочка от 30 до 60 дней" And _
           pi.Name <> "просрочка от 15 до 30 дней" And pi.Name <> "просрочка до 15 дней" Then
            pi.Visible = False
        End If
    Next pi
    
    ' Сворачиваем все элементы по столбцу "Заказчик"
    Set pf = PivotTable.PivotFields("Заказчик")
    pf.ShowDetail = False
    
End Sub




Sub ДобавитьТипСФВСводСФильтрацией()
    Dim PivotTable As PivotTable
    Dim ws As Worksheet
    Dim pf As PivotField
    Dim pi As PivotItem
    
    ' Определяем лист и сводную таблицу
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")
    
    ' Добавляем новое поле "Тип СФ" в строки
    With PivotTable
        .PivotFields("Тип СФ").Orientation = xlRowField
        .PivotFields("Тип СФ").Position = 1 ' Устанавливаем первым полем в строках
    End With
    
    ' Обновляем сводную таблицу
    PivotTable.RefreshTable
    
    ' Получаем поле "Тип СФ"
    Set pf = PivotTable.PivotFields("Тип СФ")
    
    ' Убираем все элементы, содержащие "КА" и "Кредитовое авизо"
    For Each pi In pf.PivotItems
        If InStr(pi.Name, "КА") > 0 Or InStr(pi.Name, "Кредитовое авизо") > 0 Then
            pi.Visible = False
        End If
    Next pi
End Sub





Sub ДобавитьТипСФВСвод()
    Dim PivotTable As PivotTable
    Dim ws As Worksheet

    ' Определяем лист и сводную таблицу
    Set ws = ActiveWorkbook.Sheets("Свод")
    Set PivotTable = ws.PivotTables("СводнаяТаблица")

    ' Добавляем новое поле "Тип СФ" в строки
    With PivotTable
        .PivotFields("Тип СФ").Orientation = xlRowField
        .PivotFields("Тип СФ").Position = 1 ' Устанавливаем первым полем в строках
    End With

    ' Обновляем сводную таблицу
    PivotTable.RefreshTable
End Sub















w, j), ws.Cells(i - 1, j))
            
            ' Проверяем, что диапазон не пустой
            If WorksheetFunction.CountA(sumRange) > 0 Then
                ' Вставляем формулу в ячейку
                ws.Cells(i, j).Formula = "=SUM(" & sumRange.Address(False, False) & ")"
            End If
        Next j
        ' После строки с "Итог", обновляем startRow для следующего блока данных
        startRow = i + 1
    End If
Next i



' Определяем последнюю строку и столбец на листе
Dim lastRow As Long, lastCol As Long
Dim i As Long, startRow As Long
Dim sumRange As Range

' Определяем последнюю строку и последний столбец на листе
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

' Переменная для хранения строки начала объединенного диапазона
startRow = 2 ' Первая строка данных (вторая строка листа)

' Проходим по каждой строке от 2 до последней
For i = 2 To lastRow
    ' Проверяем наличие "Итог" в столбце A
    If InStr(ws.Cells(i, "A").Value, "Итог") > 0 Then
        ' Для каждого столбца от D до последнего столбца
        For j = 4 To lastCol ' Столбец D - это 4-й столбец
            ' Определяем диапазон для суммирования от startRow до строки перед "Итог"
            Set sumRange = ws.Range(ws.Cells(startRow, j), ws.Cells(i - 1, j))
            
            ' Проверяем, что диапазон не пустой
            If WorksheetFunction.CountA(sumRange) > 0 Then
                ' Вставляем формулу в ячейку
                ws.Cells(i, j).Formula = "=SUM(" & sumRange.Address(False, False) & ")"
            End If
        Next j
        ' После строки с "Итог", обновляем startRow для следующего блока данных
        startRow = i + 1
    End If
Next i




' Добавляем формулы суммирования в строки, где в столбце A указано "Итог"
Dim lastRow As Long
Dim lastCol As Long
Dim i As Long
Dim startRow As Long
Dim sumRange As Range
Dim cell As Range

' Определяем последнюю заполненную строку и столбец на листе "Ю.В."
With ws
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
End With

' Переменная для хранения строки начала объединенной области
startRow = 2

' Проходим по всем строкам, начиная со 2-й
For i = 2 To lastRow
    ' Если в столбце A указано "Итог" (частичное или полное совпадение)
    If InStr(ws.Cells(i, "A").Value, "Итог") > 0 Then
        ' Проходим по каждому столбцу от D до последнего
        For Each cell In ws.Range(ws.Cells(i, "D"), ws.Cells(i, lastCol))
            ' Определяем диапазон для суммирования на основе объединенной ячейки выше
            Set sumRange = ws.Range(ws.Cells(startRow, cell.Column), ws.Cells(i - 1, cell.Column))
            ' Вставляем формулу суммирования
            cell.Formula = "=SUM(" & sumRange.Address(False, False) & ")"
        Next cell
    End If
    
    ' Проверяем, объединена ли ячейка в столбце A
    If ws.Cells(i, "A").MergeCells Then
        ' Если да, обновляем startRow на первую строку объединенной ячейки
        startRow = ws.Cells(i, "A").MergeArea.Row
    End If
Next i



Sub ВставитьФормулуСуммирования()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim sumRange As Range
    Dim cell As Range

    ' Определяем лист "Ю.В."
    Set ws = Worksheets("Ю.В.")

    ' Определяем последнюю заполненную строку и столбец
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Проходим по всем строкам, начиная со 2-й
    For i = 2 To lastRow
        ' Если в столбце A указано "Итог"
        If ws.Cells(i, "A").Value = "Итог" Then
            ' Суммируем значения начиная с D до последнего столбца
            For Each cell In ws.Range(ws.Cells(i, "D"), ws.Cells(i, lastCol))
                ' Определяем диапазон для суммирования
                Set sumRange = ws.Range(ws.Cells(2, cell.Column), ws.Cells(i - 1, cell.Column))
                ' Вставляем формулу суммирования
                cell.Formula = "=SUM(" & sumRange.Address(False, False) & ")"
            Next cell
        End If
    Next i
End Sub




Sub ОбъединитьСтроки()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Определяем рабочий лист "Ю.В."
    Set ws = Worksheets("Ю.В.")
    
    ' Определяем последнюю строку в столбце A (предположим, что там нет пустых строк)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Цикл по столбцу "Сегмент" для объединения одинаковых значений
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
            ws.Range(ws.Cells(i - 1, 1), ws.Cells(i, 1)).Merge
        End If
    Next i
    
    ' Цикл по столбцу "Ответственный" для объединения одинаковых значений
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 2).Value = ws.Cells(i - 1, 2).Value Then
            ws.Range(ws.Cells(i - 1, 2), ws.Cells(i, 2)).Merge
        End If
    Next i
End Sub




' Фильтруем пустые строки и строки с текстом "оптовик частный" в столбце "Сегмент"
Dim FilterRange As Range
Dim LastRow As Long

' Находим последний заполненный ряд
LastRow = DataSheet.Cells(DataSheet.Rows.Count, 1).End(xlUp).Row

' Устанавливаем диапазон фильтрации
Set FilterRange = DataSheet.Range("A1").CurrentRegion

' Проверяем, существует ли фильтр, и очищаем его
If DataSheet.AutoFilterMode Then
    DataSheet.AutoFilterMode = False
End If

' Применяем автофильтр
With FilterRange
    .AutoFilter Field:=2, Criteria1:="<>", Operator:=xlAnd ' Убираем пустые строки в столбце "Сегмент"
    .AutoFilter Field:=2, Criteria1:="<>оптовик частный" ' Убираем "оптовик частный"
End With

' Копируем данные на лист "Ю.В." как значения
Dim CopyRange As Range
Set CopyRange = FilterRange.SpecialCells(xlCellTypeVisible)

' Проверяем, есть ли данные для копирования
If CopyRange.Rows.Count > 1 Then ' Исключаем заголовок
    CopyRange.Copy
    With Worksheets("Ю.В.")
        .Range("A1").PasteSpecial Paste:=xlPasteValues
    End With
End If

' Сбрасываем фильтр
If DataSheet.AutoFilterMode Then
    DataSheet.AutoFilterMode = False
End If




' Фильтруем пустые строки и строки с текстом "оптовик частный" в столбце "Сегмент"
Dim FilterRange As Range
Set FilterRange = DataSheet.Range("A1").CurrentRegion ' Установите диапазон фильтрации

With FilterRange
    ' Применяем автофильтр
    .AutoFilter Field:=2, Criteria1:="<>", Operator:=xlAnd ' Убираем пустые строки в столбце "Сегмент"
    .AutoFilter Field:=2, Criteria1:="<>оптовик частный", Operator:=xlAnd ' Убираем "оптовик частный"
End With

' Теперь можно скопировать данные на лист "Ю.В." как значения
' Здесь добавьте ваш код для копирования данных

' После переноса данных сбрасываем фильтр
If FilterRange.Parent.AutoFilterMode Then
    FilterRange.AutoFilter
End If



' Фильтруем пустые строки и строки с текстом "оптовик частный" в столбце "Сегмент"
Dim FilterRange As Range
Set FilterRange = DataSheet.Range("A1").CurrentRegion ' Установите диапазон фильтрации

With FilterRange
    .AutoFilter Field:=1, Criteria1:="<>" ' Убираем пустые строки в первом столбце
    .AutoFilter Field:=1, Criteria1:="<>оптовик частный", Operator:=xlAnd ' Убираем "оптовик частный"
End With



' Включаем повторение подписей элементов только для столбцов "Сегмент" и "Ответственный"
With PivotTable
    .PivotFields("Сегмент").RepeatLabels = True
    .PivotFields("Ответственный").RepeatLabels = True
End With



' Включаем повторение подписей элементов для строк и столбцов
With PivotTable
    .RowGrand = True
    .ColumnGrand = True
    .RepeatAllLabels = True
End With




With Worksheets("Ю.В").UsedRange
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
End With


With Worksheets("Ю.В.").Rows(1)
    .WrapText = True ' Перенос текста
    .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
    .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
    .Font.Bold = True ' Полужирный шрифт
    .Font.Color = RGB(0, 0, 0) ' Черный цвет шрифта
End With



юWith Worksheets("Ю.В.").Rows(1)
    .WrapText = True ' Перенос текста
    .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
    .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
    .Font.Bold = True ' Полужирный шрифт
End With


юWith Worksheets("Ю.В.").Rows(1)
    .WrapText = True ' Перенос текста
    .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
    .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
End With



With Worksheets("Ю.В.")
    .Range("A1:C1").Value = .Range("A2:C2").Value ' Перенос данных из A2:C2 в A1:C1
    .Rows(2).Delete ' Удаление второй строки
End With


With Worksheets("Ю.В.")
    .Rows(1).Value = .Rows(2).Value ' Перенос данных со второй строки в первую
    .Rows(2).Delete ' Удаление второй строки
End With



With Worksheets("Ю.В.")
    .Rows(1).Delete
End With




' Создаём новый лист "Ю.В." и удаляем, если он уже существует
    On Error Resume Next
    Set NewSheet = ActiveWorkbook.Sheets("Ю.В.")
    If Not NewSheet Is Nothing Then NewSheet.Delete
    On Error GoTo 0

    ' Создаём новый лист "Ю.В."
    Set NewSheet = ActiveWorkbook.Sheets.Add
    NewSheet.Name = "Ю.В."

    ' Копируем сводную таблицу с листа "Свод" на лист "Ю.В." как значения и форматы
    With PivotSheet.UsedRange
        .Copy
        NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
        NewSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats
    End With

    ' Очищаем буфер обмена и включаем обновление экрана
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic



' Отключаем автоизменение ширины столбцов при обновлении
    .PivotTableWizard PreserveFormatting:=True
    .ManualUpdate = False
End With

' Устанавливаем фиксированную ширину 16 для всех столбцов
PivotSheet.Columns.ColumnWidth = 16

' Выполняем автоподгонку ширины только для определённых полей
With PivotSheet
    .Columns("A:A").AutoFit ' "Сегмент"
    .Columns("B:B").AutoFit ' "Ответственный"
    .Columns("C:C").AutoFit ' "Заказчик"
End With






Sub свод_1137()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Ваш код для создания и заполнения данных
    ' ...
    
    ' Макрос для создания свода
    Dim DataSheet As Worksheet
    Dim PivotSheet As Worksheet
    Dim PivotTable As PivotTable
    Dim PivotCache As PivotCache
    Dim LastRow As Long, LastCol As Long
    Dim SourceRange As String
    Dim pf As PivotField
    Dim DataField As PivotField

    ' Отключаем обновление экрана для ускорения выполнения
    Application.ScreenUpdating = False

    ' Определяем активный лист с данными
    Set DataSheet = ActiveSheet

    ' Определяем последний заполненный ряд и столбец на активном листе
    With DataSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        SourceRange = "'" & .Name & "'!A1:" & .Cells(LastRow, LastCol).Address(False, False)
    End With

    ' Преобразуем значения в столбце "Сальдо СФ на конец периода" в числовой формат
    Dim saldoRange As Range
    Set saldoRange = DataSheet.Range("D2:D" & LastRow) ' Замените "D" на букву столбца, где находится "Сальдо СФ на конец периода"
    
    saldoRange.NumberFormat = "General" ' Сначала установите общий формат
    saldoRange.Value = saldoRange.Value ' Преобразуем текстовые значения в числа

    ' Проверяем и удаляем лист "Свод", если он уже существует
    On Error Resume Next
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    If Not PivotSheet Is Nothing Then PivotSheet.Delete
    On Error GoTo 0

    ' Создаём новый лист для сводной таблицы
    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод"

' Окрашиваем весь лист "Свод" в зеленый цвет
PivotSheet.Cells.Interior.Color = RGB(144, 238, 144) ' Светло-зеленый цвет

    ' Создаём PivotCache на основе данных
    Set PivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=SourceRange)

    ' Создаём сводную таблицу на новом листе "Свод"
    Set PivotTable = PivotCache.CreatePivotTable( _
        TableDestination:=PivotSheet.Cells(3, 1), _
        TableName:="СводнаяТаблица")

    ' Настраиваем сводную таблицу
    With PivotTable
        .SmallGrid = True
        .RowAxisLayout xlTabularRow ' Устанавливаем табличную форму

        ' Отключаем промежуточные итоги только для поля "Ответственный"
        On Error Resume Next
        With .PivotFields("Ответственный")
            .Subtotals(1) = False ' Отключение всех промежуточных итогов
        End With

        ' Добавляем поля в строки
        .PivotFields("Сегмент").Orientation = xlRowField
        .PivotFields("Ответственный").Orientation = xlRowField
        .PivotFields("Заказчик").Orientation = xlRowField

        ' Добавляем поля в столбцы
        With .PivotFields("Категория просрочки")
            .Orientation = xlColumnField
            .Position = 1 ' Первым столбцом
            ' Устанавливаем порядок элементов
            .PivotItems("просрочка более 60 дней").Position = 1
            .PivotItems("просрочка от 30 до 60 дней").Position = 2
            .PivotItems("просрочка от 15 до 30 дней").Position = 3
            .PivotItems("просрочка до 15 дней").Position = 4
        End With

        ' Добавляем "Номер недели" как второй столбец
        .PivotFields("Номер недели").Orientation = xlColumnField
        .PivotFields("Номер недели").Position = 2

        ' Добавляем числовое поле
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        On Error GoTo 0

        ' Устанавливаем стиль сводной таблицы: "Средний 8"
        .TableStyle2 = "PivotStyleMedium8"
    End With

    ' Добавляем числовое поле "Сальдо СФ на конец периода" в сводную таблицу
With PivotTable
    .AddDataField .PivotFields("Сальдо СФ на конец периода"), "Итог по Сальдо", xlSum
End With

' Применяем форматирование с двумя знаками после запятой
With PivotTable.DataFields(1)
    .NumberFormat = "#,##0.00"
End With

    ' Обновляем сводную таблицу
    PivotTable.RefreshTable

    ' Сворачиваем все поля сводной таблицы, кроме строк
    For Each pf In PivotTable.PivotFields
        On Error Resume Next
        If pf.Orientation <> xlRowField Then
            pf.ShowDetail = False ' Сворачиваем, если не строка
        End If
        On Error GoTo 0
    Next pf

    ' Включаем обновление экрана
    Application.ScreenUpdating = True

    MsgBox "Сводная таблица успешно создана, отформатирована и свернута на листе 'Свод'!", vbInformation
End Sub
