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
