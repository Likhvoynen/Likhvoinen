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