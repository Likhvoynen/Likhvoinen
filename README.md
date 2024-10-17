Sub CreateTableM()
    Dim LastRow As Long, LastCol As Long
    Dim SourceRange As String
    Dim PivotSheet As Worksheet
    Dim DataSheet As Worksheet

    ' Определяем активный лист с данными
    Set DataSheet = ActiveSheet

    ' Определяем последний ряд и последний столбец на активном листе
    With DataSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        SourceRange = .Name & "!A1:" & .Cells(LastRow, LastCol).Address
    End With

    ' Создаём новый лист для сводной таблицы
    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод" ' Можно поменять имя при необходимости

    ' Создаём сводную таблицу на новом листе
    ActiveWorkbook.PivotCaches.Add( _
        SourceType:=xlDatabase, _
        SourceData:=SourceRange).CreatePivotTable _
        TableDestination:=PivotSheet.Cells(3, 1), _
        TableName:="Сводная таблица"

    ' Настраиваем сводную таблицу
    With PivotSheet.PivotTables("Сводная таблица")
        .SmallGrid = True
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        .PivotFields("Сегмент").Orientation = xlRowField
        .PivotFields("Ответственный").Orientation = xlRowField
        .PivotFields("Заказчик").Orientation = xlRowField
        .PivotFields("Категория просрочки").Orientation = xlColumnField
        .PivotFields("Номер недели").Orientation = xlColumnField
    End With

    ' Включаем обновление экрана (на случай, если было отключено)
    Application.ScreenUpdating = True
End Sub