Sub CreateTableM()
    Dim LastRow As Long, LastCol As Long
    Dim SourceRange As String
    Dim PivotSheet As Worksheet
    Dim DataSheet As Worksheet

    ' Отключаем обновление экрана для ускорения работы
    Application.ScreenUpdating = False

    ' Определяем активный лист с данными
    Set DataSheet = ActiveSheet

    ' Проверяем, есть ли данные на активном листе
    If Application.WorksheetFunction.CountA(DataSheet.Cells) = 0 Then
        MsgBox "Активный лист пуст. Пожалуйста, выберите лист с данными.", vbExclamation
        Exit Sub
    End If

    ' Определяем последний ряд и столбец на активном листе
    With DataSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        SourceRange = .Name & "!A1:" & .Cells(LastRow, LastCol).Address
    End With

    ' Проверяем, существует ли лист "Свод", и удаляем его, если нужно
    On Error Resume Next
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    If Not PivotSheet Is Nothing Then PivotSheet.Delete
    On Error GoTo 0

    ' Создаём новый лист для сводной таблицы
    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод"

    ' Создаём сводную таблицу на новом листе
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=SourceRange).CreatePivotTable _
        TableDestination:=PivotSheet.Cells(3, 1), _
        TableName:="Сводная таблица"

    ' Настраиваем сводную таблицу
    With PivotSheet.PivotTables("Сводная таблица")
        .SmallGrid = True

        ' Проверяем и добавляем поля, если они существуют
        On Error Resume Next
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        .PivotFields("Сегмент").Orientation = xlRowField
        .PivotFields("Ответственный").Orientation = xlRowField
        .PivotFields("Заказчик").Orientation = xlRowField
        .PivotFields("Категория просрочки").Orientation = xlColumnField
        .PivotFields("Номер недели").Orientation = xlColumnField
        On Error GoTo 0
    End With

    ' Включаем обновление экрана
    Application.ScreenUpdating = True

    MsgBox "Сводная таблица успешно создана!", vbInformation
End Sub