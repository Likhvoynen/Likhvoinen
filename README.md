Sub CreateTableM()
    Dim LastRow As Long, LastCol As Long
    Dim SourceRange As String
    Dim PivotSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim FieldName As Variant
    Dim Headers As Object

    ' Отключаем обновление экрана для ускорения
    Application.ScreenUpdating = False

    ' Определяем активный лист с данными
    Set DataSheet = ActiveSheet

    ' Проверяем, есть ли данные на активном листе
    If Application.WorksheetFunction.CountA(DataSheet.Cells) = 0 Then
        MsgBox "Активный лист пуст. Пожалуйста, выберите лист с данными.", vbExclamation
        Exit Sub
    End If

    ' Определяем последний ряд и последний столбец
    With DataSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        SourceRange = .Name & "!A1:" & .Cells(LastRow, LastCol).Address
    End With

    ' Сохраняем заголовки в объект Scripting.Dictionary для проверки полей
    Set Headers = CreateObject("Scripting.Dictionary")
    For Each FieldName In DataSheet.Rows(1).Columns(1).Resize(1, LastCol).Value
        Headers(FieldName) = True
    Next FieldName

    ' Проверяем и удаляем лист "Свод", если он уже существует
    On Error Resume Next
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    If Not PivotSheet Is Nothing Then PivotSheet.Delete
    On Error GoTo 0

    ' Создаём новый лист для сводной таблицы
    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод"

    ' Создаём сводную таблицу
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=SourceRange).CreatePivotTable _
        TableDestination:=PivotSheet.Cells(3, 1), _
        TableName:="Сводная таблица"

    ' Настраиваем сводную таблицу
    With PivotSheet.PivotTables("Сводная таблица")
        .SmallGrid = True

        ' Добавляем поля только если они существуют в данных
        If Headers.exists("Сальдо СФ на конец периода") Then
            .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        End If
        If Headers.exists("Сегмент") Then
            .PivotFields("Сегмент").Orientation = xlRowField
        End If
        If Headers.exists("Ответственный") Then
            .PivotFields("Ответственный").Orientation = xlRowField
        End If
        If Headers.exists("Заказчик") Then
            .PivotFields("Заказчик").Orientation = xlRowField
        End If
        If Headers.exists("Категория просрочки") Then
            .PivotFields("Категория просрочки").Orientation = xlColumnField
        End If
        If Headers.exists("Номер недели") Then
            .PivotFields("Номер недели").Orientation = xlColumnField
        End If
    End With

    ' Включаем обновление экрана
    Application.ScreenUpdating = True

    MsgBox "Сводная таблица успешно создана!", vbInformation
End Sub