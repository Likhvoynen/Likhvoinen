Sub CreateFormattedPivotTable()
    Dim DataSheet As Worksheet
    Dim PivotSheet As Worksheet
    Dim PivotTable As PivotTable
    Dim PivotCache As PivotCache
    Dim LastRow As Long, LastCol As Long
    Dim SourceRange As String

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

    ' Проверяем и удаляем лист "Свод", если он уже существует
    On Error Resume Next
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    If Not PivotSheet Is Nothing Then PivotSheet.Delete
    On Error GoTo 0

    ' Создаём новый лист для сводной таблицы
    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод"

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
        .ShowTableStyleRowStripes = True ' Полосатые строки

        ' Устанавливаем табличную форму
        .RowAxisLayout xlTabularRow

        ' Отключаем промежуточные итоги для полей "Сегмент" и "Ответственный"
        On Error Resume Next
        With .PivotFields("Сегмент")
            .Subtotals(1) = False ' Отключение всех промежуточных итогов
        End With
        With .PivotFields("Ответственный")
            .Subtotals(1) = False
        End With

        ' Добавляем поля (замени на нужные)
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        .PivotFields("Сегмент").Orientation = xlRowField
        .PivotFields("Ответственный").Orientation = xlRowField
        .PivotFields("Заказчик").Orientation = xlRowField
        .PivotFields("Категория просрочки").Orientation = xlColumnField
        .PivotFields("Номер недели").Orientation = xlColumnField
        On Error GoTo 0

        ' Устанавливаем стиль сводной таблицы: "Средний 8"
        .TableStyle2 = "PivotStyleMedium8"
    End With

    ' Включаем обновление экрана
    Application.ScreenUpdating = True

    MsgBox "Сводная таблица успешно создана и отформатирована на листе 'Свод'!", vbInformation
End Sub