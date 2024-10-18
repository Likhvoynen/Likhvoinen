Sub CreateFormattedPivotTable()
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

        ' Добавляем "Категория просрочки" как первый столбец и задаём порядок элементов
        With .PivotFields("Категория просрочки")
            .Orientation = xlColumnField
            .Position = 1 ' Первым столбцом
            ' Устанавливаем порядок элементов
            .PivotItems("просрочка более 60 дней").Position = 1
            .PivotItems("просрочка от 30 до 60 дней").Position = 2
            .PivotItems("просрочка от 15 до 30 дней").Position = 3
            .PivotItems("просрочка до 15 дней").Position = 4
        End With

        ' Добавляем "Номер недели" как второй столбец и сортируем по возрастанию
        With .PivotFields("Номер недели")
            .Orientation = xlColumnField
            .Position = 2 ' После "Категория просрочки"
            .AutoSort xlAscending, "Номер недели" ' Сортировка по возрастанию
        End With

        ' Добавляем числовое поле
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        On Error GoTo 0

        ' Устанавливаем стиль сводной таблицы: "Средний 8"
        .TableStyle2 = "PivotStyleMedium8"
    End With

    ' Преобразуем значения в столбце "Сальдо СФ на конец периода" в числовой формат
    Set DataField = PivotTable.PivotFields("Сальдо СФ на конец периода")
    With DataField
        .NumberFormat = "#,##0.00" ' Формат с двумя знаками после запятой
    End With

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