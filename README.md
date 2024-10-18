Sub CreateFormattedPivotTable()
    Dim DataSheet As Worksheet
    Dim PivotSheet As Worksheet
    Dim PivotTable As PivotTable
    Dim PivotCache As PivotCache
    Dim LastRow As Long, LastCol As Long
    Dim SourceRange As String
    Dim pf As PivotField
    Dim DataField As PivotField
    Dim pi As PivotItem
    Dim Weeks As Object
    Dim i As Long

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

        ' Отключаем промежуточные итоги для "Ответственный"
        On Error Resume Next
        With .PivotFields("Ответственный")
            .Subtotals(1) = False
        End With

        ' Добавляем поля в строки
        .PivotFields("Сегмент").Orientation = xlRowField
        .PivotFields("Ответственный").Orientation = xlRowField
        .PivotFields("Заказчик").Orientation = xlRowField

        ' Настраиваем поле "Категория просрочки"
        With .PivotFields("Категория просрочки")
            .Orientation = xlColumnField
            .Position = 1
            .PivotItems("просрочка более 60 дней").Position = 1
            .PivotItems("просрочка от 30 до 60 дней").Position = 2
            .PivotItems("просрочка от 15 до 30 дней").Position = 3
            .PivotItems("просрочка до 15 дней").Position = 4
        End With

        ' Добавляем поле "Номер недели" и готовимся к его сортировке
        With .PivotFields("Номер недели")
            .Orientation = xlColumnField
            .Position = 2
        End With
    End With

    ' Создаём коллекцию для хранения номеров недель в виде дат
    Set Weeks = CreateObject("Scripting.Dictionary")

    ' Заполняем коллекцию "Weeks" элементами из поля "Номер недели"
    For Each pi In PivotTable.PivotFields("Номер недели").PivotItems
        ' Преобразуем номер недели в дату
        Weeks.Add pi.Name, CDate(Replace(pi.Name, "-", "/") & "-1")
    Next pi

    ' Сортируем элементы по дате
    Dim SortedWeeks As Variant
    SortedWeeks = Weeks.Keys
    Call QuickSort(SortedWeeks, Weeks)

    ' Перетаскиваем элементы в нужном порядке
    For i = LBound(SortedWeeks) To UBound(SortedWeeks)
        PivotTable.PivotFields("Номер недели").PivotItems(SortedWeeks(i)).Position = i + 1
    Next i

    ' Преобразуем числовое поле "Сальдо СФ на конец периода"
    Set DataField = PivotTable.PivotFields("Сальдо СФ на конец периода")
    With DataField
        .NumberFormat = "#,##0.00"
    End With

    ' Включаем обновление экрана
    Application.ScreenUpdating = True

    MsgBox "Сводная таблица успешно создана и отсортирована!", vbInformation
End Sub

' Функция быстрой сортировки (QuickSort) для сортировки элементов
Sub QuickSort(arr As Variant, dict As Object)
    QuickSortRecursive arr, dict, LBound(arr), UBound(arr)
End Sub

Sub QuickSortRecursive(arr As Variant, dict As Object, first As Long, last As Long)
    Dim pivot As Variant, i As Long, j As Long, temp As Variant
    pivot = arr((first + last) \ 2)
    i = first
    j = last

    Do While i <= j
        Do While dict(arr(i)) < dict(pivot)
            i = i + 1
        Loop
        Do While dict(arr(j)) > dict(pivot)
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortRecursive arr, dict, first, j
    If i < last Then QuickSortRecursive arr, dict, i, last
End Sub