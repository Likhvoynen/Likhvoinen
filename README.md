Option Explicit

Sub CreateReportAndPivotTable()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Добавление и заполнение столбца "Дни просрочки с учетом доставки на дату отчета"
    Columns("R:R").Insert Shift:=xlToRight
    Range("R1").Value = "Дни просрочки с учетом доставки на дату отчета"
    ApplyHeaderStyle Range("R1")
    
    FillColumnWithFormula Range("R2"), _
        "=Параметры!R5C2-'ДЗ " & Sheets("Параметры").Range("B2").Value & "'!RC[-9]"
    
    ' Добавление и заполнение столбца "Категория просрочки"
    Columns("S:S").Insert Shift:=xlToRight
    Range("S1").Value = "Категория просрочки"
    ApplyHeaderStyle Range("S1")
    
    FillColumnWithFormula Range("S2"), _
        "=IF(RC[-1]<1, CONCATENATE(YEAR(RC[-10]), ""-"", MONTH(RC[-10])), " & _
        "IF(RC[-1]<15, ""просрочка до 15 дней"", IF(RC[-1]<30, ""просрочка от 15 до 30 дней"", " & _
        "IF(RC[-1]<60, ""просрочка от 30 до 60 дней"", ""просрочка более 60 дней""))))"

    ' Добавление и заполнение столбца "Номер"
    Columns("T:T").Insert Shift:=xlToRight
    Range("T1").Value = "Номер"
    ApplyHeaderStyle Range("T1")

    FillColumnWithFormula Range("T2"), "=WEEKNUM(RC[-11], 21)"

    ' Добавление и заполнение столбца "Номер недели"
    Columns("U:U").Insert Shift:=xlToRight
    Range("U1").Value = "Номер недели"
    ApplyHeaderStyle Range("U1")

    FillColumnWithFormula Range("U2"), "=CONCATENATE(RC[-1], "" неделя"")"

    ' Создание и настройка сводной таблицы
    CreateFormattedPivotTable

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Отчёт и сводная таблица успешно созданы!", vbInformation
End Sub

' Заполняет колонку формулой до первой пустой ячейки в соседней колонке
Private Sub FillColumnWithFormula(startCell As Range, formula As String)
    Dim checkup As Boolean: checkup = False
    Dim i As Integer: i = 0

    Do While Not checkup
        checkup = IsEmpty(startCell.Offset(i, -1))
        startCell.Offset(i, 0).FormulaR1C1 = formula
        i = i + 1
    Loop
End Sub

' Применяет стиль заголовка к ячейке
Private Sub ApplyHeaderStyle(cell As Range)
    With cell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.4
    End With
End Sub

' Создание сводной таблицы
Private Sub CreateFormattedPivotTable()
    Dim DataSheet As Worksheet, PivotSheet As Worksheet
    Dim PivotTable As PivotTable, PivotCache As PivotCache
    Dim LastRow As Long, LastCol As Long, SourceRange As String

    Set DataSheet = ActiveSheet

    ' Определение диапазона данных
    With DataSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        SourceRange = "'" & .Name & "'!A1:" & .Cells(LastRow, LastCol).Address(False, False)
    End With

    ' Создание листа "Свод"
    On Error Resume Next
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    If Not PivotSheet Is Nothing Then PivotSheet.Delete
    On Error GoTo 0

    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод"

    ' Создание сводной таблицы
    Set PivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=SourceRange)

    Set PivotTable = PivotCache.CreatePivotTable( _
        TableDestination:=PivotSheet.Cells(3, 1), TableName:="СводнаяТаблица")

    ' Настройка сводной таблицы
    With PivotTable
        .SmallGrid = True
        .RowAxisLayout xlTabularRow

        ' Настройка полей
        AddPivotField .PivotFields("Сегмент"), xlRowField
        AddPivotField .PivotFields("Ответственный"), xlRowField
        AddPivotField .PivotFields("Заказчик"), xlRowField

        ' Поле "Категория просрочки"
        With .PivotFields("Категория просрочки")
            .Orientation = xlColumnField
            .Position = 1
            .PivotItems("просрочка более 60 дней").Position = 1
            .PivotItems("просрочка от 30 до 60 дней").Position = 2
            .PivotItems("просрочка от 15 до 30 дней").Position = 3
            .PivotItems("просрочка до 15 дней").Position = 4
        End With

        ' Поле "Номер недели"
        AddPivotField .PivotFields("Номер недели"), xlColumnField, 2

        ' Числовое поле
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField

        ' Применение стиля
        .TableStyle2 = "PivotStyleMedium8"
    End With
End Sub

' Упрощённая функция добавления поля в сводную таблицу
Private Sub AddPivotField(pf As PivotField, orientation As XlPivotFieldOrientation, Optional position As Integer = 1)
    With pf
        .Orientation = orientation
        .Position = position
    End With
End Sub