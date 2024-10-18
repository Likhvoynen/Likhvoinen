Option Explicit

Sub RunCombinedMacros()
    ' Отключаем обновление экрана и автоматическое пересчет
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Выполнение первого макроса
    Call new_1137

    ' Выполнение второго макроса
    Call CreateFormattedPivotTable

    ' Включаем обновление экрана и автоматический пересчет
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Оба макроса успешно выполнены!", vbInformation
End Sub

Sub new_1137()
    Columns("R:R").Insert Shift:=xlToRight
    Range("R1").Value = "Дни просрочки с учетом доставки на дату отчета"
    
    With Range("R1").Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.4
    End With

    Dim A As String, B As String
    A = Sheets("Параметры").Range("B2").Value
    B = "=Параметры!R5C2-'ДЗ " & A & "'!RC[-9]"

    FillFormulaInColumn "R", B, "0"
    
    Columns("S:S").Insert Shift:=xlToRight
    Range("S1").Value = "Категория просрочки"
    SetHeaderFormat Range("S1")

    FillFormulaInColumn "S", "=IF(RC[-1]<1,CONCATENATE(YEAR(RC[-10]),""-"",MONTH(RC[-10])),IF(RC[-1]<15,""просрочка до 15 дней"",IF(RC[-1]<30,""просрочка от 15 до 30 дней"",IF(RC[-1]<60,""просрочка от 30 до 60 дней"",""просрочка более 60 дней""))))"
    
    Columns("T:T").Insert Shift:=xlToRight
    Range("T1").Value = "Номер"
    SetHeaderFormat Range("T1")

    FillFormulaInColumn "T", "=WEEKNUM(RC[-11],21)"
    
    Columns("U:U").Insert Shift:=xlToRight
    Range("U1").Value = "Номер недели"
    SetHeaderFormat Range("U1")

    FillFormulaInColumn "U", "=CONCATENATE(RC[-1],"" неделя"")"
    
    SetHeaderFormat Range("I1")
End Sub

Sub CreateFormattedPivotTable()
    Dim DataSheet As Worksheet, PivotSheet As Worksheet
    Dim PivotTable As PivotTable, PivotCache As PivotCache
    Dim LastRow As Long, LastCol As Long, SourceRange As String

    Set DataSheet = ActiveSheet

    With DataSheet
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        SourceRange = "'" & .Name & "'!A1:" & .Cells(LastRow, LastCol).Address(False, False)
    End With

    On Error Resume Next
    Set PivotSheet = ActiveWorkbook.Sheets("Свод")
    If Not PivotSheet Is Nothing Then PivotSheet.Delete
    On Error GoTo 0

    Set PivotSheet = ActiveWorkbook.Sheets.Add
    PivotSheet.Name = "Свод"

    Set PivotCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceRange)
    Set PivotTable = PivotCache.CreatePivotTable(PivotSheet.Cells(3, 1), "СводнаяТаблица")

    With PivotTable
        .SmallGrid = True
        .RowAxisLayout xlTabularRow

        With .PivotFields("Ответственный")
            .Subtotals(1) = False
        End With

        .PivotFields("Сегмент").Orientation = xlRowField
        .PivotFields("Ответственный").Orientation = xlRowField
        .PivotFields("Заказчик").Orientation = xlRowField

        With .PivotFields("Категория просрочки")
            .Orientation = xlColumnField
            .PivotItems("просрочка более 60 дней").Position = 1
            .PivotItems("просрочка от 30 до 60 дней").Position = 2
            .PivotItems("просрочка от 15 до 30 дней").Position = 3
            .PivotItems("просрочка до 15 дней").Position = 4
        End With

        .PivotFields("Номер недели").Orientation = xlColumnField
        .PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
        .TableStyle2 = "PivotStyleMedium8"
    End With
End Sub

Sub FillFormulaInColumn(columnLetter As String, formula As String, Optional numberFormat As String = "")
    Dim checkup As Boolean
    Dim cell As Range
    Set cell = Range(columnLetter & "2")
    checkup = False

    Do While Not checkup
        cell.FormulaR1C1 = formula
        If numberFormat <> "" Then cell.NumberFormat = numberFormat
        Set cell = cell.Offset(1, 0)
        checkup = IsEmpty(cell.Offset(0, -1))
    Loop
End Sub

Sub SetHeaderFormat(rng As Range)
    With rng.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.4
    End With
End Sub