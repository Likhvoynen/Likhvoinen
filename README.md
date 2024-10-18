Option Explicit

Sub свод_1137()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Дни просрочки с учетом доставки на дату отчета"
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("R2").Select
    
    Dim checkup As Boolean
    Dim i As Integer
    i = 0
    
    Dim A As String
    Dim B As String
    
    A = Sheets("Параметры").Range("B2")
    B = "=Параметры!R5C2-'ДЗ " + A + "'!RC[-9]"
    
    Do While checkup = False
        checkup = IsEmpty(ActiveCell.Offset(1, -1))
        ActiveCell.FormulaR1C1 = B
        Selection.numberFormat = "0"
        i = i + 1
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    Columns("S:S").Select
    Selection.Insert Shift:=xlToRight
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Категория просрочки"
   
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
    .PatternTintAndShade = 0
    End With
    
    Range("S2").Select
    
    checkup = False
    i = 0
 
    Do While checkup = False
      checkup = IsEmpty(ActiveCell.Offset(1, -1))
        ActiveCell.FormulaR1C1 = "=IF(RC[-1]<1,CONCATENATE(YEAR(RC[-10]),""-"",MONTH(RC[-10])),IF(RC[-1]<15,""просрочка до 15 дней"",IF(RC[-1]<30,""просрочка от 15 до 30 дней"",IF(RC[-1]<60,""просрочка от 30 до 60 дней"",""просрочка более 60 дней""))))"
        i = i + 1
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Номер"
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("T2").Select
    
    checkup = False
    i = 0
 
    Do While checkup = False
        checkup = IsEmpty(ActiveCell.Offset(1, -1))
        ActiveCell.FormulaR1C1 = "=WEEKNUM(RC[-11],21)"
        i = i + 1
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    Columns("U:U").Select
    Selection.Insert Shift:=xlToRight
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Номер недели"
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("U2").Select

    checkup = False
    i = 0
  
    Do While checkup = False
        checkup = IsEmpty(ActiveCell.Offset(1, -1))
        ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],"" неделя"")"
        i = i + 1
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    Range("I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic



'Макрос для создания листа свод

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
        .RowAxisLayout xlTabularRow ' Устанавливаем табличную форму

        ' Отключаем промежуточные итоги только для поля "Ответственный"
        On Error Resume Next
        With .PivotFields("Ответственный")
            .Subtotals(1) = False ' Отключение всех промежуточных итогов
        End With

        ' Добавляем поля в строки
        .PivotFields("Сегмент").orientation = xlRowField
        .PivotFields("Ответственный").orientation = xlRowField
        .PivotFields("Заказчик").orientation = xlRowField

        ' Добавляем поля в столбцы
        With .PivotFields("Категория просрочки")
            .orientation = xlColumnField
            .position = 1 ' Первым столбцом
            ' Устанавливаем порядок элементов
            .PivotItems("просрочка более 60 дней").position = 1
            .PivotItems("просрочка от 30 до 60 дней").position = 2
            .PivotItems("просрочка от 15 до 30 дней").position = 3
            .PivotItems("просрочка до 15 дней").position = 4
        End With

        ' Добавляем "Номер недели" как второй столбец
        .PivotFields("Номер недели").orientation = xlColumnField
        .PivotFields("Номер недели").position = 2

        ' Добавляем числовое поле
        .PivotFields("Сальдо СФ на конец периода").orientation = xlDataField
        On Error GoTo 0

        ' Устанавливаем стиль сводной таблицы: "Средний 8"
        .TableStyle2 = "PivotStyleMedium8"
    End With

' Преобразуем значения в столбце "Сальдо СФ на конец периода" в финансовый формат
Set DataField = PivotTable.PivotFields("Сальдо СФ на конец периода")
With DataField
    .NumberFormat = "_-* #,##0.00_₽_-;_-*
end with

    ' Сворачиваем все поля сводной таблицы, кроме строк
    For Each pf In PivotTable.PivotFields
        On Error Resume Next
        If pf.orientation <> xlRowField Then
            pf.ShowDetail = False ' Сворачиваем, если не строка
        End If
        On Error GoTo 0
    Next pf

    ' Включаем обновление экрана
    Application.ScreenUpdating = True

    MsgBox "Сводная таблица успешно создана, отформатирована и свернута на листе 'Свод'!", vbInformation
End Sub
