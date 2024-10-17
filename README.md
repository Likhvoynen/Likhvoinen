Sub CreateTableM()
ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:="ДЗ 10-24!A1:AZ10000").CreatePivotTable TableDestination:="", TableName:="Сводная таблица"
With ActiveSheet
.Name = "Свод"
.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
End With
With ActiveSheet.PivotTables("Сводная таблица")
.SmallGrid = True
.PivotFields("Сальдо СФ на конец периода").Orientation = xlDataField
.PivotFields("Сегмент").Orientation = xlRowField
.PivotFields("Ответственный").Orientation = xlRowField
.PivotFields("Заказчик").Orientation = xlRowField
.PivotFields("Категория просрочки").Orientation = xlColumnField
.PivotFields("Номер недели").Orientation = xlColumnField
End With
End Sub
