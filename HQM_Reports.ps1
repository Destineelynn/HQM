# Open New Excel Workbook / WorkSheet
$Excel = new-object -comobject excel.application
$Entries = @{}
$File = "X:/Quality/Reports/2017/HQM/Q3/PM_3_2017_Primary/Campbell_3_2017/Campbell_Diabetic Urine_3_2017.xlsx"
$Workbook = $Excel.workbooks.open($File)
$Worksheet = $Workbook.Worksheets.Item(1)
$Worksheet.Activate()
$Excel.Visible = $true



[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,6).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,8).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,5).EntireColumn.Delete()
[void]$Worksheet.Cells.Item(15,3).EntireColumn.Delete()


$range = $Worksheet.Range("D13","F500")
$range.clear()



#$rows = $Worksheet.range("C1").currentregion.rows.count
#$Worksheet.range("F2:F$rows").formula = $Worksheet.range("F2").formula





##########################################################

$Excel2 = new-object -comobject excel.application
$Entries = @{}
$File2 = "X:/Quality/Reports/2017/HQM/Q3/PM_2017_Primary/Campbell_2017/Campbell_Diabetic Urine_2017.xlsx"
$Workbook2 = $Excel2.workbooks.open($File2)
$Worksheet2 = $Workbook2.Worksheets.Item(1)
$Worksheet2.Activate()
$Excel2.Visible = $true

$Worksheet1 = $Workbook.Worksheets.Add()
$SourceRange=$Workbook2.Worksheets.Item(1).range("A1","Z1000");
$SourceRange.copy() | out-null;
$Workbook.worksheets.item(1).paste();
$Workbook2.close($false) # close source workbook w/o saving 
$Excel2.close
##########################################################


$Workbook.worksheets.item(2).cells.item(13,4).Formula = ('=VLOOKUP(B13,Sheet1!$B$1:$P$1000,12,false')
$Workbook.worksheets.item(2).cells.item(13,5).Formula = ('=VLOOKUP(B13,Sheet1!$B$1:$P$1000,13,false')
$Workbook.worksheets.item(2).cells.item(13,6).Formula = ('=VLOOKUP(B13,Sheet1!$B$1:$P$1000,15,false')



