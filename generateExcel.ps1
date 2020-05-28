$excel = New-Object -ComObject excel.application
$excel.visible = $True

$workbook = $excel.Workbooks.Add()

#$workbook.Worksheets.Item(3).Delete()

$uregwksht= $workbook.Worksheets.Item(1)
$uregwksht.Name = 'The name you choose'

$row = 1
$column = 1
$uregwksht.Cells.Item($row,$column)= 'Title'

#merging a few cells on the top row to make the title look nicer
$MergeCells = $uregwksht.Range("A1:G1")
$MergeCells.Select()
$MergeCells.MergeCells = $true
$uregwksht.Cells(1, 1).HorizontalAlignment = -4108

$uregwksht.Cells.Item(1,1).Font.Size = 18
$uregwksht.Cells.Item(1,1).Font.Bold=$True
$uregwksht.Cells.Item(1,1).Font.Name = "Cambria"
$uregwksht.Cells.Item(1,1).Font.ThemeFont = 1
$uregwksht.Cells.Item(1,1).Font.ThemeColor = 4
$uregwksht.Cells.Item(1,1).Font.ColorIndex = 55
$uregwksht.Cells.Item(1,1).Font.Color = 8210719

#create the column headers
$uregwksht.Cells.Item(3,1) = 'Date'
$uregwksht.Cells.Item(3,2) = 'Hour'
$uregwksht.Cells.Item(3,3) = 'Name'

$records = Import-Csv -Path 'C:\Users\davis\source\repos\PowershellExcelUtil - Files\CSV1.csv'

Write-Host $records[0].psobject.properties.value[0]

#iterate through entire sheet like this
$uregwksht.Cells.Item(5,1) = $records[1].psobject.properties.value[1]
$uregwksht.Cells.Item(5,2) = $records[1].psobject.properties.value[2]
$uregwksht.Cells.Item(5,3) = $records[1].psobject.properties.value[3]

#adjusting the column width so all data's properly visible
$usedRange = $uregwksht.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null

$workbook.SaveAs('C:\Users\davis\source\repos\PowershellExcelUtil - Files\myExcel.xlsx')
$excel.Quit()
