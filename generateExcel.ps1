$excel = New-Object -ComObject excel.application
$excel.visible = $True

$workbook = $excel.Workbooks.Add()

#$workbook.Worksheets.Item(3).Delete()

$uregwksht= $workbook.Worksheets.Item(1)
$uregwksht.Name = 'The name you choose'

$row = 1
$column = 1
$uregwksht.Cells.Item($row,$column)= 'Title'

#...

$workbook.SaveAs('C:\Users\davis\source\repos\PowershellExcelUtil\myExcel.xlsx')
$excel.Quit()
