$excel = New-Object -ComObject excel.application
$excel.visible = $True

$workbook = $excel.Workbooks.Add()

$uregwksht = $workbook.Worksheets.Item(1)
$uregwksht.Name = 'The name you choose'

#later: import all csv's in folder
$records = Import-Csv -Path 'C:\Users\davis\source\repos\PowershellExcelUtil - Files\CSV1.csv'

#writing headers
for ($j = 0; $j -lt $records[0].psobject.properties.name.Length; $j++) {
	$uregwksht.Cells.Item(1, $j + 1) = $records[0].psobject.properties.name[$j]
}
#filling up table
for ($i = 0; $i -lt $records.Length; $i++) {
	for ($j = 0; $j -lt $records[1].psobject.properties.value.Length; $j++) {
		$uregwksht.Cells.Item($i + 2, $j + 1) = $records[$i].psobject.properties.value[$j]
	}
}

#adjusting the column width so all data's properly visible
$usedRange = $uregwksht.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null

$workbook.SaveAs('C:\Users\davis\source\repos\PowershellExcelUtil - Files\myExcel.xlsx')
$excel.Quit()
