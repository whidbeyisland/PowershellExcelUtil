$excel = New-Object -ComObject excel.application
$excel.visible = $True

$workbook = $excel.Workbooks.Add()
$sourcePath = 'C:\Users\davis\source\repos\PowershellExcelUtil - Files'
$csvFiles = Get-ChildItem $sourcePath -Filter *.csv

for ($k = 1; $k -lt $csvFiles.Length; $k++) {
	$workbook.Worksheets.Add()
}
for ($k = 0; $k -lt $csvFiles.Length; $k++) {
	#import CSV
	$records = Import-Csv -Path ($sourcePath + "\" + $csvFiles[$k])
	
	#add a new worksheet
	$uregwksht = $workbook.Worksheets.Item($k + 1)
	$uregwksht.Name = $csvFiles[$k]
	
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
}

$workbook.SaveAs('C:\Users\davis\source\repos\PowershellExcelUtil - Files\myExcel.xlsx')
$excel.Quit()