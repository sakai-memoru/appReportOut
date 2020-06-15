# param (
#     [Parameter(Mandatory=$true)][string]$excelFile = "appOutReport.xlsm",
#     [Parameter(Mandatory=$true)][string]$macro = "FieldDefMain.Batch",
#     [Parameter(Mandatory=$true)][string]$formName = "FIELD_DEF",
#     [Parameter(Mandatory=$true)][boolean]$outTemplOn = $true,
#     [Parameter(Mandatory=$true)][boolean]$moveOn = $false
# )

$excelFile = "appReportOut.xlsm"
$macro = "ReportDefMain.Batch"
$formName = "REPORT_DEF"
$outTemplOn = $false
$outTemplHtmlOn = $false
$dumpOn = $false
$moveOn = $false
##
$curFolder = pwd 
$fullpath = Join-Path $curFolder.Path $excelFile
$excel = new-object -comobject excel.application
$excel.Visible = $false
$workbook = $excel.workbooks.open($fullpath)
$null = $excel.Run($macro, $formName, $outTemplOn, $outTemplHtmlOn, $dumpOn, $moveOn)
$workbook.close()
$excel.Quit()
echo 'finish ..... !'
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
$null = Remove-Variable excel
