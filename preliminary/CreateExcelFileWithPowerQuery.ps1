# スクリプト名: CreateExcelFileWithPowerQuery.ps1

# カレントディレクトリの取得
$currentDirectory = Get-Location
$savePath = "$currentDirectory\example_with_query.xlsx"

# 既存のExcelファイルを削除
if (Test-Path $savePath) {
    Remove-Item $savePath
}

# Excel COMオブジェクトの作成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 新しいワークブックの作成
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Name = "QueryResult"

# Power Queryの作成
$queryFormula = @"
let
    Source = {1..10},
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error)
in
    #"Converted to Table"
"@

# Power Queryの追加
$query = $workbook.Queries.Add("SampleQuery", $queryFormula)

# 接続文字列の作成
$connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=SampleQuery;Extended Properties="

# クエリテーブルの追加
$connection = $workbook.Connections.Add2("SampleQuery Connection", "", $connectionString, $queryFormula, 2)
$listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcExternal, $connection, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $worksheet.Range("A1"))
$listObject.QueryTable.CommandText = "SELECT * FROM [SampleQuery]"
$listObject.QueryTable.Refresh()

# ファイルの保存
$workbook.SaveAs($savePath)

# Excelプロセスのクリーンアップ
$workbook.Close($false)
$excel.Quit()

# COMオブジェクトの解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($listObject) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($connection) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel, workbook, worksheet, listObject, connection

Write-Output "Excel file with Power Query created at $savePath"
