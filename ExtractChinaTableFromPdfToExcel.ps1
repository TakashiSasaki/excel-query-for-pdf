# スクリプト名: ExtractChinaTableFromPdfToExcel.ps1

# カレントディレクトリの取得
$currentDirectory = Get-Location
$savePath = "$currentDirectory\china_table_data.xlsx"
$pdfPath = "$currentDirectory\china.pdf"
$pdfUrl = "https://raw.githubusercontent.com/tabulapdf/tabula-java/8bfa3ad23af34f757f72fe46584a34abfc022ed3/src/test/resources/technology/tabula/china.pdf"

# 既存のPDFファイルを削除
if (Test-Path $pdfPath) {
    Remove-Item $pdfPath
}

# PDFファイルのダウンロード
Invoke-WebRequest -Uri $pdfUrl -OutFile $pdfPath

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
$worksheet.Name = "PDFTable"

# Power Query M 式の作成
$queryFormula = @"
let
    // PDFファイルのパスを指定
    Source = Pdf.Tables(File.Contents(`"$pdfPath`"), [Implementation=`"1.3`"]),
    
    // PDF内のテーブルを選択（複数のテーブルがある場合、インデックスを変更）
    Table = Source{0}[Data]

in
    Table
"@

# Power Queryの追加
$query = $workbook.Queries.Add("ExtractChinaTableFromPdf", $queryFormula)

# 接続文字列の作成
$connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=ExtractChinaTableFromPdf;Extended Properties="

# クエリテーブルの追加
$connection = $workbook.Connections.Add2("ExtractChinaTableFromPdf Connection", "", $connectionString, $queryFormula, 2)
$listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcExternal, $connection, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $worksheet.Range("A1"))
$listObject.QueryTable.CommandText = "SELECT * FROM [ExtractChinaTableFromPdf]"
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

Write-Output "Excel file created with table data from PDF at $savePath"
