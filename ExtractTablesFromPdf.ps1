# スクリプト名: ExtractTablesFromPdf.ps1
param (
    [string]$pdfFileName,
    [string]$excelFilePath,
    [switch]$skipVersionCheck
)

function Show-Help {
    Write-Output "Usage: .\ExtractTablesFromPdf.ps1 -pdfFileName <PDF file path> [-excelFilePath <Output Excel file path>] [-skipVersionCheck]"
    Write-Output ""
    Write-Output "Arguments:"
    Write-Output "  -pdfFileName    : Path to the input PDF file"
    Write-Output "  -excelFilePath  : (Optional) Path to the output Excel file. If not specified, the output file will have the same name as the input file with .xlsx extension"
    Write-Output "  -skipVersionCheck : (Optional) Skip PowerShell version and architecture check"
}

# パラメータ検証
try {
    $params = $PSCmdlet.MyInvocation.BoundParameters.Keys
    foreach ($param in $params) {
        if ($param -notin @('pdfFileName', 'excelFilePath', 'skipVersionCheck')) {
            throw "Invalid parameter: $param"
        }
    }
} catch {
    Write-Error $_
    Show-Help
    exit
}

# PowerShellのバージョン情報を表示
Write-Output "PowerShell Version: $($PSVersionTable.PSVersion)"
if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
    Write-Output "PowerShell is running in 64-bit mode."
} else {
    Write-Output "PowerShell is running in 32-bit mode."
}

# PowerShellバージョンとアーキテクチャのチェック
if (-not $skipVersionCheck) {
    if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) {
        Write-Error "This script requires PowerShell version 5.1."
        exit
    }
    if ($env:PROCESSOR_ARCHITECTURE -ne 'x86') {
        Write-Error "This script requires 32-bit PowerShell."
        exit
    }
}

# 入力パラメータの検証
if (-not $pdfFileName) {
    Write-Error "Please provide the PDF file name as an argument."
    Show-Help
    exit
}

# PDFファイルの絶対パスを指定
$pdfFilePath = (Resolve-Path $pdfFileName).Path

# 出力先のExcelファイルのパスが指定されていない場合、PDFファイルの拡張子をxlsxに変更したものとする
if (-not $excelFilePath) {
    $excelFilePath = [System.IO.Path]::ChangeExtension($pdfFilePath, ".xlsx")
}

if (-not (Test-Path $pdfFilePath)) {
    Write-Error "The specified PDF file does not exist: $pdfFilePath"
    exit
}

if (Test-Path $excelFilePath) {
    Remove-Item $excelFilePath
}

# Excel COMオブジェクトの作成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 新しいワークブックの作成
$workbook = $excel.Workbooks.Add()

# Power Query M 式の作成（テーブルの数を取得）
$queryFormula = @"
let
    // PDFファイルのパスを指定
    Source = Pdf.Tables(File.Contents(`"$pdfFilePath`"), [Implementation=`"1.3`"])
in
    Source
"@

# Power Queryの追加
$query = $workbook.Queries.Add("GetTablesFromPdf", $queryFormula)

# 接続文字列の作成
$connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=GetTablesFromPdf;Extended Properties="

# クエリテーブルの追加
$connection = $workbook.Connections.Add2("GetTablesFromPdf Connection", "", $connectionString, $queryFormula, 2)
$listObject = $workbook.Worksheets.Item(1).ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcExternal, $connection, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $workbook.Worksheets.Item(1).Range("A1"))
$listObject.QueryTable.CommandText = "SELECT Id FROM [GetTablesFromPdf]"

# Refresh() の前後に進捗表示を追加する
Write-Output "Refreshing query table to get table IDs..."
$dummy = $listObject.QueryTable.Refresh()
Write-Output "Query table refreshed."

# Id列の値を取得して一行に表示
$idColumnValues = $listObject.Range.Columns.Item(1).Cells | Select-Object -ExpandProperty Value2
Write-Output "Id column values: $($idColumnValues -join ' ')"

# 'Table'で始まり、そのあとに数字が続く文字列をフィルタリング
$tableIds = @()
foreach ($id in $idColumnValues) {
    if ($id -match "^Table\d+$") {
        $tableIds += $id
    }
}
Write-Output "Filtered Table IDs: $($tableIds -join ' ')"

# 各テーブルIDに対して処理を追加
foreach ($tableId in $tableIds) {
    Write-Output "Processing table: $tableId"
    $worksheet = $workbook.Worksheets.Add()
    $worksheet.Name = $tableId

    $tableQueryFormula = @"
let
    // PDFファイルのパスを指定
    Source = Pdf.Tables(File.Contents(`"$pdfFilePath`"), [Implementation=`"1.3`"]),
    
    // PDF内のテーブルを選択
    Table = Source{[Id=`"$tableId`"]}[Data]

in
    Table
"@

    # Power Queryの追加
    $query = $workbook.Queries.Add("ExtractTableFromPdf_$tableId", $tableQueryFormula)

    # 接続文字列の作成
    $connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=ExtractTableFromPdf_$tableId;Extended Properties="

    # クエリテーブルの追加
    $connection = $workbook.Connections.Add2("ExtractTableFromPdf_$tableId Connection", "", $connectionString, $tableQueryFormula, 2)
    $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcExternal, $connection, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $worksheet.Range("A1"))
    $listObject.QueryTable.CommandText = "SELECT * FROM [ExtractTableFromPdf_$tableId]"
    
    # Refresh() の前後に進捗表示を追加する
    Write-Output "Refreshing query table for $tableId..."
    $dummy = $listObject.QueryTable.Refresh()
    Write-Output "Query table for $tableId refreshed."
}

# ファイルの保存
$workbook.SaveAs($excelFilePath)

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

Write-Output "Excel file created with table data from $pdfFileName at $excelFilePath"
