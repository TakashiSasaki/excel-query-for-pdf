# スクリプト名: ExtractTableFromPdfsToSeparateExcels.ps1
# このスクリプトでは1つのPDFファイルから1つのテーブルしか取り出せない。
# カレントディレクトリの取得
$currentDirectory = Get-Location
$pdfUrls = @(
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-1.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-2.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-3.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-4.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-5.pdf"
)

# PDFファイルのダウンロードと保存
foreach ($pdfUrl in $pdfUrls) {
    $pdfFileName = [System.IO.Path]::GetFileName($pdfUrl)
    $pdfFilePath = "$currentDirectory\$pdfFileName"
    $excelFilePath = "$currentDirectory\$($pdfFileName -replace '.pdf','.xlsx')"

    if (-not (Test-Path $pdfFilePath)) {
        Invoke-WebRequest -Uri $pdfUrl -OutFile $pdfFilePath
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
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "PDFTable"

    # Power Query M 式の作成
    $queryFormula = @"
let
    // PDFファイルのパスを指定
    Source = Pdf.Tables(File.Contents(`"$pdfFilePath`"), [Implementation=`"1.3`"]),
    
    // PDF内のテーブルを選択（複数のテーブルがある場合、インデックスを変更）
    Table = Source{0}[Data]

in
    Table
"@

    # Power Queryの追加
    $query = $workbook.Queries.Add("ExtractTableFromPdf", $queryFormula)

    # 接続文字列の作成
    $connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=ExtractTableFromPdf;Extended Properties="

    # クエリテーブルの追加
    $connection = $workbook.Connections.Add2("ExtractTableFromPdf Connection", "", $connectionString, $queryFormula, 2)
    $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcExternal, $connection, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $worksheet.Range("A1"))
    $listObject.QueryTable.CommandText = "SELECT * FROM [ExtractTableFromPdf]"
    $listObject.QueryTable.Refresh()

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
}
