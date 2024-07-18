# カレントディレクトリにExcelファイルを作成
$excelFileName = "HelloWorld.xlsx"
$excelFilePath = Join-Path -Path (Get-Location) -ChildPath $excelFileName

# 既存のExcelファイルがあれば削除
if (Test-Path $excelFilePath) {
    Remove-Item $excelFilePath -Force
}

# Excelアプリケーションを作成
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false  # Excelを非表示で開く

# 新しいワークブックを作成
$workbook = $excelApp.Workbooks.Add()

# ワークブックを保存
$workbook.SaveAs([ref] $excelFilePath)
$workbook.Close()
$excelApp.Quit()

# Excelアプリケーションを再度開く
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $true  # Excelを表示する
$workbook = $excelApp.Workbooks.Open($excelFilePath)

# PowerQueryのMコードを定義
$mCode = @"
let
    Source = #table({"Column1"}, {{"Hello World"}})
in
    Source
"@

# PowerQueryクエリを追加
$pqName = "HelloWorldQuery"
try {
    $workbook.Queries($pqName).Delete()  # 既存のクエリがあれば削除
} catch {
    # クエリが存在しない場合はエラーを無視
}
$workbook.Queries.Add($pqName, $mCode)

# クエリをテーブルとしてシートに追加
$sheet = $workbook.Sheets.Item(1)
$range = $sheet.Range("A1")
$queryTable = $sheet.QueryTables.Add("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$pqName;Location=Sheet1", $range)
$queryTable.CommandText = "[$pqName]"
$queryTable.Name = "HelloWorldQueryTable"
$queryTable.RefreshStyle = 2 # xlInsertEntireRows
$queryTable.AdjustColumnWidth = $true
$queryTable.RefreshOnFileOpen = $true

# クエリテーブルのプロパティを確認して表示
Write-Host "Query Table Properties:"
Write-Host "------------------------"
Write-Host "Name: $($queryTable.Name)"
Write-Host "CommandText: $($queryTable.CommandText)"
Write-Host "Connection: $($queryTable.Connection)"

# クエリをリフレッシュしながら進捗表示
Write-Host "Refreshing query table..."
try {
    $queryTable.Refresh()
    Write-Host "Query table refreshed successfully."
} catch {
    Write-Host "Error refreshing query table: $_"
}

# 保存して閉じる
$workbook.Save()
$workbook.Close($false)
$excelApp.Quit()

# リソースの解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($queryTable) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
