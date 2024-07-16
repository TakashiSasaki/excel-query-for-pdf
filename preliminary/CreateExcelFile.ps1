# スクリプト名: CreateExcelFile.ps1

# カレントディレクトリの取得
$currentDirectory = Get-Location

# Excel COMオブジェクトの作成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# 新しいワークブックの作成
$workbook = $excel.Workbooks.Add()

# シートへのデータの入力（必要に応じて）
$sheet = $workbook.Sheets.Item(1)
$sheet.Cells.Item(1,1) = "Hello"
$sheet.Cells.Item(1,2) = "World"

# ファイルの保存
$savePath = "$currentDirectory\example.xlsx"
$workbook.SaveAs($savePath)

# Excelプロセスのクリーンアップ
$workbook.Close($false)
$excel.Quit()

# COMオブジェクトの解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel, workbook, sheet

Write-Output "Excel file created at $savePath"
