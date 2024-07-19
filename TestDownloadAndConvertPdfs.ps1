# スクリプト名: TestDownloadAndConvertPdfs.ps1

# 指定されたURLリスト
$urls = @(
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-1.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-2.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-3.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-4.pdf",
    "https://raw.githubusercontent.com/thoqbk/traprange/master/_Docs/sample-5.pdf"
)

# スクリプトのディレクトリを取得
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# ExtractTablesFromPdf.ps1 の絶対パスを取得
$extractScriptPath = Join-Path -Path $scriptDir -ChildPath "ExtractTablesFromPdf.ps1"

# ダウンロードディレクトリの設定
$downloadDir = Join-Path -Path $scriptDir -ChildPath "DownloadedPdfs"
if (-not (Test-Path $downloadDir)) {
    New-Item -ItemType Directory -Path $downloadDir | Out-Null
}

# PDFファイルをダウンロード
foreach ($url in $urls) {
    $fileName = [System.IO.Path]::GetFileName($url)
    $outputPath = Join-Path -Path $downloadDir -ChildPath $fileName
    Invoke-WebRequest -Uri $url -OutFile $outputPath
    Write-Output "Downloaded $fileName to $outputPath"

    # PDFをExcelに変換
    $excelOutputPath = [System.IO.Path]::ChangeExtension($outputPath, ".xlsx")
    & $extractScriptPath -pdfFileName $outputPath -excelFilePath $excelOutputPath
    Write-Output "Converted $fileName to $excelOutputPath"
}
