# スクリプト名: ExtractTablesFromPdfs.ps1

param (
    [string[]]$pdfPaths,
    [string]$outputDir,
    [switch]$forceOverwrite
)

function Show-Help {
    Write-Output "Usage: .\ExtractTablesFromPdfs.ps1 -pdfPaths <PDF file paths or directories> [-outputDir <Output directory>] [-forceOverwrite]"
    Write-Output ""
    Write-Output "Arguments:"
    Write-Output "  -pdfPaths        : Paths to the input PDF files or directories containing PDF files. Supports wildcards."
    Write-Output "  -outputDir       : (Optional) Path to the output directory for Excel files. If not specified, the output will be in the same directory as the PDF files."
    Write-Output "  -forceOverwrite  : (Optional) Force overwrite existing Excel files."
}

# コマンドライン引数の検証
if (-not $pdfPaths) {
    Write-Error "Please provide the PDF file paths or directories as arguments."
    Show-Help
    exit
}

# スクリプトのディレクトリを取得
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# ExtractTablesFromPdf.ps1 の絶対パスを取得
$extractScriptPath = Join-Path -Path $scriptDir -ChildPath "ExtractTablesFromPdf.ps1"

# PDFファイルリストの収集
$pdfFiles = @()
foreach ($path in $pdfPaths) {
    if (Test-Path $path -PathType Container) {
        $pdfFiles += Get-ChildItem -Path $path -Recurse -Include *.pdf
    } else {
        $pdfFiles += Get-ChildItem -Path $path -Include *.pdf
    }
}

# 同名のExcelファイルが存在するかチェック
if (-not $forceOverwrite) {
    foreach ($pdfFile in $pdfFiles) {
        $fileName = $pdfFile.Name
        $outputDirectory = if ($outputDir) { $outputDir } else { $pdfFile.DirectoryName }
        $excelOutputPath = Join-Path -Path $outputDirectory -ChildPath ([System.IO.Path]::ChangeExtension($fileName, ".xlsx"))

        if (Test-Path $excelOutputPath) {
            Write-Error "The output file already exists: $excelOutputPath. Use -forceOverwrite to overwrite existing files. Script execution aborted."
            exit
        }
    }
}

# 出力ディレクトリの作成（必要な場合）
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# 各PDFファイルを処理
foreach ($pdfFile in $pdfFiles) {
    $fileName = $pdfFile.Name
    $outputDirectory = if ($outputDir) { $outputDir } else { $pdfFile.DirectoryName }
    $excelOutputPath = Join-Path -Path $outputDirectory -ChildPath ([System.IO.Path]::ChangeExtension($fileName, ".xlsx"))

    & $extractScriptPath -pdfFileName $pdfFile.FullName -excelFilePath $excelOutputPath
    Write-Output "Converted $fileName to $excelOutputPath"
}
