$ErrorActionPreference = "Stop"

# QS進捗シート同期 実行スクリプト
# - 環境変数をセットして sync_qs_list_to_sheets.py を実行します
# - credentials.json の中身はリポジトリに入れないでください（パス参照のみ）

# ===== 設定（必要に応じて変更） =====
$env:GOOGLE_SERVICE_ACCOUNT_JSON = 'C:\Users\user\credentials.json'
$env:QS_SPREADSHEET_ID           = '11wIE0lmXQv_0hnstFDKhj1X5NGfQHALf06-wvyWeuKU'
$env:QS_HEADER_ROW               = '1'
$env:QS_TEXT_MAPPING_CSV         = 'C:\Users\doctor\printviewer\text_mapping.csv'

# ===== 実行 =====
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# PDF_DIR は日本語パスになることがあるため、UTF-8で保存された pdf_dir.txt から読み取る
$pdfDirFile = Join-Path $scriptDir 'pdf_dir.txt'
if (-not (Test-Path $pdfDirFile)) {
  throw "pdf_dir.txt not found: $pdfDirFile"
}
$env:QS_PDF_DIR = (Get-Content -Encoding UTF8 $pdfDirFile | Select-Object -First 1).Trim()
if (-not $env:QS_PDF_DIR) {
  throw "QS_PDF_DIR is empty (check pdf_dir.txt)"
}

$logPath = Join-Path $scriptDir 'sync_task.log'
"START $(Get-Date -Format s) user=$env:USERNAME cwd=$scriptDir" | Out-File -Append -Encoding utf8 $logPath
"QS_PDF_DIR=$env:QS_PDF_DIR" | Out-File -Append -Encoding utf8 $logPath
try {
  python .\sync_qs_list_to_sheets.py *>> $logPath
  "PYTHON_OK exit=$LASTEXITCODE" | Out-File -Append -Encoding utf8 $logPath
} catch {
  "PYTHON_EXCEPTION $($_.Exception.Message)" | Out-File -Append -Encoding utf8 $logPath
}
"END   $(Get-Date -Format s)" | Out-File -Append -Encoding utf8 $logPath

