@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo PowerPoint自動変換ツール（簡単版）
echo ========================================
echo.

REM PowerShellを使ってファイル選択ダイアログを表示
set "psScript=Add-Type -AssemblyName System.Windows.Forms; $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Filter = 'PowerPointファイル (*.pptx)|*.pptx|すべてのファイル (*.*)|*.*'; $dialog.Title = '変換するPPTXファイルを選択してください'; if ($dialog.ShowDialog() -eq 'OK') { $dialog.FileName } else { '' }"

for /f "delims=" %%i in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "%psScript%"') do set INPUT_FILE=%%i

if "%INPUT_FILE%"=="" (
    echo ファイルが選択されませんでした。
    pause
    exit /b 1
)

REM 出力ファイル名を正しく生成
for %%F in ("%INPUT_FILE%") do set OUTPUT_FILE=%%~dpnF_converted.pptx

echo.
echo 入力ファイル: %INPUT_FILE%
echo 出力ファイル: %OUTPUT_FILE%
echo.
echo 変換を開始します...
echo.

REM ライブラリがインストールされているか確認
python -c "import pptx" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo エラー: python-pptx がインストールされていません
    echo ========================================
    echo.
    echo まず、セットアップ.bat を実行してライブラリをインストールしてください。
    echo.
    pause
    exit /b 1
)

python convert_pptx.py "%INPUT_FILE%" "%OUTPUT_FILE%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 変換が完了しました！
    echo ========================================
    echo.
    echo 出力ファイル: %OUTPUT_FILE%
    echo.
    REM 出力フォルダを開く
    explorer /select,"%OUTPUT_FILE%"
) else (
    echo.
    echo ========================================
    echo エラーが発生しました
    echo ========================================
)

echo.
pause





