@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo PowerPoint自動変換ツール
echo ========================================
echo.

if "%~1"=="" (
    echo 使用方法:
    echo   このバッチファイルにPPTXファイルをドラッグ＆ドロップしてください
    echo   または、コマンドプロンプトから以下を実行:
    echo   convert_pptx.bat "ファイルパス.pptx"
    echo.
    pause
    exit /b 1
)

set INPUT_FILE=%~1
set OUTPUT_FILE=%~dpn1_converted.pptx

echo 入力ファイル: %INPUT_FILE%
echo 出力ファイル: %OUTPUT_FILE%
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
) else (
    echo.
    echo ========================================
    echo エラーが発生しました
    echo ========================================
)

echo.
pause


