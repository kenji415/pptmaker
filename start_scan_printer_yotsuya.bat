@echo off
chcp 65001 > nul
echo 四谷校用QRスキャン自動印刷システムを起動します...
echo.

cd /d "%~dp0"

REM Python環境を確認
where python >nul 2>&1
if errorlevel 1 (
    echo Pythonが見つかりません。Anaconda環境をアクティベートしてください。
    pause
    exit /b 1
)

REM 必要なモジュールを確認してインストール
echo 必要なモジュールを確認中...

python -c "import watchdog" >nul 2>&1
if errorlevel 1 (
    echo watchdogモジュールをインストール中...
    pip install watchdog
)

python -c "import pdf2image" >nul 2>&1
if errorlevel 1 (
    echo pdf2imageモジュールをインストール中...
    pip install pdf2image
)

python -c "import pyzbar" >nul 2>&1
if errorlevel 1 (
    echo pyzbarモジュールをインストール中...
    pip install pyzbar pillow
)

python -c "import win32print" >nul 2>&1
if errorlevel 1 (
    echo pywin32モジュールをインストール中...
    pip install pywin32
)

echo.
echo モジュール確認完了。スクリプトを起動します...
echo.

python scan_printer_yotsuya.py

if errorlevel 1 (
    echo.
    echo エラーが発生しました。
    echo 詳細は上記のエラーメッセージを確認してください。
    echo.
    pause
) else (
    echo.
    echo プログラムが正常に終了しました。
    pause
)

