@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo PowerPoint自動変換ツール セットアップ
echo ========================================
echo.
echo 必要なライブラリをインストールします...
echo.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo セットアップが完了しました！
    echo ========================================
    echo.
    echo これで convert_pptx_簡単.bat が使えます。
) else (
    echo.
    echo ========================================
    echo エラーが発生しました
    echo ========================================
    echo.
    echo Pythonが正しくインストールされているか確認してください。
)

echo.
pause





