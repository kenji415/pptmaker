@echo off
chcp 932 >nul 2>&1
setlocal enabledelayedexpansion

echo ============================================================
echo PPTファイル一括PDF化ツール
echo ============================================================
echo.

REM 現在のディレクトリを表示
echo 現在のディレクトリ: %CD%
echo.

REM Pythonスクリプトのパス
set SCRIPT_PATH=%~dp0ppt_to_pdf_batch.py
echo Pythonスクリプトのパス: %SCRIPT_PATH%
echo.

REM Pythonスクリプトの存在確認
if not exist "%SCRIPT_PATH%" (
    echo エラー: Pythonスクリプトが見つかりません: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

REM Pythonの存在確認
python --version >nul 2>&1
if errorlevel 1 (
    echo エラー: Pythonが見つかりません。Pythonがインストールされているか確認してください。
    echo.
    pause
    exit /b 1
)

echo Pythonが見つかりました。
echo.

REM 引数でパスが指定されている場合はそれを使用、なければ入力待ち
if "%~1"=="" (
    set /p TARGET_PATH="PPTファイルが入っているフォルダのパスを入力してください: "
) else (
    set TARGET_PATH=%~1
)

REM パス内の引用符を削除
set TARGET_PATH=!TARGET_PATH:"=!

REM パスの検証
if "!TARGET_PATH!"=="" (
    echo エラー: パスが指定されていません。
    echo.
    pause
    exit /b 1
)

echo.
echo 処理を開始します...
echo 対象フォルダ: !TARGET_PATH!
echo.

REM Pythonスクリプトを実行
python "%SCRIPT_PATH%" "!TARGET_PATH!"

set PYTHON_EXIT_CODE=%ERRORLEVEL%

if %PYTHON_EXIT_CODE% neq 0 (
    echo.
    echo ============================================================
    echo エラーが発生しました。終了コード: %PYTHON_EXIT_CODE%
    echo ============================================================
    echo.
    pause
    exit /b %PYTHON_EXIT_CODE%
)

echo.
echo ============================================================
echo 処理が完了しました。
echo ============================================================
echo.
pause
