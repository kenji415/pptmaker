## 社内PDF → Google Sheets 同期

### 目的

社内の `PDF_DIR\算数\**\*.pdf` を走査し、Google スプレッドシート `QSリスト` タブへ同期します。

- **A列**: 相対パス（キー）例: `算数/6年/xxx.pdf`
- **B列**: タイトル（ファイル名から `.pdf` を除去）
- **C列**: 修正内容（fix系タグを `/` で連結）
- **D列**: その他タグ（通常タグを `/` で連結）
- **E列以降は手入力用**（スクリプトは触りません）
- 新規PDF: 行追加 / 削除PDF: 行削除

タグは `text_mapping.csv`（printviewerと同じ形式）から読みます（未指定なら空でOK）。

### 事前準備（1回だけ）

1. Google Cloud Console で Service Account を作成
2. Sheets API を有効化
3. JSON キーを発行して社内PCに保存
4. 対象スプレッドシートを Service Account のメールアドレスに「編集可」で共有

### 実行

PowerShell 例:

```powershell
$env:QS_SPREADSHEET_ID="11wIE0lmXQv_0hnstFDKhj1X5NGfQHALf06-wvyWeuKU"
$env:QS_SHEET_NAME="QSリスト"
$env:QS_HEADER_ROW="1"  # 見出し行（相対パス/タイトル/修正内容/その他タグ が入っている行）
$env:QS_PDF_DIR="X:\算数科作業フォルダ\QS\test_pdf"
$env:QS_TEXT_MAPPING_CSV="C:\Users\doctor\printviewer\text_mapping.csv"
$env:GOOGLE_SERVICE_ACCOUNT_JSON="C:\path\to\service-account.json"

python .\sync_qs_list_to_sheets.py
```

### 自動更新（推奨）

Windows タスクスケジューラで `python sync_qs_list_to_sheets.py` を 5分ごと等で実行してください。

