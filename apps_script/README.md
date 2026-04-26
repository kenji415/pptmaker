## Googleスプレッドシート「QSリスト」同期（Apps Script）

### 目的

Google Drive 上の「算数」配下の PDF を走査し、スプレッドシートの `QSリスト` タブへ同期します。

- **A列**: PDF相対パス（例: `算数/6年/xxx.pdf`）
- **B列**: タイトル（ファイル名の `.pdf` を除去）
- **C列**: 修正内容（fix系タグを `/` で連結）
- **D列**: その他タグ（通常タグを `/` で連結）

**E列以降は手入力用**として残し、スクリプトは触りません。

### 手順

1. スプレッドシートを開く
2. `拡張機能` → `Apps Script`
3. 新規プロジェクトに `apps_script/qs_sync.gs` の内容を貼り付け
4. `CONFIG` を設定
   - `DRIVE_ROOT_FOLDER_ID`: Drive上のルートフォルダID（算数フォルダ、またはその親）
   - `DRIVE_SCAN_SUBPATH`: 例 `算数`
   - `TEXT_MAPPING_FILE_ID`（任意）: `text_mapping.csv` をDriveに置いた場合のファイルID
5. `syncQsList()` を実行（初回は権限の許可が必要）
6. `トリガー` で時間主導（例: 5分ごと）を設定すると自動更新になります

### タグの入力元

タグも埋めたい場合は、Drive上に `text_mapping.csv` を置き、`TEXT_MAPPING_FILE_ID` を指定してください。

※ `file_path` が `算数/6年/xxx.pdf` のような相対パスになっている必要があります。

