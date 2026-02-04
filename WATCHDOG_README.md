# printviewer 監視（落ちたら復活）

## 仕組み

- **1分おき**に「printviewer のプロセスが生きているか」を確認
- 死んでいたら **起動し直す**
- 異常終了・Windows再起動後も自動復帰

## ファイル

| ファイル | 役割 |
|----------|------|
| `start_printviewer.bat` | printviewer を起動（venv があれば有効化、バックグラウンドで `python app.py`） |
| `watch_printviewer.bat` | `python.exe` がいるか確認し、いなければ `start_printviewer.bat` を実行 |

## タスクスケジューラ設定（Windows）

1. **タスクスケジューラ**を開く
2. **基本タスクの作成**（または「タスクの作成」）
3. **トリガー**
   - 「ログオン時」
   - **さらに**「1分ごとに繰り返す」を追加（トリガーを編集して「間隔: 1分」に設定）
4. **操作**
   - プログラム/スクリプト: `C:\Users\doctor\printviewer\watch_printviewer.bat`  
     （※ printviewer を置いた実際のパスに変更）
   - **開始**（作業ディレクトリ）: `C:\Users\doctor\printviewer`  
     **※ここを空にするとタスクから Python が見つからないことがあります。必ず printviewer フォルダを指定してください。**
5. **全般**
   - 「最上位の特権で実行」にチェック
   - 「ユーザーがログオンしているかどうかに関わらず実行」にチェック

これで「落ちたら次の1分以内に復活」します。

### タスクの「実行」で動かしたとき

- タスクは**窓を開かず**バックグラウンドで動きます。「実行」を押しても画面は出ません。
- 動いたか確認する方法:
  1. **ログを見る** … printviewer フォルダに `watch_printviewer.log` と `start_printviewer.log` ができます。タスク実行後に開いて「Starting printviewer」「Started app.py」などが出ているか確認。
  2. **アプリにアクセス** … ブラウザで http://127.0.0.1:5000 を開いて表示されるか確認。
- ログに「Python not found」と出る場合 … タスクの**「開始」**に printviewer フォルダ（例: `C:\Users\doctor\printviewer`）が入っているか確認してください。Anaconda 利用時は、venv なしでも `%USERPROFILE%\anaconda3\python.exe` を自動で探します。

## プロセス名を変える場合

`uvicorn` や `gunicorn`、または `printviewer.exe` など別のプロセスで動かす場合は、  
`watch_printviewer.bat` の `findstr` 対象を変更してください。

```batch
tasklist /FI "IMAGENAME eq uvicorn.exe" 2>nul | findstr /I "uvicorn.exe" >nul
```

など。
