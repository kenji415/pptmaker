# QRスキャン自動印刷システム（方式C） 使用方法

## 概要

校舎別スキャンフォルダに保存されたPDFからQRコードを読み取り、対応する印刷用PDFを自動で印刷するシステムです。

## ディレクトリ構成

```
\\server\scan\
 ├─ yotsuya\
 │   ├─ in          # 複合機のスキャン保存先（監視対象）
 │   ├─ processing  # 処理中
 │   ├─ done        # 成功
 │   └─ error       # QR不明・印刷失敗
 ├─ jiyugaoka\
 │   ├─ in
 │   ├─ processing
 │   ├─ done
 │   └─ error
 └─ kichijoji\
     ├─ in
     ├─ processing
     ├─ done
     └─ error
```

## 設定

### 1. プリンタ設定（printers.yaml）

```yaml
yotsuya:
  printer_name: "RICOH_MP_C6004_YOTSUYA"
  max_copies: 5

jiyugaoka:
  printer_name: "Canon_iR_ADV_JIYUGAOKA"
  max_copies: 5

kichijoji:
  printer_name: "EPSON_BIZ_KICHIJOJI"
  max_copies: 5
```

### 2. 環境変数（オプション）

- `POPPLER_PATH`: popplerのパス（未設定時は `C:\tools\poppler-25.12.0\Library\bin` を使用）

## インストール

```bash
pip install -r requirements.txt
```

## 実行方法

```bash
python scan_printer.py
```

## 処理フロー

1. 各校舎の `in` フォルダを監視
2. 新規PDFファイルを検知
3. `processing` フォルダに移動（二重処理防止）
4. PDFの1ページ目からQRコードを読み取り
5. `PRINT_ID=QS_YYYY_NNNNN` 形式からPRINT_IDを抽出
6. `\\server\print_materials\{PRINT_ID}.pdf` を検索
7. フォルダ名から校舎を判定
8. `printers.yaml` からプリンタ名を取得
9. 印刷ジョブを投入
10. ログをCSVに記録
11. 成功時は `done`、失敗時は `error` フォルダに移動

## ログ

印刷ログは `print_log.csv` に記録されます。

フォーマット：
```csv
timestamp,campus,scan_file,print_id,printer,result,error_message
```

例：
```csv
2025-12-11 14:23:01,yotsuya,scan_001.pdf,QS_2025_03421,RICOH_MP_C6004_YOTSUYA,success,
2025-12-11 14:25:12,jiyugaoka,scan_002.pdf,,Canon_iR_ADV_JIYUGAOKA,error,QR not found
```

## 注意事項

- QRコードはPDFの1ページ目のみを読み取ります
- QRコードが複数検出された場合はエラーとして処理されます
- 同一ファイルの再検知は自動的にスキップされます
- 印刷対象PDFが見つからない場合、`error` フォルダに移動します

## トラブルシューティング

### QRコードが読み取れない場合

- PDFの1ページ目にQRコードが存在するか確認
- QRコードのサイズや品質を確認
- ログファイル（scan_printer.log）を確認

### プリンタが見つからない場合

- `printers.yaml` の `printer_name` が正しいか確認
- Windowsのプリンタ一覧でプリンタ名を確認
- ネットワークプリンタの場合は接続を確認

### 印刷が実行されない場合

- `pywin32` が正しくインストールされているか確認
- Windowsの印刷キューを確認
- ログファイルを確認



