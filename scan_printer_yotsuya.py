"""
四谷校用QRスキャン自動印刷システム（実験版）
監視フォルダを監視し、スキャンされたPDFを自動印刷
"""

import os
import csv
import shutil
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    from PIL import Image
    HAS_PYZBAR = True
except ImportError:
    HAS_PYZBAR = False
    logging.warning("pyzbar not available. Install: pip install pyzbar pillow")

try:
    from pdf2image import convert_from_path
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False
    logging.warning("pdf2image not available. Install: pip install pdf2image")

try:
    import win32print
    import win32api
    HAS_WIN32PRINT = True
except ImportError:
    HAS_WIN32PRINT = False
    logging.warning("win32print not available. Install: pip install pywin32")


# ==========================
# 設定
# ==========================
# 監視フォルダ（四谷校）
SCAN_DIR = Path(r"Y:\QS印刷\QS印刷四谷校")

# 処理済みフォルダ
PROCESSED_DIR = SCAN_DIR / "processed"
ERROR_DIR = SCAN_DIR / "error"

# 印刷対象PDFの保存場所（printviewerのPDF_DIR）
PDF_DIR = Path(r"C:\Users\doctor\printviewer\test_pdfs")

# 印刷対象PDFの中央リポジトリ（オプション）
PRINT_MATERIALS_ROOT = None  # Path(r"\\server\print_materials")  # 必要に応じて設定

# ログファイル
LOG_CSV = Path("print_log_yotsuya.csv")

# PRINT_IDとファイル名のマッピングファイル（printviewer側で生成される想定）
PRINT_ID_MAPPING_FILE = Path(r"C:\Users\doctor\printviewer\print_id_mapping.csv")

# プリンタ名（四谷校）
PRINTER_NAME = "執務室"  # FF Apeos C7070 PS H2

# PDF→画像変換設定
POPPLER_PATH = os.environ.get("POPPLER_PATH", None)
if POPPLER_PATH is None:
    default_poppler_path = r"C:\tools\poppler-25.12.0\Library\bin"
    if os.path.exists(default_poppler_path):
        POPPLER_PATH = default_poppler_path

# ファイル安定化チェック設定
STABLE_CHECK_INTERVAL_SEC = 0.5
STABLE_CHECK_COUNT = 6  # 0.5秒×6回=3秒間サイズ変化なしで安定扱い

# 処理中のファイルを追跡（二重処理防止）
processing_files = set()


# ==========================
# ログ設定
# ==========================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("scan_printer_yotsuya.log", encoding="utf-8")
    ],
)


# ==========================
# ユーティリティ
# ==========================
def wait_until_file_stable(path: Path) -> bool:
    """書き込み中ファイルを避けるため、サイズが一定になるまで待つ"""
    last = -1
    stable = 0
    for _ in range(STABLE_CHECK_COUNT * 5):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return False

        if size == last and size > 0:
            stable += 1
            if stable >= STABLE_CHECK_COUNT:
                return True
        else:
            stable = 0
            last = size

        time.sleep(STABLE_CHECK_INTERVAL_SEC)

    return False


def load_print_id_mapping() -> dict:
    """PRINT_IDとファイル名のマッピングを読み込む"""
    mapping = {}
    if PRINT_ID_MAPPING_FILE.exists():
        try:
            with open(PRINT_ID_MAPPING_FILE, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    print_id = row.get("print_id", "").strip()
                    filename = row.get("filename", "").strip()
                    if print_id and filename:
                        mapping[print_id] = filename
            logging.info(f"PRINT_IDマッピングを読み込みました: {len(mapping)}件")
        except Exception as e:
            logging.warning(f"マッピングファイルの読み込みエラー: {e}")
    return mapping


def get_print_pdf_path(print_id: str, original_filename: Optional[str] = None) -> Optional[Path]:
    """PRINT_IDに対応するPDFファイルのパスを取得"""
    
    # 方法1: マッピングファイルから検索
    mapping = load_print_id_mapping()
    if print_id in mapping:
        mapped_filename = mapping[print_id]
        pdf_path = PDF_DIR / mapped_filename
        if pdf_path.exists():
            logging.info(f"印刷対象PDFを発見（マッピング）: {pdf_path}")
            return pdf_path
        else:
            logging.warning(f"マッピングに記載されているファイルが見つかりません: {mapped_filename}")
    
    # 方法2: 元のファイル名が提供されている場合、そのファイルを検索
    if original_filename:
        # ファイル名から拡張子を除いた部分で検索
        base_name = Path(original_filename).stem
        for pdf_file in PDF_DIR.rglob(f"*{base_name}*.pdf"):
            if pdf_file.is_file():
                logging.info(f"印刷対象PDFを発見（元ファイル名）: {pdf_file}")
                return pdf_file
    
    # 方法3: {PRINT_ID}.pdf という名前のファイルを検索
    if PDF_DIR.exists():
        for pdf_file in PDF_DIR.rglob(f"{print_id}.pdf"):
            if pdf_file.is_file():
                logging.info(f"印刷対象PDFを発見（PRINT_ID名）: {pdf_file}")
                return pdf_file
    
    # 方法4: 中央リポジトリから検索
    if PRINT_MATERIALS_ROOT and PRINT_MATERIALS_ROOT.exists():
        pdf_path = PRINT_MATERIALS_ROOT / f"{print_id}.pdf"
        if pdf_path.exists():
            logging.info(f"印刷対象PDFを発見（中央リポジトリ）: {pdf_path}")
            return pdf_path
    
    # 方法5: test_pdfsフォルダ内のすべてのPDFをリストアップしてログに出力
    if PDF_DIR.exists():
        all_pdfs = list(PDF_DIR.rglob("*.pdf"))
        logging.warning(f"印刷対象PDFが見つかりません: {print_id}.pdf")
        logging.info(f"test_pdfsフォルダ内のPDFファイル数: {len(all_pdfs)}")
        if len(all_pdfs) <= 10:
            logging.info(f"利用可能なPDFファイル: {[str(p.relative_to(PDF_DIR)) for p in all_pdfs]}")
    
    return None


def extract_print_id_from_qr(pdf_path: Path) -> tuple[Optional[str], Optional[str]]:
    """PDFの1ページ目からQRコードを読み取り、PRINT_IDを抽出"""
    if not HAS_PDF2IMAGE or not HAS_PYZBAR:
        logging.error("必要なライブラリがインストールされていません")
        return None
    
    try:
        # PDFの1ページ目を画像に変換
        images = convert_from_path(
            str(pdf_path),
            first_page=1,
            last_page=1,
            poppler_path=POPPLER_PATH
        )
        
        if not images:
            logging.warning(f"PDFから画像を取得できませんでした: {pdf_path}")
            return None
        
        # QRコードを検出
        for image in images:
            # pyzbarでQRコードを検出（PIL Imageをそのまま使用）
            qr_codes = pyzbar_decode(image)
            
            if not qr_codes:
                continue
            
            # QRコードが複数ある場合はエラー
            if len(qr_codes) > 1:
                logging.warning(f"QRコードが複数検出されました: {pdf_path}")
                return None
            
            # QRコードデータを取得
            qr_data = qr_codes[0].data.decode('utf-8')
            logging.info(f"QRコード検出: {qr_data}")
            
            # PRINT_ID=QS_...,FILE=ファイル名 の形式から情報を抽出
            import re
            print_id = None
            original_filename = None
            
            # PRINT_IDを抽出
            match = re.search(r'PRINT_ID=([^,]+)', qr_data)
            if match:
                print_id = match.group(1)
            
            # FILEを抽出
            match = re.search(r'FILE=([^,]+)', qr_data)
            if match:
                original_filename = match.group(1)
            
            # 旧形式（PRINT_ID=のみ）にも対応
            if not print_id:
                match = re.match(r'PRINT_ID=(\S+)', qr_data)
                if match:
                    print_id = match.group(1)
            
            if print_id:
                return print_id, original_filename
            else:
                # PRINT_IDがなければエラー
                return None, None
        
        logging.warning(f"QRコードが検出されませんでした: {pdf_path}")
        return None, None
        
    except Exception as e:
        logging.exception(f"QRコード読み取りエラー: {pdf_path}, {e}")
        return None, None


def print_pdf(pdf_path: Path, printer_name: str, copies: int = 1) -> bool:
    """PDFをWindows印刷キューに送信"""
    if not HAS_WIN32PRINT:
        logging.error("win32printが利用できません")
        return False
    
    try:
        # プリンタが存在するか確認
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        printers_network = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_NETWORK)]
        all_printers = printers + printers_network
        
        if printer_name not in all_printers:
            logging.error(f"プリンタが見つかりません: {printer_name}")
            logging.info(f"利用可能なプリンタ: {all_printers}")
            return False
        
        # ネットワークパス（UNCパス）の場合は一時的にローカルにコピー
        temp_pdf_path = None
        try:
            source_path = str(pdf_path.resolve())
            
            # UNCパス（\\で始まる）の場合は一時ファイルにコピー
            if source_path.startswith('\\\\'):
                import tempfile
                temp_dir = Path(tempfile.gettempdir())
                temp_pdf_path = temp_dir / f"print_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{pdf_path.name}"
                
                logging.info(f"ネットワークパスを検出。一時ファイルにコピー: {temp_pdf_path}")
                shutil.copy2(source_path, str(temp_pdf_path))
                abs_path = str(temp_pdf_path)
            else:
                abs_path = source_path
            
            logging.info(f"印刷を試行中: {abs_path} -> {printer_name}")
            
            # 方法1: printtoコマンドを使用
            try:
                result = win32api.ShellExecute(
                    0,
                    "printto",
                    abs_path,
                    f'"{printer_name}"',
                    str(Path(abs_path).parent),
                    0
                )
                
                # ShellExecuteは成功時に32以上の値を返す
                if result > 32:
                    logging.info(f"印刷ジョブを投入しました: {pdf_path.name} -> {printer_name} (部数: {copies})")
                    # 一時ファイルは印刷キューに送信後、少し待ってから削除
                    if temp_pdf_path:
                        import threading
                        def delete_temp_file():
                            time.sleep(10)  # 10秒待ってから削除
                            try:
                                if temp_pdf_path.exists():
                                    temp_pdf_path.unlink()
                                    logging.info(f"一時ファイルを削除: {temp_pdf_path}")
                            except Exception as e:
                                logging.warning(f"一時ファイル削除エラー: {e}")
                        threading.Thread(target=delete_temp_file, daemon=True).start()
                    return True
                else:
                    logging.warning(f"ShellExecuteの戻り値: {result}")
                    raise Exception(f"ShellExecute returned {result}")
                    
            except Exception as e1:
                logging.warning(f"printtoでの印刷に失敗: {e1}")
                logging.info("代替方法を試行中...")
                
                # 方法2: デフォルトプリンタを一時的に変更して印刷
                try:
                    # 現在のデフォルトプリンタを取得
                    default_printer = win32print.GetDefaultPrinter()
                    logging.info(f"現在のデフォルトプリンタ: {default_printer}")
                    
                    # 一時的にデフォルトプリンタを変更
                    win32print.SetDefaultPrinter(printer_name)
                    logging.info(f"デフォルトプリンタを変更: {printer_name}")
                    
                    # PDFを印刷（デフォルトプリンタに送信）
                    result = win32api.ShellExecute(
                        0,
                        "print",
                        abs_path,
                        None,  # デフォルトプリンタを使用
                        str(Path(abs_path).parent),
                        0
                    )
                    
                    # 元のデフォルトプリンタに戻す
                    win32print.SetDefaultPrinter(default_printer)
                    
                    if result > 32:
                        logging.info(f"印刷ジョブを投入しました: {pdf_path.name} -> {printer_name} (部数: {copies})")
                        # 一時ファイルを削除
                        if temp_pdf_path:
                            import threading
                            def delete_temp_file():
                                time.sleep(10)
                                try:
                                    if temp_pdf_path.exists():
                                        temp_pdf_path.unlink()
                                        logging.info(f"一時ファイルを削除: {temp_pdf_path}")
                                except Exception as e:
                                    logging.warning(f"一時ファイル削除エラー: {e}")
                            threading.Thread(target=delete_temp_file, daemon=True).start()
                        return True
                    else:
                        logging.error(f"印刷に失敗しました。ShellExecute戻り値: {result}")
                        return False
                        
                except Exception as e2:
                    logging.exception(f"代替方法での印刷も失敗: {e2}")
                    # 元のデフォルトプリンタに戻す（念のため）
                    try:
                        if 'default_printer' in locals():
                            win32print.SetDefaultPrinter(default_printer)
                    except:
                        pass
                    return False
                    
        finally:
            # エラー時も一時ファイルをクリーンアップ（ただし印刷キューに送信済みの場合は少し待つ）
            pass  # 削除は別スレッドで行う
        
    except Exception as e:
        logging.exception(f"印刷エラー: {pdf_path}, {e}")
        # 一時ファイルを削除
        if temp_pdf_path and temp_pdf_path.exists():
            try:
                temp_pdf_path.unlink()
            except:
                pass
        return False


def log_print_result(
    scan_file: str,
    print_id: Optional[str],
    printer: str,
    result: str,
    error_message: str = ""
):
    """印刷結果をCSVログに記録"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # ログファイルが存在しない場合はヘッダーを書き込む
    file_exists = LOG_CSV.exists()
    
    with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow([
                "timestamp", "scan_file", "print_id",
                "printer", "result", "error_message"
            ])
        writer.writerow([
            timestamp, scan_file, print_id or "",
            printer, result, error_message
        ])


def handle_pdf(pdf_path: Path) -> None:
    """PDF1つの処理：安定化待ち→QR読取→印刷→ログ記録"""
    # 既に処理中のファイルはスキップ（二重処理防止）
    file_key = str(pdf_path)
    if file_key in processing_files:
        logging.info(f"処理中のファイルをスキップ: {pdf_path}")
        return
    
    processing_files.add(file_key)
    
    try:
        # ファイルがPDFか確認
        if pdf_path.suffix.lower() != ".pdf":
            return
        
        logging.info(f"PDF検出: {pdf_path}")
        
        # ファイル安定化待ち
        if not wait_until_file_stable(pdf_path):
            logging.warning(f"ファイルが安定しませんでした: {pdf_path}")
            return
        
        # QRコードからPRINT_IDと元のファイル名を抽出
        print_id, original_filename = extract_print_id_from_qr(pdf_path)
        
        if not print_id:
            # QRコードが読めない場合、errorフォルダへ
            ERROR_DIR.mkdir(parents=True, exist_ok=True)
            error_file = ERROR_DIR / pdf_path.name
            
            # 同名ファイルがある場合はタイムスタンプを追加
            if error_file.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = error_file.stem
                suffix = error_file.suffix
                error_file = ERROR_DIR / f"{stem}_{timestamp}{suffix}"
            
            try:
                shutil.move(str(pdf_path), str(error_file))
                log_print_result(
                    scan_file=pdf_path.name,
                    print_id=None,
                    printer=PRINTER_NAME,
                    result="error",
                    error_message="QR not found"
                )
                logging.warning(f"QRコードが読めませんでした: {pdf_path}")
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
            return
        
        # PRINT_IDに対応する印刷対象PDFを取得
        print_pdf_path = get_print_pdf_path(print_id, original_filename)
        
        if not print_pdf_path:
            # 印刷対象PDFが見つからない場合、errorフォルダへ
            ERROR_DIR.mkdir(parents=True, exist_ok=True)
            error_file = ERROR_DIR / pdf_path.name
            
            # 同名ファイルがある場合はタイムスタンプを追加
            if error_file.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = error_file.stem
                suffix = error_file.suffix
                error_file = ERROR_DIR / f"{stem}_{timestamp}{suffix}"
            
            try:
                shutil.move(str(pdf_path), str(error_file))
                log_print_result(
                    scan_file=pdf_path.name,
                    print_id=print_id,
                    printer=PRINTER_NAME,
                    result="error",
                    error_message=f"PDF not found: {print_id}.pdf"
                )
                logging.error(f"印刷対象PDFが見つかりません: {print_id}.pdf")
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
            return
        
        # 印刷対象PDFを印刷（部数は1部固定）
        print_success = print_pdf(print_pdf_path, PRINTER_NAME, 1)
        
        if print_success:
            # 成功時、processedフォルダへ移動
            PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
            processed_file = PROCESSED_DIR / pdf_path.name
            
            # 同名ファイルがある場合はタイムスタンプを追加
            if processed_file.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = processed_file.stem
                suffix = processed_file.suffix
                processed_file = PROCESSED_DIR / f"{stem}_{timestamp}{suffix}"
            
            try:
                shutil.move(str(pdf_path), str(processed_file))
                log_print_result(
                    scan_file=pdf_path.name,
                    print_id=print_id,
                    printer=PRINTER_NAME,
                    result="success",
                    error_message=""
                )
                logging.info(f"処理完了: {pdf_path.name} -> {print_id or 'QRなし'}")
            except Exception as e:
                logging.error(f"processedフォルダへの移動エラー: {e}")
        else:
            # 印刷失敗時、errorフォルダへ
            ERROR_DIR.mkdir(parents=True, exist_ok=True)
            error_file = ERROR_DIR / pdf_path.name
            
            # 同名ファイルがある場合はタイムスタンプを追加
            if error_file.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = error_file.stem
                suffix = error_file.suffix
                error_file = ERROR_DIR / f"{stem}_{timestamp}{suffix}"
            
            try:
                shutil.move(str(pdf_path), str(error_file))
                log_print_result(
                    scan_file=pdf_path.name,
                    print_id=print_id,
                    printer=PRINTER_NAME,
                    result="error",
                    error_message="Print failed"
                )
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
    
    except Exception as e:
        logging.exception(f"処理エラー: {pdf_path}, {e}")
        # エラー時もerrorフォルダへ移動を試みる
        try:
            ERROR_DIR.mkdir(parents=True, exist_ok=True)
            error_file = ERROR_DIR / pdf_path.name
            
            if pdf_path.exists():
                if error_file.exists():
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    stem = error_file.stem
                    suffix = error_file.suffix
                    error_file = ERROR_DIR / f"{stem}_{timestamp}{suffix}"
                shutil.move(str(pdf_path), str(error_file))
        except Exception:
            pass
    finally:
        # 処理中リストから削除
        processing_files.discard(file_key)


# ==========================
# Watchdog ハンドラ
# ==========================
class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() == ".pdf":
            # 処理は別スレッドで実行（非ブロッキング）
            import threading
            threading.Thread(target=self._handle_pdf_delayed, args=(path,), daemon=True).start()
    
    def _handle_pdf_delayed(self, path: Path):
        """少し待ってから処理（ファイル作成完了を待つ）"""
        time.sleep(0.3)
        handle_pdf(path)
    
    def on_moved(self, event):
        if event.is_directory:
            return
        path = Path(event.dest_path)
        if path.suffix.lower() == ".pdf":
            import threading
            threading.Thread(target=self._handle_pdf_delayed, args=(path,), daemon=True).start()


def main():
    """メイン処理"""
    try:
        logging.info("四谷校用QRスキャン自動印刷システム（実験版）を起動します")
        print("=" * 60)
        print("四谷校用QRスキャン自動印刷システム（実験版）")
        print("=" * 60)
        
        # 設定確認
        print(f"監視フォルダを確認: {SCAN_DIR}")
        if not SCAN_DIR.exists():
            error_msg = f"監視フォルダが存在しません: {SCAN_DIR}"
            logging.error(error_msg)
            print(f"ERROR: {error_msg}")
            print("\nフォルダを作成するか、パスを確認してください。")
            input("\nEnterキーを押して終了してください...")
            return
        
        print(f"✓ 監視フォルダ: {SCAN_DIR}")
        
        if not HAS_PDF2IMAGE:
            error_msg = "pdf2imageがインストールされていません: pip install pdf2image"
            logging.error(error_msg)
            print(f"ERROR: {error_msg}")
            input("\nEnterキーを押して終了してください...")
            return
        print("✓ pdf2image: OK")
        
        if not HAS_PYZBAR:
            logging.warning("pyzbarがインストールされていません: pip install pyzbar pillow (QRコード読み取りができません)")
            print("WARNING: pyzbarがインストールされていません（QRコード読み取りができません）")
        else:
            print("✓ pyzbar: OK")
        
        if not HAS_WIN32PRINT:
            error_msg = "pywin32がインストールされていません: pip install pywin32"
            logging.error(error_msg)
            print(f"ERROR: {error_msg}")
            input("\nEnterキーを押して終了してください...")
            return
        print("✓ pywin32: OK")
    
        # フォルダを作成
        PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
        ERROR_DIR.mkdir(parents=True, exist_ok=True)
        
        logging.info(f"監視フォルダ: {SCAN_DIR}")
        logging.info(f"処理済みフォルダ: {PROCESSED_DIR}")
        logging.info(f"エラーフォルダ: {ERROR_DIR}")
        logging.info(f"プリンタ: {PRINTER_NAME}")
        
        print(f"\n✓ 処理済みフォルダ: {PROCESSED_DIR}")
        print(f"✓ エラーフォルダ: {ERROR_DIR}")
        print(f"✓ プリンタ: {PRINTER_NAME}")
        print(f"✓ PDF検索フォルダ: {PDF_DIR}")
        if not PDF_DIR.exists():
            print(f"WARNING: PDF検索フォルダが存在しません: {PDF_DIR}")
        if PRINT_MATERIALS_ROOT:
            print(f"✓ 中央リポジトリ: {PRINT_MATERIALS_ROOT}")
            if not PRINT_MATERIALS_ROOT.exists():
                print(f"WARNING: 中央リポジトリが存在しません: {PRINT_MATERIALS_ROOT}")
        
        # プリンタの存在確認
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
            printers_network = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_NETWORK)]
            all_printers = printers + printers_network
            
            if PRINTER_NAME not in all_printers:
                print(f"\nWARNING: プリンタ '{PRINTER_NAME}' が見つかりませんでした。")
                print(f"利用可能なプリンタ:")
                for p in all_printers:
                    print(f"  - {p}")
                print(f"\nプリンタ名を確認して、スクリプト内の PRINTER_NAME を修正してください。")
            else:
                print(f"✓ プリンタ '{PRINTER_NAME}' が見つかりました")
        except Exception as e:
            logging.warning(f"プリンタ確認エラー: {e}")
            print(f"WARNING: プリンタの確認中にエラーが発生しました: {e}")
        
        print("\n" + "=" * 60)
        print("フォルダ監視を開始しました。")
        print("停止するには Ctrl+C を押してください。")
        print("=" * 60 + "\n")
        
        # 監視開始
        observer = Observer()
        observer.schedule(PDFHandler(), str(SCAN_DIR), recursive=False)
        observer.start()
        logging.info("フォルダ監視を開始しました")
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n停止中...")
            logging.info("停止中...")
            observer.stop()
        
        observer.join()
        print("プログラムを終了しました。")
        
    except Exception as e:
        import traceback
        error_msg = f"予期しないエラーが発生しました: {e}"
        logging.exception(error_msg)
        print(f"\nERROR: {error_msg}")
        print("\n詳細:")
        traceback.print_exc()
        input("\nEnterキーを押して終了してください...")


if __name__ == "__main__":
    main()

