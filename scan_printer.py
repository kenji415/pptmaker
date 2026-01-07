"""
QRスキャン自動印刷システム（方式C）
校舎別スキャンフォルダを監視し、QRコードからPRINT_IDを読み取り、
対応するPDFを自動印刷する
"""

import os
import re
import csv
import shutil
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, Union

import yaml
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
# 校舎別スキャンフォルダのルート（UNCパス対応）
SCAN_ROOT = Path(r"\\server\scan")

# 印刷用PDFの中央リポジトリ
PRINT_MATERIALS_ROOT = Path(r"\\server\print_materials")

# 設定ファイル
PRINTERS_CONFIG = Path("printers.yaml")

# ログファイル
LOG_CSV = Path("print_log.csv")

# 処理中のファイルを追跡（二重処理防止）
processing_files = set()

# ファイル安定化チェック設定
STABLE_CHECK_INTERVAL_SEC = 0.5
STABLE_CHECK_COUNT = 6  # 0.5秒×6回=3秒間サイズ変化なしで安定扱い

# PDF→画像変換設定
POPPLER_PATH = os.environ.get("POPPLER_PATH", None)
if POPPLER_PATH is None:
    default_poppler_path = r"C:\tools\poppler-25.12.0\Library\bin"
    if os.path.exists(default_poppler_path):
        POPPLER_PATH = default_poppler_path


# ==========================
# ログ設定
# ==========================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("scan_printer.log", encoding="utf-8")
    ],
)


# ==========================
# ユーティリティ
# ==========================
def load_printer_config() -> dict:
    """プリンタ設定を読み込む"""
    if not PRINTERS_CONFIG.exists():
        logging.error(f"プリンタ設定ファイルが見つかりません: {PRINTERS_CONFIG}")
        return {}
    
    try:
        with open(PRINTERS_CONFIG, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            return config or {}
    except Exception as e:
        logging.error(f"プリンタ設定ファイルの読み込みエラー: {e}")
        return {}


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


def extract_print_id_from_qr(pdf_path: Path) -> Tuple[Optional[str], Optional[str]]:
    """PDFの1ページ目からQRコードを読み取り、PRINT_IDとプリンター名を抽出
    戻り値: (print_id, printer_name)"""
    if not HAS_PDF2IMAGE or not HAS_PYZBAR:
        logging.error("必要なライブラリがインストールされていません")
        return None, None
    
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
            return None, None
        
        # QRコードを検出
        for image in images:
            # pyzbarでQRコードを検出（PIL Imageをそのまま使用）
            qr_codes = pyzbar_decode(image)
            
            if not qr_codes:
                continue
            
            # QRコードが複数ある場合はエラー
            if len(qr_codes) > 1:
                logging.warning(f"QRコードが複数検出されました: {pdf_path}")
                return None, None
            
            # QRコードデータを取得
            qr_data = qr_codes[0].data.decode('utf-8')
            logging.info(f"QRコード検出: {qr_data}")
            
            print_id = None
            printer_name = None
            
            # PRINT_ID=QS_... の形式からPRINT_IDを抽出
            match = re.search(r'PRINT_ID=([^,\s]+)', qr_data)
            if match:
                print_id = match.group(1).strip()
                logging.info(f"PRINT_ID抽出: {print_id}")
            
            # PRINTER=プリンター名 の形式からプリンター名を抽出
            match = re.search(r'PRINTER=([^,\s]+)', qr_data)
            if match:
                printer_name = match.group(1).strip()
                logging.info(f"PRINTER抽出: {printer_name}")
            
            return print_id, printer_name
        
        logging.warning(f"QRコードが検出されませんでした: {pdf_path}")
        return None, None
        
    except Exception as e:
        logging.exception(f"QRコード読み取りエラー: {pdf_path}, {e}")
        return None, None


def get_print_pdf_path(print_id: str) -> Optional[Path]:
    """PRINT_IDに対応するPDFファイルのパスを取得"""
    pdf_path = PRINT_MATERIALS_ROOT / f"{print_id}.pdf"
    
    if pdf_path.exists():
        return pdf_path
    else:
        logging.warning(f"印刷対象PDFが見つかりません: {pdf_path}")
        return None


def get_campus_from_folder_path(file_path: Path) -> Optional[str]:
    """ファイルパスから校舎名（フォルダ名）を取得"""
    # \\server\scan\yotsuya\in\file.pdf のようなパスから yotsuya を抽出
    parts = file_path.parts
    
    # SCAN_ROOTのパス部分を除去
    try:
        scan_root_parts = Path(SCAN_ROOT).parts
        # UNCパスの場合、\\server\share は parts が ['\\\\server\\share'] になることがある
        if len(parts) > len(scan_root_parts):
            # 校舎名は SCAN_ROOT の直下のフォルダ
            campus_index = len(scan_root_parts)
            if campus_index < len(parts):
                return parts[campus_index]
    except Exception:
        pass
    
    return None


def print_pdf(pdf_path: Path, printer_name: str, copies: int = 1) -> bool:
    """PDFをWindows印刷キューに送信"""
    if not HAS_WIN32PRINT:
        logging.error("win32printが利用できません")
        return False
    
    try:
        # プリンタが存在するか確認
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        if printer_name not in printers:
            # ネットワークプリンタも確認
            printers_network = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_NETWORK)]
            if printer_name not in printers_network:
                logging.error(f"プリンタが見つかりません: {printer_name}")
                return False
        
        # PDFを印刷
        win32api.ShellExecute(
            0,
            "print",
            str(pdf_path),
            f'/d:"{printer_name}"',
            ".",
            0
        )
        
        logging.info(f"印刷ジョブを投入しました: {pdf_path} -> {printer_name} (部数: {copies})")
        return True
        
    except Exception as e:
        logging.exception(f"印刷エラー: {pdf_path}, {e}")
        return False


def log_print_result(
    campus: str,
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
                "timestamp", "campus", "scan_file", "print_id",
                "printer", "result", "error_message"
            ])
        writer.writerow([
            timestamp, campus, scan_file, print_id or "",
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
        
        # 校舎名を取得
        campus = get_campus_from_folder_path(pdf_path)
        if not campus:
            logging.error(f"校舎名を取得できませんでした: {pdf_path}")
            return
        
        # プリンタ設定を読み込む
        printer_config = load_printer_config()
        if campus not in printer_config:
            logging.error(f"校舎 '{campus}' のプリンタ設定が見つかりません")
            return
        
        printer_name = printer_config[campus]["printer_name"]
        max_copies = printer_config[campus].get("max_copies", 5)
        
        # processingフォルダに移動
        campus_dir = pdf_path.parent.parent  # inフォルダの親（校舎フォルダ）
        processing_dir = campus_dir / "processing"
        processing_dir.mkdir(parents=True, exist_ok=True)
        
        processing_file = processing_dir / pdf_path.name
        
        try:
            shutil.move(str(pdf_path), str(processing_file))
            pdf_path = processing_file  # 以降はprocessing内のファイルを使用
        except Exception as e:
            logging.error(f"processingフォルダへの移動エラー: {e}")
            return
        
        # QRコードからPRINT_IDとプリンター名を抽出
        print_id, qr_printer_name = extract_print_id_from_qr(pdf_path)
        
        # プリンター名の決定: QRコードに含まれていればそれを使用、なければ設定ファイルから取得
        if qr_printer_name:
            printer_name = qr_printer_name
            max_copies = 5  # デフォルト値
            logging.info(f"QRコードからプリンター名を取得: {printer_name}")
        elif campus in printer_config:
            printer_name = printer_config[campus]["printer_name"]
            max_copies = printer_config[campus].get("max_copies", 5)
            logging.info(f"設定ファイルからプリンター名を取得: {printer_name}")
        else:
            logging.error(f"プリンター名を取得できませんでした（校舎: {campus}、QRコードにも含まれていません）")
            return
        
        if not print_id:
            # QRコードが読めない場合、errorフォルダへ
            error_dir = campus_dir / "error"
            error_dir.mkdir(parents=True, exist_ok=True)
            error_file = error_dir / pdf_path.name
            
            try:
                shutil.move(str(pdf_path), str(error_file))
                log_print_result(
                    campus=campus,
                    scan_file=pdf_path.name,
                    print_id=None,
                    printer=printer_name,
                    result="error",
                    error_message="QR not found"
                )
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
            return
        
        # 印刷対象PDFを取得
        print_pdf_path = get_print_pdf_path(print_id)
        
        if not print_pdf_path:
            # 印刷対象PDFが見つからない場合、errorフォルダへ
            error_dir = campus_dir / "error"
            error_dir.mkdir(parents=True, exist_ok=True)
            error_file = error_dir / pdf_path.name
            
            try:
                shutil.move(str(pdf_path), str(error_file))
                log_print_result(
                    campus=campus,
                    scan_file=pdf_path.name,
                    print_id=print_id,
                    printer=printer_name,
                    result="error",
                    error_message=f"PDF not found: {print_id}.pdf"
                )
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
            return
        
        # 印刷実行（部数は1部固定）
        copies = 1
        if copies > max_copies:
            logging.warning(f"印刷部数が最大値を超えています: {copies} > {max_copies}, {max_copies}に制限")
            copies = max_copies
        
        print_success = print_pdf(print_pdf_path, printer_name, copies)
        
        if print_success:
            # 成功時、doneフォルダへ
            done_dir = campus_dir / "done"
            done_dir.mkdir(parents=True, exist_ok=True)
            done_file = done_dir / pdf_path.name
            
            try:
                shutil.move(str(pdf_path), str(done_file))
                log_print_result(
                    campus=campus,
                    scan_file=pdf_path.name,
                    print_id=print_id,
                    printer=printer_name,
                    result="success",
                    error_message=""
                )
                logging.info(f"処理完了: {pdf_path.name} -> {print_id}")
            except Exception as e:
                logging.error(f"doneフォルダへの移動エラー: {e}")
        else:
            # 印刷失敗時、errorフォルダへ
            error_dir = campus_dir / "error"
            error_dir.mkdir(parents=True, exist_ok=True)
            error_file = error_dir / pdf_path.name
            
            try:
                shutil.move(str(pdf_path), str(error_file))
                log_print_result(
                    campus=campus,
                    scan_file=pdf_path.name,
                    print_id=print_id,
                    printer=printer_name,
                    result="error",
                    error_message="Print failed"
                )
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
    
    except Exception as e:
        logging.exception(f"処理エラー: {pdf_path}, {e}")
        # エラー時もerrorフォルダへ移動を試みる
        try:
            campus = get_campus_from_folder_path(pdf_path)
            if campus:
                campus_dir = Path(SCAN_ROOT) / campus
                error_dir = campus_dir / "error"
                error_dir.mkdir(parents=True, exist_ok=True)
                error_file = error_dir / pdf_path.name
                
                if pdf_path.exists():
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


def monitor_campus_folders():
    """校舎別フォルダを監視"""
    if not SCAN_ROOT.exists():
        logging.error(f"スキャンフォルダルートが存在しません: {SCAN_ROOT}")
        return
    
    observer = Observer()
    
    # 各校舎フォルダの in フォルダを監視
    for campus_dir in SCAN_ROOT.iterdir():
        if not campus_dir.is_dir():
            continue
        
        in_dir = campus_dir / "in"
        if in_dir.exists():
            observer.schedule(PDFHandler(), str(in_dir), recursive=False)
            logging.info(f"監視開始: {in_dir}")
        else:
            logging.warning(f"inフォルダが見つかりません: {in_dir}")
    
    observer.start()
    logging.info("フォルダ監視を開始しました")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("停止中...")
        observer.stop()
    
    observer.join()


def main():
    """メイン処理"""
    logging.info("QRスキャン自動印刷システム（方式C）を起動します")
    
    # 設定確認
    if not HAS_PDF2IMAGE:
        logging.error("pdf2imageがインストールされていません: pip install pdf2image")
        return
    
    if not HAS_PYZBAR:
        logging.error("pyzbarがインストールされていません: pip install pyzbar pillow")
        return
    
    if not HAS_WIN32PRINT:
        logging.error("pywin32がインストールされていません: pip install pywin32")
        return
    
    # プリンタ設定確認
    config = load_printer_config()
    if not config:
        logging.warning("プリンタ設定が空です。printers.yamlを確認してください")
    else:
        logging.info(f"プリンタ設定を読み込みました: {len(config)}校舎")
        for campus, settings in config.items():
            logging.info(f"  {campus}: {settings.get('printer_name', 'N/A')}")
    
    # 監視開始
    monitor_campus_folders()


if __name__ == "__main__":
    main()

