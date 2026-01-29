"""
四谷校用QRスキャン自動印刷システム（実験版）
監視フォルダを監視し、スキャンされたPDFを自動印刷
"""

import os
import subprocess
import csv
import shutil
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import unquote

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False
    logging.warning("yaml not available. Install: pip install PyYAML")

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
    import win32ui
    import win32con
    HAS_WIN32PRINT = True
    # 用紙サイズの定数
    DMPAPER_A4 = 9  # A4用紙
    DMORIENT_LANDSCAPE = 2  # 横向き
    DMORIENT_PORTRAIT = 1  # 縦向き
    DM_PAPERSIZE = 0x00000002
    DM_ORIENTATION = 0x00000001
    DM_COPIES = 0x00000100
    # ページレイアウト（複数プリンターで共通）
    DM_DUPLEX = 0x00001000
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
# 環境変数PDF_DIRが設定されていればそれを使用、なければapp.pyと同じデフォルト値を使用
PDF_DIR_ENV = os.environ.get("PDF_DIR")
if PDF_DIR_ENV:
    PDF_DIR = Path(PDF_DIR_ENV)
else:
    # app.pyと同じデフォルト値: Y:\算数科作業フォルダ\10分テスト\test_pdf
    # ただし、test_pdfsフォルダもフォールバックとして検索対象に含める
    PDF_DIR = Path(r"Y:\算数科作業フォルダ\10分テスト\test_pdf")

# 印刷対象PDFの中央リポジトリ（オプション）
PRINT_MATERIALS_ROOT = None  # Path(r"\\server\print_materials")  # 必要に応じて設定

# SumatraPDF（正式インストール先を優先）
SUMATRA_PATH = r"C:\Program Files\SumatraPDF\SumatraPDF.exe"

# ログファイル
LOG_CSV = Path("print_log_yotsuya.csv")

# PRINT_IDとファイル名のマッピングファイル（printviewer側で生成される想定）
PRINT_ID_MAPPING_FILE = Path(r"C:\Users\doctor\printviewer\print_id_mapping.csv")

# プリンタ設定ファイル
PRINTERS_CONFIG = Path(r"C:\Users\doctor\printviewer\printers.yaml")

# プリンタ名（四谷校）
PRINTER_NAME = "執務室"  # FF Apeos C7070 PS H2


def load_printer_config():
    """プリンタ設定を読み込む"""
    if not PRINTERS_CONFIG.exists():
        return {}
    
    if not HAS_YAML:
        logging.warning("yamlモジュールが利用できないため、プリンタ設定ファイルを読み込めません")
        return {}
    
    try:
        with open(PRINTERS_CONFIG, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            return config or {}
    except Exception as e:
        logging.warning(f"プリンタ設定ファイルの読み込みエラー: {e}")
        return {}


def find_printer_by_name(qr_printer_name: str, available_printers: list) -> Optional[str]:
    """QRコードのプリンター名から実際のWindowsプリンター名を検索"""
    try:
        logging.debug(f"find_printer_by_name: QR='{qr_printer_name}', 利用可能プリンタ数={len(available_printers)}")
        
        # 完全一致を試行
        if qr_printer_name in available_printers:
            logging.info(f"完全一致で発見: '{qr_printer_name}'")
            return qr_printer_name
        
        # printers.yamlからマッピングを確認
        try:
            printer_config = load_printer_config()
            logging.debug(f"printers.yamlからマッピング確認中...")
            for campus_key, campus_config in printer_config.items():
                config_printer_name = campus_config.get("printer_name", "")
                if config_printer_name == qr_printer_name:
                    # 設定ファイルのプリンター名と一致する場合、実際のプリンター名を検索
                    # 部分一致で検索（設定ファイルの名前が含まれるプリンターを探す）
                    logging.debug(f"printers.yamlで一致: campus_key='{campus_key}', config_printer_name='{config_printer_name}'")
                    for printer in available_printers:
                        if config_printer_name in printer or printer in config_printer_name:
                            logging.info(f"プリンター名マッピング: '{qr_printer_name}' -> '{printer}'")
                            return printer
        except Exception as e:
            logging.debug(f"プリンター設定ファイル読み込みエラー（無視）: {e}")
        
        # 部分一致で検索（QRコードの名前が含まれるプリンターを探す）
        logging.debug(f"部分一致で検索中...")
        for printer in available_printers:
            if qr_printer_name in printer or printer in qr_printer_name:
                logging.info(f"プリンター名部分一致: '{qr_printer_name}' -> '{printer}'")
                return printer
        
        # 逆方向の部分一致（実際のプリンター名にQRコードの名前が含まれる場合）
        for printer in available_printers:
            if qr_printer_name in printer:
                logging.info(f"プリンター名部分一致（逆）: '{qr_printer_name}' -> '{printer}'")
                return printer
        
        logging.warning(f"プリンター名が見つかりませんでした: '{qr_printer_name}'")
        logging.debug(f"利用可能なプリンタ: {available_printers}")
        return None
    except Exception as e:
        logging.error(f"find_printer_by_name エラー: {e}")
        return None

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
# ルートロガーを設定
root_logger = logging.getLogger()
root_logger.setLevel(logging.INFO)

# 既存のハンドラーをクリア
root_logger.handlers.clear()

# コンソールハンドラー（必ず表示されるように）
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
console_handler.setFormatter(console_formatter)
root_logger.addHandler(console_handler)

# ファイルハンドラー
file_handler = logging.FileHandler("scan_printer_yotsuya.log", encoding="utf-8")
file_handler.setLevel(logging.INFO)
file_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
file_handler.setFormatter(file_formatter)
root_logger.addHandler(file_handler)


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


def guess_filename_from_scan_name(scan_filename: str) -> Optional[str]:
    """スキャンされたファイル名から元のファイル名を推測"""
    # 例: "算数_6年_数の性質_連続する数の積_応用.pdf - QS printer.pdf"
    # → "算数/6年/数の性質_連続する数の積_応用.pdf"
    
    # "- QS printer.pdf" を除去
    base_name = scan_filename.replace(" - QS printer.pdf", "").replace(".pdf", "")
    
    # アンダースコアをスラッシュに変換して推測
    # "算数_6年_数の性質_連続する数の積_応用" → "算数/6年/数の性質_連続する数の積_応用.pdf"
    guessed = base_name.replace("_", "/") + ".pdf"
    
    return guessed


def get_print_pdf_path(original_filename: Optional[str] = None, scan_filename: Optional[str] = None) -> Optional[Path]:
    """ファイル名でPDFファイルのパスを取得"""
    import difflib
    
    # PDF_DIRの存在確認
    logging.info(f"PDF_DIR存在確認: {PDF_DIR} -> {PDF_DIR.exists()}")
    if not PDF_DIR.exists():
        logging.warning(f"PDF_DIRが存在しません: {PDF_DIR}")
    
    # 方法1: QRコードから取得したファイル名で直接検索（最優先）
    if original_filename:
        # ファイル名の前後の空白を除去
        original_filename = original_filename.strip()
        # ファイル名をそのまま使って検索（パス区切りを考慮）
        # Pathオブジェクトで結合（Windowsでも正しく動作）
        pdf_path = PDF_DIR / original_filename
        # デバッグ: 検索パスをログに出力
        logging.info(f"検索パス1: {pdf_path} (存在確認: {pdf_path.exists()})")
        logging.info(f"PDF_DIR: {PDF_DIR} (型: {type(PDF_DIR)}), original_filename: '{original_filename}' (長さ: {len(original_filename)})")
        
        if pdf_path.exists():
            logging.info(f"印刷対象PDFを発見（QRコードのFILE）: {pdf_path}")
            return pdf_path
        else:
            logging.warning(f"QRコードのFILEに記載されているファイルが見つかりません: {original_filename}")
            logging.warning(f"検索したパス: {pdf_path}")
            # 絶対パスでも試行
            abs_path = pdf_path.resolve()
            logging.info(f"絶対パスで再試行: {abs_path} (存在確認: {abs_path.exists()})")
            if abs_path.exists():
                logging.info(f"印刷対象PDFを発見（絶対パス）: {abs_path}")
                return abs_path
            
            # 相対パスでの検索も試行
            if "/" in original_filename or "\\" in original_filename:
                # パス区切りがある場合は、最後のファイル名部分で検索
                filename_part = Path(original_filename).name
                for pdf_file in PDF_DIR.rglob(f"*{filename_part}"):
                    if pdf_file.is_file():
                        logging.info(f"印刷対象PDFを発見（ファイル名部分一致）: {pdf_file}")
                        return pdf_file
                
                # フォルダパスが一致するファイルを検索
                folder_path = "/".join(original_filename.split("/")[:-1]) if "/" in original_filename else ""
                if folder_path:
                    folder_dir = PDF_DIR / folder_path
                    if folder_dir.exists():
                        # 同じフォルダ内でファイル名のキーワードで検索
                        original_name = Path(original_filename).stem  # 拡張子なし
                        # キーワードを抽出（アンダースコアやハイフンで分割）
                        keywords = [k for k in original_name.replace("_", " ").replace("-", " ").split() if len(k) > 1]
                        
                        best_match = None
                        best_score = 0.0
                        
                        for pdf_file in folder_dir.glob("*.pdf"):
                            pdf_name = pdf_file.stem
                            # 類似度を計算
                            score = difflib.SequenceMatcher(None, original_name.lower(), pdf_name.lower()).ratio()
                            
                            # キーワードマッチング（キーワードが含まれている場合はボーナス）
                            keyword_match = sum(1 for kw in keywords if kw in pdf_name)
                            if keyword_match > 0:
                                score += keyword_match * 0.2
                            
                            if score > best_score:
                                best_score = score
                                best_match = pdf_file
                        
                        if best_match and best_score > 0.3:  # 類似度30%以上
                            logging.info(f"印刷対象PDFを発見（類似度マッチング、スコア: {best_score:.2f}）: {best_match}")
                            return best_match
    
    # 方法2: スキャンされたファイル名から推測（FILE=がない場合のフォールバック）
    if scan_filename:
        guessed_filename = guess_filename_from_scan_name(scan_filename)
        if guessed_filename:
            guessed_path = PDF_DIR / guessed_filename
            if guessed_path.exists():
                logging.info(f"印刷対象PDFを発見（ファイル名推測）: {guessed_path}")
                return guessed_path
            else:
                # 部分一致で検索
                base_parts = guessed_filename.split("/")
                if len(base_parts) >= 2:
                    # 最後の2つの部分（例: "6年/数の性質_連続する数の積_応用.pdf"）で検索
                    search_pattern = "/".join(base_parts[-2:])
                    for pdf_file in PDF_DIR.rglob(f"*{search_pattern}"):
                        if pdf_file.is_file():
                            logging.info(f"印刷対象PDFを発見（部分一致検索）: {pdf_file}")
                            return pdf_file
    
    # 方法3: フォールバック用のtest_pdfsフォルダを検索（PDF_DIRが異なる場合）
    fallback_pdf_dir = Path(r"C:\Users\doctor\printviewer\test_pdfs")
    if fallback_pdf_dir.exists() and fallback_pdf_dir != PDF_DIR and original_filename:
        # 同じ検索ロジックを適用
        pdf_path = fallback_pdf_dir / original_filename
        if pdf_path.exists():
            logging.info(f"印刷対象PDFを発見（フォールバックフォルダ）: {pdf_path}")
            return pdf_path
        
        # 類似度マッチングも試行
        if "/" in original_filename or "\\" in original_filename:
            folder_path = "/".join(original_filename.split("/")[:-1]) if "/" in original_filename else ""
            if folder_path:
                folder_dir = fallback_pdf_dir / folder_path
                if folder_dir.exists():
                    import difflib
                    original_name = Path(original_filename).stem
                    keywords = [k for k in original_name.replace("_", " ").replace("-", " ").split() if len(k) > 1]
                    
                    best_match = None
                    best_score = 0.0
                    
                    for pdf_file in folder_dir.glob("*.pdf"):
                        pdf_name = pdf_file.stem
                        score = difflib.SequenceMatcher(None, original_name.lower(), pdf_name.lower()).ratio()
                        keyword_match = sum(1 for kw in keywords if kw in pdf_name)
                        if keyword_match > 0:
                            score += keyword_match * 0.2
                        
                        if score > best_score:
                            best_score = score
                            best_match = pdf_file
                    
                    if best_match and best_score > 0.3:
                        logging.info(f"印刷対象PDFを発見（フォールバックフォルダ、類似度マッチング、スコア: {best_score:.2f}）: {best_match}")
                        return best_match
    
    # 方法4: 中央リポジトリから検索（オプション）
    if PRINT_MATERIALS_ROOT and PRINT_MATERIALS_ROOT.exists() and original_filename:
        pdf_path = PRINT_MATERIALS_ROOT / original_filename
        if pdf_path.exists():
            logging.info(f"印刷対象PDFを発見（中央リポジトリ）: {pdf_path}")
            return pdf_path
    
    # 方法5: すべてのPDFをリストアップしてログに出力
    all_pdfs = []
    if PDF_DIR.exists():
        all_pdfs.extend(list(PDF_DIR.rglob("*.pdf")))
    if fallback_pdf_dir.exists() and fallback_pdf_dir != PDF_DIR:
        all_pdfs.extend(list(fallback_pdf_dir.rglob("*.pdf")))
    
    logging.warning(f"印刷対象PDFが見つかりません: {original_filename or scan_filename}")
    logging.info(f"検索対象フォルダ内のPDFファイル数: {len(all_pdfs)}")
    if len(all_pdfs) <= 20:
        logging.info(f"利用可能なPDFファイル: {[str(p) for p in all_pdfs[:20]]}")
    
    return None


def save_print_id_mapping(print_id: str, filename: str):
    """PRINT_IDとファイル名のマッピングを保存"""
    file_exists = PRINT_ID_MAPPING_FILE.exists()
    
    # 既存のマッピングを読み込む
    existing_mappings = {}
    if file_exists:
        try:
            with open(PRINT_ID_MAPPING_FILE, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    existing_mappings[row["print_id"]] = row["filename"]
        except Exception as e:
            logging.warning(f"マッピングファイル読み込みエラー: {e}")
    
    # 既に存在する場合はスキップ
    if print_id in existing_mappings:
        return
    
    # 新しいマッピングを追加
    existing_mappings[print_id] = filename
    
    # 保存
    try:
        with open(PRINT_ID_MAPPING_FILE, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["print_id", "filename"])
            writer.writeheader()
            for pid, fname in existing_mappings.items():
                writer.writerow({
                    "print_id": pid,
                    "filename": fname
                })
        logging.info(f"マッピングを保存しました: {print_id} -> {filename}")
    except Exception as e:
        logging.warning(f"マッピングファイル保存エラー: {e}")


def extract_print_id_from_qr(pdf_path: Path) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """PDFの1ページ目からQRコードを読み取り、PRINT_ID、FILE、PRINTERを抽出
    戻り値: (print_id, original_filename, printer_name)"""
    if not HAS_PDF2IMAGE or not HAS_PYZBAR:
        logging.error("必要なライブラリがインストールされていません")
        return None, None, None
    
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
            return None, None, None
        
        # QRコードを検出
        for image in images:
            # pyzbarでQRコードを検出（PIL Imageをそのまま使用）
            qr_codes = pyzbar_decode(image)
            
            if not qr_codes:
                continue
            
            # QRコードが複数ある場合はエラー
            if len(qr_codes) > 1:
                logging.warning(f"QRコードが複数検出されました: {pdf_path}")
                return None, None, None
            
            # QRコードデータを取得
            qr_data = qr_codes[0].data.decode('utf-8')
            logging.info(f"QRコード検出（全内容）: {qr_data}")
            
            # PRINT_ID=QS_...,FILE=ファイル名,PRINTER=プリンター名 の形式から情報を抽出
            import re
            print_id = None
            original_filename = None
            printer_name = None
            
            # PRINT_IDを抽出（カンマまたは空白で区切られる）
            match = re.search(r'PRINT_ID=([^,\s]+)', qr_data)
            if match:
                print_id = match.group(1).strip()
                logging.info(f"PRINT_ID抽出: {print_id}")
            else:
                # 旧形式（PRINT_ID=のみ）にも対応
                match = re.match(r'PRINT_ID=(\S+)', qr_data)
                if match:
                    print_id = match.group(1).strip()
                    logging.info(f"PRINT_ID抽出（旧形式）: {print_id}")
            
            # FILEを抽出（カンマで区切られる、または末尾まで）
            # FILE=から次のカンマ、または文字列の終わりまでの部分を取得
            match = re.search(r'FILE=([^,]+)', qr_data)
            if match:
                encoded_filename = match.group(1).strip()
                # URLデコードして元のファイル名に戻す
                original_filename = unquote(encoded_filename)
                logging.info(f"FILE抽出（エンコード前）: {encoded_filename}")
                logging.info(f"FILE抽出（デコード後）: {original_filename}")
            else:
                logging.warning(f"QRコードにFILE=が含まれていません: {qr_data}")
                original_filename = None
            
            # PRINTERを抽出（カンマで区切られる）
            match = re.search(r'PRINTER=([^,\s]+)', qr_data)
            if match:
                printer_name = match.group(1).strip()
                logging.info(f"PRINTER抽出: {printer_name}")
            
            if print_id:
                return print_id, original_filename, printer_name
            else:
                # PRINT_IDがなければエラー
                logging.warning(f"PRINT_IDを抽出できませんでした: {qr_data}")
                return None, None, None
        
        logging.warning(f"QRコードが検出されませんでした: {pdf_path}")
        return None, None, None
        
    except Exception as e:
        logging.exception(f"QRコード読み取りエラー: {pdf_path}, {e}")
        return None, None, None


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
        
        # QRコードのプリンター名を実際のWindowsプリンター名にマッピング
        original_printer_name = printer_name
        logging.info(f"プリンター名解決開始: QRコード='{original_printer_name}', 利用可能なプリンタ数={len(all_printers)}")
        logging.info(f"利用可能なプリンタ一覧: {all_printers}")
        try:
            actual_printer_name = find_printer_by_name(printer_name, all_printers)
            if actual_printer_name:
                logging.info(f"プリンター名を解決: '{original_printer_name}' -> '{actual_printer_name}'")
                printer_name = actual_printer_name
                # 解決後のプリンター名が実際に存在するか再確認
                if printer_name not in all_printers:
                    logging.error(f"解決後のプリンター名が利用可能なプリンタに存在しません: '{printer_name}'")
                    logging.error(f"利用可能なプリンタ: {all_printers}")
                    return False
            else:
                # find_printer_by_nameがNoneを返した場合
                logging.warning(f"find_printer_by_nameがNoneを返しました: '{printer_name}'")
                if printer_name not in all_printers:
                    logging.error(f"プリンタが見つかりません: {printer_name}")
                    logging.error(f"利用可能なプリンタ: {all_printers}")
                    # 部分一致で再試行
                    logging.info(f"部分一致で再検索中...")
                    found = False
                    for printer in all_printers:
                        if printer_name in printer or printer in printer_name:
                            logging.info(f"部分一致で発見: '{printer_name}' -> '{printer}'")
                            printer_name = printer
                            found = True
                            break
                    if not found:
                        logging.error(f"部分一致でも見つかりませんでした: '{printer_name}'")
                        return False
        except Exception as e:
            logging.warning(f"プリンター名解決エラー: {e}。元のプリンター名 '{printer_name}' を使用します")
            # エラーが発生した場合、元のプリンター名で続行を試みる
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
            
            # 印刷に使うアプリを検索（SumatraPDF優先、なければAdobe Reader）
            sumatra_paths = [
                SUMATRA_PATH,
                os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
                r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            ]
            sumatra_path = None
            for path in sumatra_paths:
                if path and Path(path).exists():
                    sumatra_path = path
                    break
            
            acrobat_paths = [
                r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            ]
            acrobat_path = None
            for path in acrobat_paths:
                if Path(path).exists():
                    acrobat_path = path
                    break
            
            # SumatraPDFを優先して印刷（指定プリンターに確実に送りやすい）
            if sumatra_path:
                try:
                    logging.info(f"SumatraPDFを使用して印刷: {sumatra_path}")
                    
                    # 印刷前にNup設定を強制的に1に設定
                    original_nup_values = {}
                    try:
                        printer_handle = win32print.OpenPrinter(printer_name)
                        try:
                            printer_info = win32print.GetPrinter(printer_handle, 2)
                            if printer_info and 'pDevMode' in printer_info and printer_info['pDevMode']:
                                devmode = printer_info['pDevMode']
                                logging.info(f"プリンター設定（現在）: PaperSize={devmode.PaperSize}, Orientation={devmode.Orientation}, Copies={devmode.Copies}")
                                
                                # リコーApeos C7070固有のNup設定を確認し、強制的に1に設定
                                devmode_attrs = dir(devmode)
                                nup_attrs = ['PagesPerSheet', 'PagesPerSheetN', 'Nup', 'NupOrientOrder', 'PagesPerSheetNup']
                                
                                needs_update = False
                                for attr_name in nup_attrs:
                                    if attr_name in devmode_attrs:
                                        try:
                                            current_value = getattr(devmode, attr_name)
                                            logging.info(f"リコーApeos C7070: {attr_name} = {current_value} (現在の設定)")
                                            
                                            # 1以外の場合は1に強制設定が必要
                                            if current_value != 1:
                                                # 元の値を保存
                                                original_nup_values[attr_name] = current_value
                                                setattr(devmode, attr_name, 1)
                                                logging.info(f"リコーApeos C7070: {attr_name} を {current_value} -> 1 に変更します")
                                                needs_update = True
                                        except Exception as e:
                                            logging.debug(f"{attr_name} の設定をスキップ: {e}")
                                
                                # 設定変更が必要な場合のみSetPrinterで保存を試行
                                # ただし、権限不足の場合はスキップして印刷を継続
                                if needs_update:
                                    try:
                                        # printer_infoのpDevModeを更新
                                        printer_info['pDevMode'] = devmode
                                        # SetPrinterで設定を保存
                                        win32print.SetPrinter(printer_handle, 2, printer_info, 0)
                                        logging.info(f"Nup設定を1に強制設定しました: {printer_name}")
                                    except Exception as e:
                                        error_code = e.args[0] if e.args else None
                                        if error_code == 5:  # アクセス拒否
                                            logging.info(f"SetPrinter権限不足（エラー5）: プリンター設定の変更をスキップして印刷を継続します")
                                            logging.info(f"プリンター '{printer_name}' のNup設定は現在のまま（{original_nup_values}）で印刷されます")
                                            # SetPrinterが失敗した場合、devmodeの変更を元に戻す（メモリ上の変更をクリア）
                                            for attr_name, original_value in original_nup_values.items():
                                                try:
                                                    setattr(devmode, attr_name, original_value)
                                                    logging.debug(f"{attr_name} を元の値 {original_value} に戻しました")
                                                except Exception:
                                                    pass
                                        else:
                                            logging.warning(f"SetPrinterで設定を保存できませんでした: {e}")
                                            logging.warning(f"プリンター '{printer_name}' の設定を手動で1ページ/枚に設定してください")
                                            # SetPrinterが失敗した場合、devmodeの変更を元に戻す
                                            for attr_name, original_value in original_nup_values.items():
                                                try:
                                                    setattr(devmode, attr_name, original_value)
                                                    logging.debug(f"{attr_name} を元の値 {original_value} に戻しました")
                                                except Exception:
                                                    pass
                                else:
                                    logging.info(f"プリンター '{printer_name}' のNup設定は既に1です（変更不要）")
                        finally:
                            win32print.ClosePrinter(printer_handle)
                    except Exception as e:
                        logging.warning(f"プリンター設定の変更をスキップしました: {e}")
                        logging.warning(f"プリンター '{printer_name}' の設定を手動で1ページ/枚に設定してください")
                    
                    # Nup設定を1に強制設定した状態で印刷
                    # SumatraPDFで印刷（-print-to で指定プリンターに直接送信）
                    logging.info(f"SumatraPDF印刷コマンド実行: ファイル='{abs_path}', プリンター='{printer_name}'")
                    logging.info(f"利用可能なプリンタリスト: {all_printers}")
                    if printer_name not in all_printers:
                        logging.error(f"警告: プリンター名 '{printer_name}' が利用可能なプリンタリストに存在しません")
                        logging.error(f"利用可能なプリンタ: {all_printers}")
                        return False
                    
                    # subprocessで起動（ShellExecuteはAppData配下でアクセス拒否になることがあるため）
                    # -print-settings "simplex,monochrome,fit": 片面・グレースケール・用紙に収める（校舎差・ドライバ既定を無視）
                    sumatra_ok = False
                    try:
                        creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
                        proc = subprocess.run(
                            [sumatra_path, "-print-to", printer_name, "-print-settings", "simplex,monochrome,fit", abs_path, "-silent"],
                            cwd=str(Path(abs_path).parent),
                            capture_output=True,
                            timeout=60,
                            creationflags=creationflags,
                        )
                        logging.info(f"SumatraPDF 終了コード: {proc.returncode}")
                        if proc.returncode == 0:
                            logging.info(f"印刷ジョブを投入しました（SumatraPDF経由）: {pdf_path.name} -> {printer_name}")
                            sumatra_ok = True
                        else:
                            if proc.stderr:
                                logging.warning(f"SumatraPDF stderr: {proc.stderr.decode(errors='replace')}")
                            raise Exception(f"SumatraPDF returned {proc.returncode}")
                    except subprocess.TimeoutExpired:
                        logging.warning("SumatraPDFがタイムアウトしました（印刷は送信済みの可能性あり）")
                        sumatra_ok = True
                    except OSError as e:
                        # WinError 193 = 32bit Python から 64bit exe を起動した場合など
                        if getattr(e, "winerror", None) == 193:
                            logging.warning("WinError 193: 32bit Python と 64bit SumatraPDF の組み合わせの可能性があります。ShellExecuteで再試行します。")
                            result = win32api.ShellExecute(
                                0, "open", sumatra_path,
                                f'-print-to "{printer_name}" -print-settings "simplex,monochrome,fit" "{abs_path}" -silent',
                                str(Path(abs_path).parent), 0
                            )
                            if result > 32:
                                logging.info(f"印刷ジョブを投入しました（SumatraPDF経由・ShellExecute）: {pdf_path.name} -> {printer_name}")
                                sumatra_ok = True
                            else:
                                logging.warning(f"ShellExecuteも失敗: 戻り値={result}")
                                raise Exception(f"SumatraPDF ShellExecute returned {result}")
                        else:
                            raise
                    if sumatra_ok:
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
                except Exception as sumatra_error:
                    logging.warning(f"SumatraPDFでの印刷に失敗: {sumatra_error}")
                    logging.info("デフォルトプリンタを一時的に変更する方法にフォールバックします...")
                    
                    # 方法3: デフォルトプリンタを一時的に変更して印刷（最後の手段）
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
            elif acrobat_path:
                # SumatraPDFが無い場合のフォールバック: Adobe Readerで印刷
                logging.info("SumatraPDFが見つからないためAdobe Readerを使用します")
                try:
                    logging.info(f"Adobe Readerを使用して印刷: {acrobat_path}")
                    if printer_name not in all_printers:
                        logging.error(f"警告: プリンター名 '{printer_name}' が利用可能なプリンタリストに存在しません")
                        return False
                    result = win32api.ShellExecute(
                        0, "open", acrobat_path,
                        f'/t "{abs_path}" "{printer_name}"',
                        str(Path(abs_path).parent), 0
                    )
                    if result > 32:
                        logging.info(f"印刷ジョブを投入しました（Adobe Reader経由）: {pdf_path.name} -> {printer_name}")
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
                        raise Exception(f"Adobe Reader ShellExecute returned {result}")
                except Exception as acrobat_error:
                    logging.warning(f"Adobe Readerでの印刷に失敗: {acrobat_error}")
                    logging.info("デフォルトプリンタを一時的に変更する方法にフォールバックします...")
                    try:
                        default_printer = win32print.GetDefaultPrinter()
                        win32print.SetDefaultPrinter(printer_name)
                        result = win32api.ShellExecute(0, "print", abs_path, None, str(Path(abs_path).parent), 0)
                        win32print.SetDefaultPrinter(default_printer)
                        if result > 32:
                            logging.info(f"印刷ジョブを投入しました（デフォルト変更）: {pdf_path.name} -> {printer_name}")
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
                        return False
                    except Exception as e2:
                        logging.exception(f"代替方法での印刷も失敗: {e2}")
                        try:
                            if 'default_printer' in locals():
                                win32print.SetDefaultPrinter(default_printer)
                        except Exception:
                            pass
                        return False
            else:
                logging.error("SumatraPDFもAdobe Readerも見つかりません。SumatraPDFのインストールを推奨します。")
                return False
                    
        finally:
            # プリンター設定を元に戻す（オプション：必要に応じてコメントアウト）
            # 元の設定を残したい場合は以下のコードを有効化
            # if original_orientation is not None and printer_name:
            #     try:
            #         restore_handle = win32print.OpenPrinter(printer_name)
            #         if restore_handle:
            #             printer_info = win32print.GetPrinter(restore_handle, 2)
            #             if printer_info and 'pDevMode' in printer_info and printer_info['pDevMode']:
            #                 devmode = printer_info['pDevMode']
            #                 devmode.Orientation = original_orientation
            #                 if original_paper_size is not None:
            #                     devmode.PaperSize = original_paper_size
            #                 if original_copies is not None:
            #                     devmode.Copies = original_copies
            #                 devmode.Fields = devmode.Fields | DM_ORIENTATION | DM_PAPERSIZE | DM_COPIES
            #                 win32print.SetPrinter(restore_handle, 2, printer_info, 0)
            #                 logging.info(f"プリンター設定を元に戻しました")
            #             win32print.ClosePrinter(restore_handle)
            #     except Exception as restore_error:
            #         logging.warning(f"プリンター設定の復元に失敗: {restore_error}")
            pass
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
    logging.info(f"handle_pdf呼び出し: {pdf_path} (存在: {pdf_path.exists()})")
    # 既に処理中のファイルはスキップ（二重処理防止）
    file_key = str(pdf_path)
    if file_key in processing_files:
        logging.info(f"処理中のファイルをスキップ: {pdf_path}")
        return
    
    processing_files.add(file_key)
    
    try:
        # ファイルがPDFか確認
        if pdf_path.suffix.lower() != ".pdf":
            logging.warning(f"PDF以外のファイルをスキップ: {pdf_path}")
            return
        
        logging.info(f"PDF検出: {pdf_path}")
        
        # ファイル安定化待ち
        if not wait_until_file_stable(pdf_path):
            logging.warning(f"ファイルが安定しませんでした: {pdf_path}")
            return
        
        # QRコードからファイル名（FILE=）とプリンター名（PRINTER=）を抽出（PRINT_IDはログ用にのみ使用）
        print_id, original_filename, qr_printer_name = extract_print_id_from_qr(pdf_path)
        
        # プリンター名の決定: QRコードに含まれていればそれを使用、なければデフォルト値
        if qr_printer_name:
            printer_name = qr_printer_name
            logging.info(f"QRコードからプリンター名を取得: {printer_name}")
        else:
            printer_name = PRINTER_NAME  # デフォルトのプリンター名
            logging.info(f"デフォルトのプリンター名を使用: {printer_name}")
        
        if not original_filename:
            # FILE=が含まれていない場合、errorフォルダへ
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
                    print_id=print_id or "unknown",
                    printer=printer_name,
                    result="error",
                    error_message="FILE not found in QR"
                )
                logging.warning(f"QRコードにFILE=が含まれていません: {pdf_path}")
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
            return
        
        # ファイル名で印刷対象PDFを取得
        scan_filename = pdf_path.name
        print_pdf_path = get_print_pdf_path(original_filename, scan_filename)
        
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
                    print_id=print_id or "unknown",
                    printer=printer_name,
                    result="error",
                    error_message=f"PDF not found: {original_filename}"
                )
                logging.error(f"印刷対象PDFが見つかりません: {original_filename}")
            except Exception as e:
                logging.error(f"errorフォルダへの移動エラー: {e}")
            return
        
        # 印刷対象PDFを印刷（部数は1部固定）
        print_success = print_pdf(print_pdf_path, printer_name, 1)
        
        if print_success:
            # 成功時、processedフォルダへ移動
            # 移動元が既に無い場合は「他スレッドで移動済み」として処理完了扱い、異常終了にしない
            if not pdf_path.exists():
                logging.info(f"移動元ファイルが既に存在しません（他スレッドで移動済みの可能性）: {pdf_path.name}。処理完了扱いとします。")
                log_print_result(
                    scan_file=pdf_path.name,
                    print_id=print_id or "unknown",
                    printer=printer_name,
                    result="success",
                    error_message=""
                )
                logging.info(f"処理完了: {pdf_path.name} -> {original_filename} (PRINT_ID: {print_id or 'N/A'})")
            else:
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
                        print_id=print_id or "unknown",
                        printer=printer_name,
                        result="success",
                        error_message=""
                    )
                    logging.info(f"処理完了: {pdf_path.name} -> {original_filename} (PRINT_ID: {print_id or 'N/A'})")
                except FileNotFoundError:
                    logging.info(f"移動元が移動中に削除されました: {pdf_path.name}。処理完了扱いとします。")
                    log_print_result(
                        scan_file=pdf_path.name,
                        print_id=print_id or "unknown",
                        printer=printer_name,
                        result="success",
                        error_message=""
                    )
                    logging.info(f"処理完了: {pdf_path.name} -> {original_filename} (PRINT_ID: {print_id or 'N/A'})")
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
                    print_id=print_id or "unknown",
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
        logging.info(f"ファイル検出イベント: {path} (拡張子: {path.suffix})")
        if path.suffix.lower() == ".pdf":
            logging.info(f"PDFファイルを検出しました: {path}")
            # 処理は別スレッドで実行（非ブロッキング）
            import threading
            threading.Thread(target=self._handle_pdf_delayed, args=(path,), daemon=True).start()
        else:
            logging.debug(f"PDF以外のファイルをスキップ: {path}")
    
    def _handle_pdf_delayed(self, path: Path):
        """少し待ってから処理（ファイル作成完了を待つ）"""
        logging.info(f"PDF処理を開始: {path}")
        time.sleep(0.3)
        handle_pdf(path)
    
    def on_modified(self, event):
        """ファイル変更イベント（ネットワークドライブでon_createdが発火しない場合の対策）"""
        if event.is_directory:
            return
        path = Path(event.src_path)
        logging.info(f"ファイル変更イベント: {path} (拡張子: {path.suffix})")
        if path.suffix.lower() == ".pdf":
            # ファイルサイズが0でない場合のみ処理（作成中はスキップ）
            try:
                if path.exists() and path.stat().st_size > 0:
                    logging.info(f"PDFファイル変更を検出しました: {path}")
                    import threading
                    threading.Thread(target=self._handle_pdf_delayed, args=(path,), daemon=True).start()
            except Exception as e:
                logging.debug(f"ファイル変更イベント処理エラー: {e}")
    
    def on_moved(self, event):
        if event.is_directory:
            return
        path = Path(event.dest_path)
        logging.info(f"ファイル移動イベント: {path}")
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
        
        # 定期的なポーリング（ネットワークドライブ対策）
        last_poll_time = time.time()
        poll_interval = 2.0  # 2秒ごとにポーリング
        
        def poll_for_new_pdfs():
            """フォルダ内のPDFファイルを定期的にチェック"""
            try:
                if not SCAN_DIR.exists():
                    return
                for pdf_file in SCAN_DIR.glob("*.pdf"):
                    # 処理済みフォルダやエラーフォルダのファイルはスキップ
                    if pdf_file.parent == PROCESSED_DIR or pdf_file.parent == ERROR_DIR:
                        continue
                    # 処理中のファイルはスキップ
                    file_key = str(pdf_file)
                    if file_key in processing_files:
                        continue
                    # ファイルサイズが0でない場合のみ処理
                    try:
                        if pdf_file.exists() and pdf_file.stat().st_size > 0:
                            # ファイルの最終更新時刻をチェック（最近変更されたファイルのみ）
                            mtime = pdf_file.stat().st_mtime
                            if time.time() - mtime < 10:  # 10秒以内に変更されたファイル
                                logging.info(f"ポーリングでPDFを検出: {pdf_file}")
                                import threading
                                threading.Thread(target=handle_pdf, args=(pdf_file,), daemon=True).start()
                    except Exception as e:
                        logging.debug(f"ポーリング処理エラー ({pdf_file}): {e}")
            except Exception as e:
                logging.debug(f"ポーリングエラー: {e}")
        
        try:
            while True:
                current_time = time.time()
                # 定期的にポーリング
                if current_time - last_poll_time >= poll_interval:
                    poll_for_new_pdfs()
                    last_poll_time = current_time
                time.sleep(0.5)  # 0.5秒ごとにチェック
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

