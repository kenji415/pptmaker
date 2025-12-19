import os
import re
import time
import shutil
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

import fitz  # pymupdf
import cv2
import numpy as np

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


# ==========================
# 設定
# ==========================
# ショートカットファイルのパス（実パスに解決される）
INBOX_SHORTCUT = Path(r"C:\Users\doctor\Desktop\99 共有スキャン画像.lnk")
INBOX_DIR = None  # 後で解決される

# 講師別出力先（ここ直下に「山口」「田中」などの講師フォルダがある想定）
OUT_ROOT  = Path(r"Y:\QSスキャン")

# エラー退避先・ログ
ERROR_DIR = Path(r"C:\Scan\ERROR")          # QR読めない/形式不正はここへ退避
LOG_PATH  = Path(r"C:\Scan\scan_router.log")

STUDENT_SUFFIX = "さん"  # 生徒名の敬称
FNAME_TEMPLATE = "QS_{date}_{student}{suffix}_{teacher}.pdf"

DUPLICATE_SUFFIX_TEMPLATE = "_{n}"  # QS_xxx_yyy_2.pdf のように付番
TARGET_EXT = ".pdf"

# 「コピー中の未完成PDF」を避けるための待機
STABLE_CHECK_INTERVAL_SEC = 0.5
STABLE_CHECK_COUNT = 6  # 0.5秒×6回=3秒間サイズ変化なしで安定扱い


# ==========================
# ログ設定
# ==========================
# 先にログフォルダを作成しておく（例: C:\Scan\）
LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler()
    ],
)


# ==========================
# ユーティリティ
# ==========================
def resolve_shortcut(shortcut_path: Path) -> Optional[Path]:
    """Windowsショートカット（.lnk）の実パスを解決"""
    if not shortcut_path.exists():
        return None
    
    if shortcut_path.suffix.lower() != ".lnk":
        # ショートカットでない場合はそのまま返す
        return shortcut_path if shortcut_path.is_dir() else None
    
    if not HAS_WIN32COM:
        logging.warning("win32com not available. Install pywin32: pip install pywin32")
        return None
    
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(shortcut_path))
        target_path = Path(shortcut.TargetPath)
        if target_path.is_dir():
            return target_path
        else:
            logging.warning(f"Shortcut target is not a directory: {target_path}")
            return None
    except Exception as e:
        logging.error(f"Failed to resolve shortcut {shortcut_path}: {e}")
        return None


def sanitize_filename_part(s: str) -> str:
    """WindowsでNGな文字を除去・整形"""
    s = s.strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s)
    return s


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


def render_first_page_to_image(pdf_path: Path, zoom: float = 3.0) -> np.ndarray:
    """PDFの1ページ目を画像化（OpenCV BGR）"""
    doc = fitz.open(pdf_path)
    if doc.page_count < 1:
        doc.close()
        raise ValueError("PDF has no pages")

    page = doc.load_page(0)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    doc.close()

    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    img_bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return img_bgr


def decode_qr_from_image(img_bgr: np.ndarray) -> Optional[str]:
    """
    OpenCVでQRコードを検出＆デコード

    - 用紙左下にQRがある前提で、その周辺を優先的にトライ
    - コントラストが低い場合に備えて、二値化・拡大など複数パターンを試す
    """

    def _try_decode(detector: cv2.QRCodeDetector, img: np.ndarray, label: str) -> Optional[str]:
        try:
            data, _, _ = detector.detectAndDecode(img)
            data = (data or "").strip()
            if data:
                logging.info(f"QR decoded ({label}): '{data}'")
            return data if data else None
        except Exception as e:
            logging.debug(f"QR decode failed ({label}): {e}")
            return None

    detector = cv2.QRCodeDetector()
    h, w = img_bgr.shape[:2]

    # 1) まず左下の広めのROIを試す（高さ下側40%、左側40%）
    roi = img_bgr[int(h * 0.6): h, 0: int(w * 0.4)]
    if roi.size > 0:
        # 1-1) ROIそのまま
        data = _try_decode(detector, roi, "roi_raw")
        if data:
            return data

        # 1-2) ROIを拡大
        roi_big = cv2.resize(roi, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
        data = _try_decode(detector, roi_big, "roi_big")
        if data:
            return data

        # 1-3) ROI → グレースケール＋二値化（背景がザラザラなとき用）
        gray = cv2.cvtColor(roi_big, cv2.COLOR_BGR2GRAY)
        # Otsu閾値で二値化
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        data = _try_decode(detector, th, "roi_big_otsu")
        if data:
            return data

    # 2) ページ全体でトライ
    data = _try_decode(detector, img_bgr, "full_bgr")
    if data:
        return data

    # 3) ページ全体をグレースケール＋二値化
    gray_full = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    _, th_full = cv2.threshold(gray_full, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    data = _try_decode(detector, th_full, "full_otsu")

    return data if data else None


def parse_qr_payload(payload: str) -> Optional[Tuple[str, str, Optional[str]]]:
    """
    QR形式：生徒名,講師名[,テキスト名]
    例：山中秀悟,田中
    例：山中秀悟,田中,数の性質_応用
    戻り値：(teacher, student, text_name)
    """
    p = payload.strip()
    if "," not in p:
        return None

    parts = [x.strip() for x in p.split(",")]
    
    # 2つの場合（旧形式）：生徒名,講師名
    if len(parts) == 2:
        student, teacher = parts
        if not student or not teacher:
            return None
        return teacher, student, None
    
    # 3つ以上の場合（新形式）：生徒名,講師名,テキスト名
    if len(parts) >= 3:
        student, teacher, text_name = parts[0], parts[1], parts[2]
        if not student or not teacher:
            return None
        return teacher, student, text_name
    
    return None


def build_destination(teacher: str, student: str, text_name: Optional[str] = None) -> Path:
    """移動先（講師フォルダ + 新ファイル名）を決定"""
    teacher_s = sanitize_filename_part(teacher)
    student_s = sanitize_filename_part(student)

    teacher_dir = OUT_ROOT / teacher_s
    teacher_dir.mkdir(parents=True, exist_ok=True)

    # スキャン日付を取得（YYYY.MM.DD形式）
    scan_date = datetime.now().strftime("%Y.%m.%d")

    # テキスト名がある場合はファイル名に含める
    if text_name:
        text_s = sanitize_filename_part(text_name)
        # テキスト名をファイル名に追加（例: QS_2024.01.01_山中秀悟さん_田中_数の性質_応用.pdf）
        base_name = FNAME_TEMPLATE.format(date=scan_date, student=student_s, suffix=STUDENT_SUFFIX, teacher=teacher_s)
        # テキスト名を追加（拡張子の前に挿入）
        base_name = base_name.replace(".pdf", f"_{text_s}.pdf")
    else:
        base_name = FNAME_TEMPLATE.format(date=scan_date, student=student_s, suffix=STUDENT_SUFFIX, teacher=teacher_s)
    
    base_name = sanitize_filename_part(base_name)

    dst = teacher_dir / base_name

    # 同名があれば連番
    if dst.exists():
        stem = dst.stem
        suffix = dst.suffix
        n = 2
        while True:
            candidate = teacher_dir / f"{stem}{DUPLICATE_SUFFIX_TEMPLATE.format(n=n)}{suffix}"
            if not candidate.exists():
                dst = candidate
                break
            n += 1

    return dst


def move_to_error(pdf_path: Path, tag: str) -> None:
    """
    QRなし / 読み取れない / 形式不正 などの場合の処理。
    いまは「INBOXからは動かさず、ログだけ残す」運用にする。
    """
    logging.error(f"{tag}: QR error on file (NOT moved): {pdf_path}")


def handle_pdf(pdf_path: Path) -> None:
    """PDF1つの処理：安定化待ち→QR読取→リネーム＆移動（=move）"""
    try:
        if pdf_path.suffix.lower() != TARGET_EXT:
            return

        logging.info(f"Detected: {pdf_path}")

        if not wait_until_file_stable(pdf_path):
            logging.warning(f"File not stable or disappeared: {pdf_path}")
            return

        img = render_first_page_to_image(pdf_path, zoom=2.0)
        qr = decode_qr_from_image(img)

        if not qr:
            move_to_error(pdf_path, "NOQR")
            return

        parsed = parse_qr_payload(qr)
        if not parsed:
            move_to_error(pdf_path, "BADQR")
            logging.error(f"Bad QR payload: '{qr}'")
            return

        teacher, student, text_name = parsed
        dst = build_destination(teacher, student, text_name)

        # move = 移動しつつファイル名変更（リネーム+移動を一度に）
        shutil.move(str(pdf_path), str(dst))
        if text_name:
            logging.info(f"OK teacher='{teacher}' student='{student}' text='{text_name}' => {dst}")
        else:
            logging.info(f"OK teacher='{teacher}' student='{student}' => {dst}")

    except Exception as e:
        logging.exception(f"Failed handling {pdf_path}: {e}")
        # 例外でもPDFが残ってたらエラーへ退避（事故防止）
        try:
            if pdf_path.exists():
                move_to_error(pdf_path, "EXCEPTION")
        except Exception:
            pass


# ==========================
# Watchdog ハンドラ
# ==========================
class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() == TARGET_EXT:
            time.sleep(0.3)  # 作成直後の書込中対策
            handle_pdf(path)

    def on_moved(self, event):
        if event.is_directory:
            return
        path = Path(event.dest_path)
        if path.suffix.lower() == TARGET_EXT:
            time.sleep(0.3)
            handle_pdf(path)


def main():
    global INBOX_DIR
    
    # ショートカットから実パスを解決
    resolved_inbox = resolve_shortcut(INBOX_SHORTCUT)
    if resolved_inbox is None:
        logging.error(f"Failed to resolve INBOX shortcut: {INBOX_SHORTCUT}")
        logging.error("Please check if the shortcut exists and points to a valid directory.")
        return
    
    INBOX_DIR = resolved_inbox
    logging.info(f"Resolved INBOX from shortcut: {INBOX_SHORTCUT} -> {INBOX_DIR}")
    
    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    OUT_ROOT.mkdir(parents=True, exist_ok=True)
    ERROR_DIR.mkdir(parents=True, exist_ok=True)

    logging.info("Starting scan router (rename+move, up to step ②)...")
    logging.info(f"INBOX: {INBOX_DIR}")
    logging.info(f"OUT:   {OUT_ROOT}")
    logging.info(f"ERROR: {ERROR_DIR}")

    observer = Observer()
    observer.schedule(PDFHandler(), str(INBOX_DIR), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("Stopping...")
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()

