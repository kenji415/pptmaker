import os
import csv
import re
import locale
import hashlib
import uuid
import yaml
from datetime import datetime
from urllib.parse import quote, unquote
from functools import wraps
from pathlib import Path

from flask import Flask, render_template, send_file, request, abort, session, redirect, url_for, flash, jsonify
from pdf2image import convert_from_path
from werkzeug.security import generate_password_hash, check_password_hash
from PIL import ImageDraw, ImageFont, Image
import qrcode
import fitz  # PyMuPDF

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "your-secret-key-change-this-in-production")

# 設定: 環境変数で上書き可能、なければローカルパスを使用
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PDF_DIR = os.environ.get("PDF_DIR", r"Y:\算数科作業フォルダ\10分テスト\test_pdf")
CACHE_DIR = os.path.join(BASE_DIR, "cache")
LOG_FILE = os.path.join(BASE_DIR, "logs", "print_log.csv")
POPPLER_PATH = os.environ.get("POPPLER_PATH", None)
if POPPLER_PATH is None:
    default_poppler_path = r"C:\tools\poppler-25.12.0\Library\bin"
    if os.path.exists(default_poppler_path):
        POPPLER_PATH = default_poppler_path
USERS_FILE = os.path.join(BASE_DIR, "users.csv")
STUDENTS_DIR = os.path.join(BASE_DIR, "students")
TEXT_MAPPING_FILE = os.path.join(BASE_DIR, "text_mapping.csv")
FILE_NAME_HISTORY_FILE = os.path.join(BASE_DIR, "file_name_history.csv")
PRINT_ID_MAPPING_FILE = os.path.join(BASE_DIR, "print_id_mapping.csv")
PRINTERS_CONFIG = os.path.join(BASE_DIR, "printers.yaml")

# 必要なディレクトリを作成
os.makedirs(STUDENTS_DIR, exist_ok=True)
os.makedirs(CACHE_DIR, exist_ok=True)
os.makedirs(os.path.dirname(LOG_FILE) if os.path.dirname(LOG_FILE) else ".", exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)


def get_current_user():
    """セッションからユーザー名を取得"""
    return session.get("username", "unknown")


def load_users():
    """ユーザー情報を読み込む"""
    users = {}
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # ヘッダー行が存在し、必要なキーがあることを確認
                    if "username" in row and "password_hash" in row:
                        users[row["username"]] = row["password_hash"]
                    elif len(row) >= 2:
                        # ヘッダー行がない場合のフォールバック（最初の列がusername、2番目がpassword_hash）
                        keys = list(row.keys())
                        if len(keys) >= 2:
                            users[row[keys[0]]] = row[keys[1]]
        except Exception as e:
            import traceback
            print(f"ERROR: ユーザーファイル読み込みエラー: {e}")
            print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
    return users


def save_user(username, password_hash):
    """ユーザー情報を保存"""
    users = load_users()
    users[username] = password_hash
    
    file_exists = os.path.exists(USERS_FILE)
    with open(USERS_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["username", "password_hash"])
        for user, pwd_hash in users.items():
            writer.writerow([user, pwd_hash])


def login_required(f):
    """ログイン必須デコレータ"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


def load_printer_config():
    """プリンタ設定を読み込む"""
    if not os.path.exists(PRINTERS_CONFIG):
        return {}
    
    try:
        with open(PRINTERS_CONFIG, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            return config or {}
    except Exception as e:
        print(f"プリンタ設定ファイルの読み込みエラー: {e}")
        return {}


def get_printer_name_by_campus(campus_name: str) -> str:
    """校舎名からプリンター名を取得"""
    config = load_printer_config()
    if campus_name and campus_name in config:
        return config[campus_name].get("printer_name", "")
    return ""


def generate_print_id():
    """一意なPRINT_IDを生成（形式: QS_YYYY_NNNNN）"""
    year = datetime.now().strftime("%Y")
    # 短い一意ID（UUIDの最初の5文字 + タイムスタンプの一部）
    unique_part = hashlib.md5(uuid.uuid4().bytes + datetime.now().isoformat().encode()).hexdigest()[:5].upper()
    return f"QS_{year}_{unique_part}"


def create_header_with_qr(filename, username, text_name, campus_name=None):
    """頭紙PDFにQRコードを重ねて画像を生成"""
    # 頭紙PDFのパス
    header_template_path = os.path.join(BASE_DIR, "templates", "頭紙.pdf")
    if not os.path.exists(header_template_path):
        raise FileNotFoundError("頭紙テンプレートが見つかりません")
    
    # 頭紙PDFを画像に変換（1ページのみ）
    header_images = convert_from_path(header_template_path, poppler_path=POPPLER_PATH, first_page=1, last_page=1)
    if not header_images:
        raise ValueError("頭紙PDFの変換に失敗しました")
    
    img = header_images[0]
    draw = ImageDraw.Draw(img)
    img_width, img_height = img.size
    
    # PRINT_IDを生成
    print_id = generate_print_id()
    
    # 元のファイル名を取得
    original_filename = filename.replace('\\', '/')
    
    # PRINT_IDとファイル名のマッピングを保存
    save_print_id_mapping(print_id, original_filename)
    
    # QRコードのデータを生成
    encoded_filename = quote(original_filename, safe='/')
    qr_data = f"PRINT_ID={print_id},FILE={encoded_filename}"
    
    # 校舎が選択されている場合、プリンター名をQRコードに追加
    if campus_name:
        printer_name = get_printer_name_by_campus(campus_name)
        if printer_name:
            qr_data += f",PRINTER={printer_name}"
    
    # QRコードを生成
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=15,
        border=4,
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    
    # QRコード画像を生成
    qr_img = qr.make_image(fill_color="black", back_color="white")
    
    # QRコードのサイズを調整（画像サイズの約15%に増加）
    qr_size = int(min(img_width, img_height) * 0.15)
    qr_img = qr_img.resize((qr_size, qr_size), Image.Resampling.LANCZOS)
    
    # フォントパス
    font_paths = [
        "C:/Windows/Fonts/msgothic.ttc",
        "C:/Windows/Fonts/meiryo.ttc",
        "C:/Windows/Fonts/msmincho.ttc",
        "arial.ttf"
    ]
    
    # ファイル名のタイトルを中央に表示
    # ファイル名から拡張子を除く
    file_title = os.path.splitext(os.path.basename(original_filename))[0]
    # パス区切り文字をスペースに変換（例: "算数/6年/数の性質" → "算数 6年 数の性質"）
    file_title = file_title.replace('/', ' ').replace('\\', ' ')
    
    # タイトル用フォントを準備（大きめのサイズ）
    title_font_size = max(32, int(img_width / 40))
    title_font = None
    for font_path in font_paths:
        try:
            title_font = ImageFont.truetype(font_path, title_font_size)
            break
        except Exception:
            continue
    if title_font is None:
        title_font = ImageFont.load_default()
    
    # タイトルのサイズを取得
    title_bbox = draw.textbbox((0, 0), file_title, font=title_font)
    title_width = title_bbox[2] - title_bbox[0]
    title_height = title_bbox[3] - title_bbox[1]
    
    # 画面中央にタイトルを配置
    title_x = (img_width - title_width) / 2
    title_y = (img_height - title_height) / 2
    
    # タイトルを描画
    draw.text(
        (int(title_x), int(title_y)),
        file_title,
        fill=(0, 0, 0, 255),
        font=title_font
    )
    
    # テキスト用フォントを準備（PRINT_ID用）
    text_font_size = max(14, int(img_width / 80))
    text_font = None
    for font_path in font_paths:
        try:
            text_font = ImageFont.truetype(font_path, text_font_size)
            break
        except Exception:
            continue
    if text_font is None:
        text_font = ImageFont.load_default()
    
    # PRINT_IDのテキストのサイズを取得
    text_id = print_id
    text_bbox = draw.textbbox((0, 0), text_id, font=text_font)
    text_text_width = text_bbox[2] - text_bbox[0]
    text_text_height = text_bbox[3] - text_bbox[1]
    
    # テキストの高さを考慮してQRコードの位置を決定（以前と同じ位置：左下）
    bottom_margin = 15
    text_margin = 10
    total_height = qr_size + text_margin + text_text_height + bottom_margin
    
    # 左下に配置
    margin = 20
    qr_x = margin
    qr_y = img_height - total_height + bottom_margin
    
    # QRコードを画像に貼り付け
    img.paste(qr_img, (int(qr_x), int(qr_y)))
    
    # QRコードの下、中央揃えでテキストを配置
    text_x = qr_x + (qr_size - text_text_width) / 2
    text_y = qr_y + qr_size + text_margin
    
    # テキストを描画
    draw.text(
        (int(text_x), int(text_y)),
        text_id,
        fill=(0, 0, 0, 255),
        font=text_font
    )
    
    return img


def pdf_to_images(filename, username=None, student_name=None, student_number=None, text_name=None, campus_name=None, include_qr=False):
    """PDFを画像に変換"""
    # URLデコード
    filename = unquote(filename)
    base, ext = os.path.splitext(filename)
    if ext.lower() != ".pdf":
        raise ValueError("PDFファイルではありません")

    pdf_path = os.path.join(PDF_DIR, filename)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError("PDF が見つかりません")

    out_dir = os.path.join(CACHE_DIR, base)
    os.makedirs(out_dir, exist_ok=True)

    # キャッシュキーを生成（ユーザー名、生徒名、生徒番号、テキスト名、校舎名を含む）
    # バージョン19: 画面右上に「生徒名：○○　講師名：○○」の形式で表示、PDF内容全体を下と右にシフト
    cache_key = f"v19_{username or ''}_{student_name or ''}_{student_number or ''}_{text_name or ''}_{campus_name or ''}_{include_qr}"
    cache_suffix = ""
    if cache_key.strip():
        # ハッシュ値を生成してキャッシュサフィックスとして使用
        cache_hash = hashlib.md5(cache_key.encode('utf-8')).hexdigest()[:8]
        cache_suffix = f"_{cache_hash}"
    
    # 既存の PNG ファイルをチェック（キャッシュキーに基づく）
    if cache_suffix:
        existing = [f for f in os.listdir(out_dir) if f.lower().endswith(".png") and cache_suffix in f]
        if existing:
            existing.sort()
            return [os.path.join(out_dir, f) for f in existing]
    else:
        # キャッシュサフィックスがない場合（ユーザー名も生徒情報もない場合）
        existing = [f for f in os.listdir(out_dir) if f.lower().endswith(".png") and not "_" in f.replace("page_", "").replace(".png", "")]
        if existing:
            existing.sort()
            return [os.path.join(out_dir, f) for f in existing]

    # PDFを画像に変換
    images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
    image_paths = []
    # 印刷時の位置調整：PDF内容全体を下にシフトするための余白
    bottom_padding = 60  # 上に追加する余白（ピクセル）- 画像を下にシフトするため
    right_padding = 30  # 左に追加する余白（ピクセル）- 画像を右にシフトするため
    
    for i, img in enumerate(images, start=1):
        # 元の画像サイズを取得
        original_width, original_height = img.size
        
        # 上と左に余白を追加した新しい画像を作成（画像を下と右にシフトするため）
        new_img = Image.new('RGB', (original_width + right_padding, original_height + bottom_padding), color='white')
        # 元の画像を新しい画像の右下側に配置（上と左に余白ができ、画像が下と右にシフトされる）
        new_img.paste(img, (right_padding, bottom_padding))
        img = new_img  # 以降は新しい画像を使用
        
        # 1枚目でテキスト名がある場合、またはユーザー名/生徒情報が指定されている場合、テキストを描画
        if i == 1 and (username or student_name or student_number or text_name):
            try:
                draw = ImageDraw.Draw(img)
                img_width, img_height = img.size
                # フォントサイズを少し大きく（画面右上に表示）
                font_size = max(20, int(img_width / 80))
                
                font = None
                font_paths = [
                    "C:/Windows/Fonts/msgothic.ttc",
                    "C:/Windows/Fonts/meiryo.ttc",
                    "C:/Windows/Fonts/msmincho.ttc",
                    "arial.ttf"
                ]
                for font_path in font_paths:
                    try:
                        font = ImageFont.truetype(font_path, font_size)
                        break
                    except Exception:
                        continue
                
                if font is None:
                    font = ImageFont.load_default()
                
                # 画面右上に「生徒名：○○　講師名：○○」の形式で表示
                top_margin = 50  # 画面上端からの余白（印刷時のマージンを考慮して下げる）
                right_margin = 20  # 右端からの余白
                text_spacing = 15  # テキスト間のスペース
                
                # 表示するテキストを組み立て
                display_text_parts = []
                if student_name:
                    display_text_parts.append(f"生徒名：{student_name}")
                if username:
                    display_text_parts.append(f"講師名：{username}")
                
                if display_text_parts:
                    display_text = "　".join(display_text_parts)  # 全角スペースで区切る
                    
                    # テキストのサイズを取得
                    bbox = draw.textbbox((0, 0), display_text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                    
                    # 画面右上に配置（右揃え）
                    x_pos = img_width - text_width - right_margin
                    y_pos = top_margin
                    
                    # テキストを描画
                    draw.text(
                        (x_pos, y_pos),
                        display_text,
                        fill=(0, 0, 0, 255),
                        font=font
                    )
                
                # QRコードを生成して左下に配置（PRINT_ID形式）
                # ※QRコードにはPRINT_IDのみを含み、生徒名・講師名は含まない
                # include_qrがTrueの場合のみQRコードを表示（頭紙印刷時のみ）
                if include_qr and username and text_name:
                    try:
                        # PRINT_IDを生成（一意なID）
                        print_id = generate_print_id()
                        
                        # 元のファイル名を取得（filenameは既にunquote済み）
                        # 相対パスをそのまま使用（例: "算数/6年/数の性質_連続する数の積_応用.pdf"）
                        # パス区切り文字を統一（Windows形式をスラッシュに）
                        original_filename = filename.replace('\\', '/')
                        
                        # PRINT_IDとファイル名のマッピングを保存
                        save_print_id_mapping(print_id, original_filename)
                        
                        # QRコードのデータ: PRINT_ID=QS_YYYY_NNNNN,FILE=元のファイル名（URLエンコード）,PRINTER=プリンター名（校舎選択時のみ）
                        # 日本語ファイル名を正しく扱うため、URLエンコードしてから埋め込む
                        encoded_filename = quote(original_filename, safe='/')
                        qr_data = f"PRINT_ID={print_id},FILE={encoded_filename}"
                        
                        # 校舎が選択されている場合、プリンター名をQRコードに追加
                        if campus_name:
                            printer_name = get_printer_name_by_campus(campus_name)
                            if printer_name:
                                qr_data += f",PRINTER={printer_name}"
                        
                        # QRコードを生成
                        qr = qrcode.QRCode(
                            version=1,
                            error_correction=qrcode.constants.ERROR_CORRECT_L,
                            box_size=15,
                            border=4,
                        )
                        qr.add_data(qr_data)
                        qr.make(fit=True)
                        
                        # QRコード画像を生成
                        qr_img = qr.make_image(fill_color="black", back_color="white")
                        
                        # QRコードのサイズを調整（画像サイズの約20%）
                        qr_size = int(min(img_width, img_height) * 0.2)
                        qr_img = qr_img.resize((qr_size, qr_size), Image.Resampling.LANCZOS)
                        
                        # QRコードの下にテキストIDを表示するためのフォントを準備
                        # テキスト用フォント（QRコードより小さく）
                        text_font_size = max(14, int(img_width / 80))
                        text_font = None
                        for font_path in font_paths:
                            try:
                                text_font = ImageFont.truetype(font_path, text_font_size)
                                break
                            except Exception:
                                continue
                        if text_font is None:
                            text_font = ImageFont.load_default()
                        
                        # PRINT_IDのテキストのサイズを取得
                        text_id = print_id  # テキストIDとしてPRINT_IDを使用
                        text_bbox = draw.textbbox((0, 0), text_id, font=text_font)
                        text_text_width = text_bbox[2] - text_bbox[0]
                        text_text_height = text_bbox[3] - text_bbox[1]
                        
                        # テキストの高さを考慮してQRコードの位置を決定
                        bottom_margin = 15  # 画面下端との最小余白
                        text_margin = 10  # QRコードとテキストの間のマージン
                        total_height = qr_size + text_margin + text_text_height + bottom_margin
                        
                        # 左下に配置（マージンを考慮、テキスト分のスペースを確保）
                        margin = 20
                        qr_x = margin
                        qr_y = img_height - total_height + bottom_margin
                        
                        # QRコードを画像に貼り付け（一度だけ）
                        img.paste(qr_img, (int(qr_x), int(qr_y)))
                        
                        # QRコードの下、中央揃えでテキストを配置
                        text_x = qr_x + (qr_size - text_text_width) / 2
                        text_y = qr_y + qr_size + text_margin
                        
                        # テキストを描画
                        draw.text(
                            (int(text_x), int(text_y)),
                            text_id,
                            fill=(0, 0, 0, 255),
                            font=text_font
                        )
                        
                    except Exception as e:
                        import traceback
                        print(f"ERROR: QRコード生成エラー: {e}")
                        print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
                    
            except Exception as e:
                import traceback
                print(f"ERROR: テキスト描画エラー: {e}")
                print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
        
        img_name = f"page_{i}{cache_suffix}.png"
        img_path = os.path.join(out_dir, img_name)
        img.save(img_path, "PNG")
        image_paths.append(img_path)

    return image_paths


def get_folders_and_files(folder_path=""):
    """フォルダとPDFファイルを取得（Windows Explorerの順序でソート）"""
    full_path = os.path.join(PDF_DIR, folder_path) if folder_path else PDF_DIR
    
    if not os.path.exists(full_path):
        return [], []
    
    folders = []
    files = []
    
    try:
        # Windowsのロケール設定を使用して自然な順序でソート
        locale.setlocale(locale.LC_ALL, 'Japanese_Japan.932')
        
        for item in os.scandir(full_path):
            if item.is_dir():
                folders.append(item.name)
            elif item.name.lower().endswith(".pdf"):
                files.append(item.name)
        
        # ロケールベースの自然な順序でソート
        folders.sort(key=lambda x: locale.strxfrm(x))
        files.sort(key=lambda x: locale.strxfrm(x))
    except Exception:
        # ロケール設定に失敗した場合は通常のソート
        for item in os.scandir(full_path):
            if item.is_dir():
                folders.append(item.name)
            elif item.name.lower().endswith(".pdf"):
                files.append(item.name)
        folders.sort()
        files.sort()
    
    return folders, files


def get_all_pdf_files(subject_filter=""):
    """指定された科目のすべてのPDFファイルを再帰的に取得"""
    results = []  # [{"file_path": "算数/6年/file.pdf", "file_name": "file.pdf", "folder_path": "算数/6年"}, ...]
    
    def scan_directory(directory_path, relative_path=""):
        """ディレクトリを再帰的にスキャン"""
        if not os.path.exists(directory_path):
            return
        
        try:
            for item in os.scandir(directory_path):
                if item.is_dir():
                    # 科目フィルターが指定されている場合、最初の階層でフィルタリング
                    if subject_filter and not relative_path:
                        if item.name != subject_filter:
                            continue
                    new_relative_path = os.path.join(relative_path, item.name) if relative_path else item.name
                    new_relative_path = new_relative_path.replace('\\', '/')
                    scan_directory(item.path, new_relative_path)
                elif item.name.lower().endswith(".pdf"):
                    # 科目フィルターが指定されている場合、最初の階層でフィルタリング
                    if subject_filter and not relative_path:
                        # ファイルが直接PDF_DIRにある場合はスキップ（科目フォルダ内のファイルのみ）
                        continue
                    # パスを正規化（Windowsパス区切り文字を統一）
                    file_path = os.path.join(relative_path, item.name) if relative_path else item.name
                    file_path = file_path.replace('\\', '/')
                    results.append({
                        "file_path": file_path,
                        "file_name": item.name,
                        "folder_path": relative_path.replace('\\', '/') if relative_path else ""
                    })
        except Exception as e:
            print(f"Error scanning directory {directory_path}: {e}")
    
    scan_directory(PDF_DIR, "")
    return results


def get_students_file(username):
    """ユーザーごとの生徒ファイルパスを取得"""
    return os.path.join(STUDENTS_DIR, f"{username}.csv")


def load_students(username):
    """ユーザーごとの生徒リストを読み込む"""
    students = []
    students_file = get_students_file(username)
    
    if os.path.exists(students_file):
        try:
            with open(students_file, "r", encoding="utf-8", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    students.append({
                        "student_name": row.get("student_name", "").strip(),
                        "student_number": row.get("student_number", "").strip()
                    })
        except Exception as e:
            print(f"生徒データ読み込みエラー: {e}")
    
    return students


def save_students(username, students):
    """ユーザーごとの生徒リストを保存"""
    students_file = get_students_file(username)
    
    with open(students_file, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["student_name", "student_number"])
        writer.writeheader()
        for student in students:
            writer.writerow({
                "student_name": student["student_name"],
                "student_number": student.get("student_number", "")
            })


def load_file_name_history():
    """ファイル名変更履歴を読み込む"""
    history = {}  # {current_path: old_path}
    if os.path.exists(FILE_NAME_HISTORY_FILE):
        with open(FILE_NAME_HISTORY_FILE, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                old_path = normalize_file_path(row["old_path"])
                current_path = normalize_file_path(row["current_path"])
                if current_path not in history:
                    history[current_path] = []
                history[current_path].append(old_path)
    return history


def save_file_name_history(history):
    """ファイル名変更履歴を保存"""
    with open(FILE_NAME_HISTORY_FILE, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["old_path", "current_path"])
        writer.writeheader()
        for current_path, old_paths in history.items():
            for old_path in old_paths:
                writer.writerow({
                    "old_path": old_path,
                    "current_path": current_path
                })


def save_print_id_mapping(print_id: str, filename: str):
    """PRINT_IDとファイル名のマッピングを保存"""
    file_exists = os.path.exists(PRINT_ID_MAPPING_FILE)
    
    # 既存のマッピングを読み込む
    existing_mappings = {}
    if file_exists:
        try:
            with open(PRINT_ID_MAPPING_FILE, "r", encoding="utf-8", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    existing_mappings[row["print_id"]] = row["filename"]
        except Exception as e:
            print(f"マッピングファイル読み込みエラー: {e}")
    
    # 新しいマッピングを追加（既に存在する場合は更新）
    existing_mappings[print_id] = filename
    
    # 保存
    with open(PRINT_ID_MAPPING_FILE, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["print_id", "filename"])
        writer.writeheader()
        for pid, fname in existing_mappings.items():
            writer.writerow({
                "print_id": pid,
                "filename": fname
            })


def find_mappings_by_folder_and_index(folder_path, file_index, text_mappings, all_files_in_folder):
    """フォルダパスとファイルのインデックスを使ってマッピング情報を検索"""
    # 同じフォルダ内のファイルで、マッピング情報があるものを探す
    matched_mappings = []
    
    # フォルダパスが一致するマッピングを全て取得
    folder_mappings = {}
    for saved_path, mappings_list in text_mappings.items():
        saved_folder = '/'.join(saved_path.split('/')[:-1]) if '/' in saved_path else ''
        if saved_folder == folder_path or (not folder_path and not saved_folder):
            saved_filename = saved_path.split('/')[-1] if '/' in saved_path else saved_path
            folder_mappings[saved_filename] = mappings_list
    
    # フォルダ内のファイル数とマッピングがあるファイル数を比較
    # もしマッピングがあるファイルが1つだけで、現在のフォルダにもファイルが1つだけなら、それを引き継ぐ
    if len(all_files_in_folder) == 1 and len(folder_mappings) == 1:
        # 唯一のマッピングを引き継ぐ
        matched_mappings = list(folder_mappings.values())[0]
    
    return matched_mappings


def load_text_mappings():
    """テキスト対応情報を読み込む（正規化されたパスでマッピング）"""
    mappings = {}  # {file_path: [{"juku_name": "...", "text_name": "..."}, ...]}
    if os.path.exists(TEXT_MAPPING_FILE):
        with open(TEXT_MAPPING_FILE, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                file_path = row["file_path"]
                # 読み込み時も正規化して一貫性を保つ
                normalized_path = normalize_file_path(file_path)
                juku_name = row["juku_name"]
                text_name = row["text_name"]
                if normalized_path not in mappings:
                    mappings[normalized_path] = []
                mappings[normalized_path].append({
                    "juku_name": juku_name,
                    "text_name": text_name
                })
    return mappings


def save_text_mappings(mappings):
    """テキスト対応情報を保存"""
    with open(TEXT_MAPPING_FILE, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["file_path", "juku_name", "text_name"])
        writer.writeheader()
        for file_path, items in mappings.items():
            for item in items:
                writer.writerow({
                    "file_path": file_path,
                    "juku_name": item["juku_name"],
                    "text_name": item["text_name"]
                })


def add_text_mapping(file_path, juku_name, text_name):
    """テキスト対応情報を追加"""
    mappings = load_text_mappings()
    if file_path not in mappings:
        mappings[file_path] = []
    # 既に同じ組み合わせが存在しないかチェック
    for item in mappings[file_path]:
        if item["juku_name"] == juku_name and item["text_name"] == text_name:
            return  # 既に存在する場合は追加しない
    mappings[file_path].append({
        "juku_name": juku_name,
        "text_name": text_name
    })
    save_text_mappings(mappings)


def delete_text_mapping(file_path, juku_name, text_name):
    """テキスト対応情報を削除"""
    mappings = load_text_mappings()
    if file_path in mappings:
        mappings[file_path] = [
            item for item in mappings[file_path]
            if not (item["juku_name"] == juku_name and item["text_name"] == text_name)
        ]
        if not mappings[file_path]:
            del mappings[file_path]
        save_text_mappings(mappings)


@app.route("/login", methods=["GET", "POST"])
def login():
    """ログインページ"""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        
        users = load_users()
        if username in users and check_password_hash(users[username], password):
            session["username"] = username
            return redirect(url_for("index"))
        else:
            flash("ユーザー名またはパスワードが正しくありません", "error")
    
    return render_template("login.html")


@app.route("/logout")
def logout():
    """ログアウト"""
    session.pop("username", None)
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    """PDF一覧（フォルダ表示）"""
    # 空のパスでフォルダ表示を直接呼び出す
    return folder_view("")


def normalize_file_path(file_path):
    """ファイルパスを正規化（\を/に変換、連続するスラッシュを1つに、先頭・末尾のスラッシュを削除）"""
    normalized = file_path.replace('\\', '/')
    while '//' in normalized:
        normalized = normalized.replace('//', '/')
    return normalized.strip('/')


@app.route("/api/text-mapping", methods=["POST"])
@login_required
def add_text_mapping_api():
    """テキスト対応情報を追加するAPI"""
    data = request.get_json()
    file_path = data.get("file_path", "")
    juku_name = data.get("juku_name", "").strip()
    text_name = data.get("text_name", "").strip()
    
    if not file_path or not juku_name or not text_name:
        return {"success": False, "error": "必要な情報が不足しています"}, 400
    
    # ファイルパスを正規化
    file_path = normalize_file_path(file_path)
    
    add_text_mapping(file_path, juku_name, text_name)
    return {"success": True}


@app.route("/api/text-mapping", methods=["DELETE"])
@login_required
def delete_text_mapping_api():
    """テキスト対応情報を削除するAPI"""
    data = request.get_json()
    file_path = data.get("file_path", "")
    juku_name = data.get("juku_name", "").strip()
    text_name = data.get("text_name", "").strip()
    
    if not file_path or not juku_name or not text_name:
        return {"success": False, "error": "必要な情報が不足しています"}, 400
    
    # ファイルパスを正規化
    file_path = normalize_file_path(file_path)
    
    delete_text_mapping(file_path, juku_name, text_name)
    return {"success": True}


@app.route("/api/search", methods=["GET"])
@login_required
def search_files():
    """ファイル検索API"""
    subject = request.args.get("subject", "").strip()
    keyword = request.args.get("keyword", "").strip()
    
    if not keyword:
        return {"results": []}
    
    # すべてのPDFファイルを取得
    all_files = get_all_pdf_files(subject_filter=subject)
    
    # テキスト対応情報を読み込む
    text_mappings = load_text_mappings()
    
    results = []
    keyword_lower = keyword.lower()
    
    for file_info in all_files:
        file_path_normalized = normalize_file_path(file_info["file_path"])
        file_name = file_info["file_name"]
        folder_path = file_info["folder_path"]
        
        # ファイル名で部分一致検索
        matched = keyword_lower in file_name.lower()
        matching_text_mappings = []
        
        # テキストマッピング情報で検索
        if file_path_normalized in text_mappings:
            for mapping in text_mappings[file_path_normalized]:
                if keyword_lower in mapping["juku_name"].lower() or keyword_lower in mapping["text_name"].lower():
                    matched = True
                    matching_text_mappings.append(mapping)
        else:
            # ファイル名だけで検索（フォルダパスが異なる場合に対応）
            for saved_path, mappings_list in text_mappings.items():
                saved_filename = saved_path.split('/')[-1] if '/' in saved_path else saved_path
                if saved_filename == file_name:
                    for mapping in mappings_list:
                        if keyword_lower in mapping["juku_name"].lower() or keyword_lower in mapping["text_name"].lower():
                            matched = True
                            matching_text_mappings.append(mapping)
        
        if matched:
            results.append({
                "file_path": file_path_normalized,
                "file_name": file_name,
                "folder_path": folder_path,
                "matching_text_mappings": matching_text_mappings
            })
    
    return {"results": results}


@app.route("/folder/")
@app.route("/folder/<path:folder_path>")
@login_required
def folder_view(folder_path=""):
    """フォルダ表示"""
    # URLデコード
    decoded_folder_path = unquote(folder_path) if folder_path else ""
    folders, files = get_folders_and_files(decoded_folder_path)
    
    # 表示用には元のフォルダ名・ファイル名を使用
    # URL用にはエンコード
    encoded_folders = [quote(f, safe="") for f in folders]
    encoded_files = [quote(f, safe="") for f in files]
    
    # 現在のパスもエンコード（親フォルダへのリンク用）
    # 先頭・末尾のスラッシュを除去し、連続するスラッシュを正規化（パス結合用）
    if folder_path:
        # 連続するスラッシュを繰り返し1つに統一してから、先頭・末尾を削除
        normalized = folder_path
        while '//' in normalized:
            normalized = normalized.replace('//', '/')
        current_path_encoded = normalized.strip('/')
    else:
        current_path_encoded = ""
    
    # 表示用のパス（デコード済み、連続するスラッシュを正規化して先頭・末尾を除去）
    if decoded_folder_path:
        normalized_display = decoded_folder_path
        while '//' in normalized_display:
            normalized_display = normalized_display.replace('//', '/')
        current_path_display = normalized_display.strip('/')
    else:
        current_path_display = ""
    
    # テキスト対応情報を読み込む
    text_mappings = load_text_mappings()
    # ファイルパスをキーとして対応情報を取得
    file_text_mappings = {}
    for file in files:
        # ファイルパスを生成（フォルダパス + ファイル名）
        if decoded_folder_path:
            file_path = decoded_folder_path.replace('\\', '/') + '/' + file
        else:
            file_path = file
        # 正規化
        file_path = normalize_file_path(file_path)
        
        # 正規化されたファイルパスでマッピング情報を取得
        matched_mappings = []
        
        # 完全一致で検索
        if file_path in text_mappings:
            matched_mappings = text_mappings[file_path]
        else:
            # ファイル名だけで検索（フォルダパスが異なる場合に対応）
            matched_by_filename = False
            for saved_path, mappings_list in text_mappings.items():
                # ファイル名の部分だけを比較
                saved_filename = saved_path.split('/')[-1] if '/' in saved_path else saved_path
                # ファイル名が一致するか確認
                if saved_filename == file:
                    matched_mappings = mappings_list
                    matched_by_filename = True
                    break
            
            # ファイル名でマッチしなかった場合、フォルダ内で唯一のファイルなら引き継ぐ
            if not matched_by_filename:
                folder_path_for_search = decoded_folder_path if decoded_folder_path else ""
                folder_path_for_search = normalize_file_path(folder_path_for_search)
                folder_mappings = find_mappings_by_folder_and_index(folder_path_for_search, 0, text_mappings, files)
                if folder_mappings:
                    matched_mappings = folder_mappings
        
        file_text_mappings[file] = matched_mappings
    
    return render_template(
        "index.html",
        folders=folders,  # 表示用（デコード済み）
        files=files,  # 表示用（デコード済み）
        encoded_folders=encoded_folders,  # URL用（エンコード済み）
        encoded_files=encoded_files,  # URL用（エンコード済み）
        current_path=current_path_encoded,  # URL用（エンコード済み）
        current_path_display=current_path_display,  # 表示用（デコード済み）
        file_text_mappings=file_text_mappings,  # ファイルごとのテキスト対応情報
        username=session.get("username", "unknown")
    )


@app.route("/view/<path:filename>")
@login_required
def view(filename):
    """PDFを表示"""
    # セキュリティチェック
    if ".." in filename or filename.startswith("\\") or filename.startswith("/"):
        abort(400)

    # URLデコード
    decoded_filename = unquote(filename)
    pdf_path = os.path.join(PDF_DIR, decoded_filename)
    if not os.path.exists(pdf_path):
        abort(404, description="PDFファイルが見つかりません")

    user = get_current_user()
    
    # 生徒データを取得
    students = load_students(user)
    
    # クエリパラメータから生徒名を取得（選択された場合）
    selected_student_name = request.args.get("student_name", "")
    selected_student_number = ""
    if selected_student_name:
        for student in students:
            if student["student_name"] == selected_student_name:
                selected_student_number = student.get("student_number", "")
                break

    # テキスト名を取得（PDFファイル名から拡張子を除く）
    text_name = os.path.splitext(os.path.basename(decoded_filename))[0]

    try:
        image_paths = pdf_to_images(
            filename,
            username=user,
            student_name=selected_student_name if selected_student_name else None,
            student_number=selected_student_number if selected_student_number else None,
            text_name=text_name,  # 生徒名選択に関係なく常に渡す
            campus_name=None,  # 通常のプレビューでは校舎情報は不要
            include_qr=False  # 通常のプレビューではQRコードは不要
        )
    except Exception as e:
        return f"画像変換エラー: {e}", 500

    base, _ = os.path.splitext(decoded_filename)
    image_urls = []
    for p in image_paths:
        img_name = os.path.basename(p)
        # baseをURLエンコードしてから結合
        base_parts = base.split(os.sep)
        base_encoded = "/".join([quote(part, safe="") for part in base_parts])
        image_urls.append(f"/image/{base_encoded}/{quote(img_name, safe='')}")

    # 親フォルダのパスを取得（一つ前のフォルダ一覧に戻るため）
    parent_folder_path = ""
    if os.sep in decoded_filename or "/" in decoded_filename:
        # パス区切り文字で分割して、最後のファイル名を除く
        path_parts = decoded_filename.replace("\\", "/").split("/")
        if len(path_parts) > 1:
            parent_folder_path = "/".join(path_parts[:-1])
    
    return render_template(
        "view.html",
        username=user,
        filename=decoded_filename,
        image_urls=image_urls,
        students=students,
        selected_student_name=selected_student_name,
        parent_folder_path=parent_folder_path
    )


@app.route("/header-print")
@login_required
def header_print():
    """頭紙印刷ページ"""
    user = get_current_user()
    
    # プリンタ設定を読み込んで校舎リストを取得
    printer_config = load_printer_config()
    campuses = []
    # 校舎名のマッピング
    campus_name_mapping = {
        "yotsuya": "四谷校",
        "azabujuban": "麻布十番校",
        "yoyogi": "代々木校",
        "jiyugaoka": "自由が丘校",
        "kichijoji": "吉祥寺校",
        "tokyo": "東京校",
        "yokohama": "横浜校",
        "seijogakuen": "成城学園校",
        "ochanomizu": "お茶の水校",
        "kyotoekimae": "京都駅前校"
    }
    for campus_key, campus_config in printer_config.items():
        if campus_key in campus_name_mapping:
            campuses.append({
                "key": campus_key,
                "name": campus_name_mapping[campus_key]
            })
    
    # すべてのPDFファイルを取得
    pdf_files = get_all_pdf_files()
    
    return render_template(
        "header_print.html",
        username=user,
        campuses=campuses,
        pdf_files=pdf_files
    )


@app.route("/header/<path:filename>")
@login_required
def header(filename):
    """頭紙を表示・印刷（単一ファイル）"""
    # セキュリティチェック
    if ".." in filename or filename.startswith("\\") or filename.startswith("/"):
        abort(400)
    
    # URLデコード
    decoded_filename = unquote(filename)
    pdf_path = os.path.join(PDF_DIR, decoded_filename)
    if not os.path.exists(pdf_path):
        abort(404, description="PDFファイルが見つかりません")
    
    user = get_current_user()
    
    # テキスト名を取得（PDFファイル名から拡張子を除く）
    text_name = os.path.splitext(os.path.basename(decoded_filename))[0]
    
    # クエリパラメータから校舎名を取得（選択された場合）
    selected_campus_name = request.args.get("campus", "")
    
    try:
        # 頭紙PDFにQRコードを重ねて画像を生成
        header_img = create_header_with_qr(
            decoded_filename,
            username=user,
            text_name=text_name,
            campus_name=selected_campus_name if selected_campus_name else None
        )
        
        # 画像サイズを取得
        img_width, img_height = header_img.size
        
        # 画像を2倍のサイズに拡大（印刷時に200%倍率が必要な問題を解決）
        header_img_large = header_img.resize((img_width * 2, img_height * 2), Image.Resampling.LANCZOS)
        
        # 画像をキャッシュディレクトリに保存（DPI情報を設定）
        base, _ = os.path.splitext(decoded_filename)
        cache_dir = os.path.join(CACHE_DIR, base)
        os.makedirs(cache_dir, exist_ok=True)
        header_file_path = os.path.join(cache_dir, "header.png")
        # DPI情報を設定（72 DPI = 1インチ = 72ピクセル）
        # A4 landscape: 297mm x 210mm = 11.69インチ x 8.27インチ
        # 72 DPIの場合: 842px x 595px
        # 200 DPIの場合: 2339px x 1654px（pdf2imageのデフォルト）
        # pdf2imageのデフォルトDPI（200）に合わせる
        header_img_large.save(header_file_path, "PNG", dpi=(200, 200))
        
        # 画像URLを生成
        base_parts = base.split(os.sep)
        base_encoded = "/".join([quote(part, safe="") for part in base_parts])
        image_url = f"/image/{base_encoded}/header.png"
        
        # HTMLテンプレートで表示（画像サイズを渡す）
        return render_template("header.html", image_url=image_url, img_width=img_width, img_height=img_height)
    except Exception as e:
        import traceback
        print(f"ERROR: 頭紙生成エラー: {e}")
        print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
        return f"頭紙生成エラー: {e}", 500


@app.route("/headers-batch", methods=["POST"])
@login_required
def headers_batch():
    """複数の頭紙を一括表示・印刷"""
    user = get_current_user()
    
    # JSONデータからファイルリストと校舎名を取得
    try:
        data = request.get_json()
        if data is None:
            # JSONが送信されていない場合、フォームデータから取得を試みる
            data = request.form.to_dict()
            if "files" in data:
                import json
                data["files"] = json.loads(data["files"])
        
        if data is None:
            return jsonify({"error": "リクエストデータが無効です"}), 400
        
        file_paths = data.get("files", [])
        selected_campus_name = data.get("campus", "")
        
        if not file_paths or not selected_campus_name:
            return jsonify({"error": "ファイルと校舎を選択してください"}), 400
    except Exception as e:
        import traceback
        print(f"ERROR: リクエスト解析エラー: {e}")
        print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
        return jsonify({"error": f"リクエスト解析エラー: {str(e)}"}), 400
    
    image_urls = []
    img_width = None
    img_height = None
    
    try:
        for file_path in file_paths:
            # セキュリティチェック
            if ".." in file_path or file_path.startswith("\\") or file_path.startswith("/"):
                continue
            
            decoded_filename = unquote(file_path)
            pdf_path = os.path.join(PDF_DIR, decoded_filename)
            if not os.path.exists(pdf_path):
                continue
            
            # テキスト名を取得
            text_name = os.path.splitext(os.path.basename(decoded_filename))[0]
            
            # 頭紙PDFにQRコードを重ねて画像を生成
            header_img = create_header_with_qr(
                decoded_filename,
                username=user,
                text_name=text_name,
                campus_name=selected_campus_name if selected_campus_name else None
            )
            
            # 画像サイズを取得（最初の1つを基準にする）
            if img_width is None:
                img_width, img_height = header_img.size
            
            # 画像を2倍のサイズに拡大
            header_img_large = header_img.resize((img_width * 2, img_height * 2), Image.Resampling.LANCZOS)
            
            # 画像をキャッシュディレクトリに保存
            base, _ = os.path.splitext(decoded_filename)
            cache_dir = os.path.join(CACHE_DIR, base)
            os.makedirs(cache_dir, exist_ok=True)
            header_file_path = os.path.join(cache_dir, "header.png")
            header_img_large.save(header_file_path, "PNG", dpi=(200, 200))
            
            # 画像URLを生成
            base_parts = base.split(os.sep)
            base_encoded = "/".join([quote(part, safe="") for part in base_parts])
            image_url = f"/image/{base_encoded}/header.png"
            image_urls.append(image_url)
        
        if not image_urls:
            return jsonify({"error": "有効なファイルが見つかりませんでした"}), 400
        
        # 一括印刷用のURLを生成（クエリパラメータで画像URLを渡す）
        # または、セッションに保存してからリダイレクト
        import base64
        import json as json_module
        
        # 画像URLをJSONエンコードしてbase64エンコード
        image_urls_json = json_module.dumps(image_urls)
        image_urls_encoded = base64.urlsafe_b64encode(image_urls_json.encode('utf-8')).decode('utf-8')
        
        # 一括印刷ページのURLを返す
        batch_url = f"/headers-batch-view?images={image_urls_encoded}&width={img_width}&height={img_height}"
        
        return jsonify({"url": batch_url})
    except Exception as e:
        import traceback
        print(f"ERROR: 頭紙一括生成エラー: {e}")
        print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
        return f"頭紙一括生成エラー: {e}", 500


@app.route("/headers-batch-view")
@login_required
def headers_batch_view():
    """一括印刷用の頭紙を表示"""
    import base64
    import json as json_module
    
    # クエリパラメータから画像URLを取得
    images_encoded = request.args.get("images", "")
    if not images_encoded:
        abort(400, description="画像データが指定されていません")
    
    try:
        # base64デコードしてJSONを取得
        image_urls_json = base64.urlsafe_b64decode(images_encoded.encode('utf-8')).decode('utf-8')
        image_urls = json_module.loads(image_urls_json)
        
        # 画像サイズを取得（オプション）
        img_width = request.args.get("width", type=int)
        img_height = request.args.get("height", type=int)
        
        return render_template("header_batch.html", image_urls=image_urls, img_width=img_width, img_height=img_height)
    except Exception as e:
        import traceback
        print(f"ERROR: 頭紙一括表示エラー: {e}")
        print(f"ERROR: トレースバック:\n{traceback.format_exc()}")
        return f"頭紙一括表示エラー: {e}", 500


@app.route("/image/<path:base>/<path:img_name>")
def image(base, img_name):
    """画像を返す"""
    # セキュリティチェック
    if ".." in base or ".." in img_name:
        abort(400)
    
    # URLデコード
    base_decoded = unquote(base)
    img_name_decoded = unquote(img_name)
    
    dir_path = os.path.join(CACHE_DIR, base_decoded)
    img_path = os.path.join(dir_path, img_name_decoded)

    if not os.path.exists(img_path):
        abort(404)

    return send_file(img_path, mimetype="image/png")


@app.route("/download/<path:filename>")
@login_required
def download_pdf(filename):
    """PDFファイルをダウンロード"""
    # セキュリティチェック
    if ".." in filename or filename.startswith("\\") or filename.startswith("/"):
        abort(400)
    
    # URLデコード
    decoded_filename = unquote(filename)
    pdf_path = os.path.join(PDF_DIR, decoded_filename)
    
    if not os.path.exists(pdf_path):
        abort(404, description="PDFファイルが見つかりません")
    
    # ファイル名を取得（パス区切り文字をアンダースコアに変換してファイル名として使用）
    safe_filename = os.path.basename(decoded_filename)
    if not safe_filename:
        safe_filename = "download.pdf"
    
    return send_file(
        pdf_path,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=safe_filename
    )


@app.route("/log_print", methods=["POST"])
@login_required
def log_print():
    """印刷ログを記録"""
    user = get_current_user()
    filename = request.form.get("filename", "")
    copies = request.form.get("copies", "1")
    student_name = request.form.get("student_name", "")
    client_ip = request.remote_addr or ""

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    file_exists = os.path.exists(LOG_FILE)
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["printed_at", "user", "filename", "copies", "student_name", "client_ip"])
        writer.writerow([now, user, filename, copies, student_name, client_ip])

    return "OK"


@app.route("/logs")
@login_required
def logs():
    """印刷ログを表示"""
    log_entries = []
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            log_entries = list(reader)
            log_entries.reverse()  # 新しい順に
    
    return render_template("logs.html", logs=log_entries, username=session.get("username", "unknown"))


@app.route("/students", methods=["GET", "POST"])
@login_required
def students():
    """生徒登録ページ"""
    username = get_current_user()
    
    if request.method == "POST":
        action = request.form.get("action", "add")
        student_name = request.form.get("student_name", "").strip()
        student_number = request.form.get("student_number", "").strip()
        
        if not student_name:
            flash("生徒名を入力してください。", "error")
            students_list = load_students(username)
            return render_template("students.html", students=students_list, username=username)
        
        students_list = load_students(username)
        
        if action == "add":
            # 重複チェック
            if any(s["student_name"] == student_name for s in students_list):
                flash(f"生徒「{student_name}」は既に登録されています。", "error")
            else:
                students_list.append({
                    "student_name": student_name,
                    "student_number": student_number
                })
                save_students(username, students_list)
                flash(f"生徒「{student_name}」を登録しました。", "success")
        
        elif action == "edit":
            # 既存の生徒を更新
            found = False
            for student in students_list:
                if student["student_name"] == student_name:
                    student["student_number"] = student_number
                    found = True
                    break
            
            if found:
                save_students(username, students_list)
                flash(f"生徒「{student_name}」を更新しました。", "success")
            else:
                flash(f"生徒「{student_name}」が見つかりません。", "error")
        
        elif action == "delete":
            # 生徒を削除
            original_name = request.form.get("student_name", "").strip()
            students_list = [s for s in students_list if s["student_name"] != original_name]
            save_students(username, students_list)
            flash(f"生徒「{original_name}」を削除しました。", "success")
    
    # GETリクエストまたはPOST処理後の表示
    students_list = load_students(username)
    return render_template("students.html", students=students_list, username=username)


@app.route("/logo")
def logo():
    """ロゴ画像を返す"""
    logo_path = os.path.join(BASE_DIR, "qslogo.png")
    if os.path.exists(logo_path):
        return send_file(logo_path, mimetype="image/png")
    abort(404)


@app.route("/favicon.ico")
def favicon():
    """ファビコン"""
    # ファビコンが存在しない場合は静かに404を返す（エラーログを出さない）
    from flask import Response
    return Response(status=404)


@app.errorhandler(500)
def internal_error(error):
    """内部サーバーエラーのハンドラー"""
    import traceback
    error_msg = str(error)
    traceback_str = traceback.format_exc()
    print(f"ERROR: 内部サーバーエラー: {error_msg}")
    print(f"ERROR: トレースバック:\n{traceback_str}")
    return jsonify({"error": "内部サーバーエラーが発生しました。詳細はサーバーログを確認してください。"}), 500


@app.errorhandler(Exception)
def handle_exception(e):
    """すべての例外をキャッチするハンドラー"""
    import traceback
    error_msg = str(e)
    traceback_str = traceback.format_exc()
    print(f"ERROR: 予期しないエラー: {error_msg}")
    print(f"ERROR: トレースバック:\n{traceback_str}")
    
    # JSONリクエストの場合はJSONで返す
    if request.is_json or request.path.startswith('/api') or request.path.startswith('/headers-batch'):
        return jsonify({"error": f"エラーが発生しました: {error_msg}"}), 500
    
    # それ以外はHTMLエラーページを返す
    return f"エラーが発生しました: {error_msg}", 500


if __name__ == "__main__":
    # 開発用: IISではwfastcgiを使用
    debug_mode = os.environ.get("FLASK_DEBUG", "False").lower() == "true"
    app.run(host="0.0.0.0", port=5000, debug=debug_mode)
