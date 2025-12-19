import os
import csv
import re
import locale
import hashlib
from datetime import datetime
from urllib.parse import quote, unquote
from functools import wraps

from flask import Flask, render_template, send_file, request, abort, session, redirect, url_for, flash
from pdf2image import convert_from_path
from werkzeug.security import generate_password_hash, check_password_hash
from PIL import ImageDraw, ImageFont, Image
import qrcode

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "your-secret-key-change-this-in-production")

# 設定: 環境変数で上書き可能、なければローカルパスを使用
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PDF_DIR = os.path.join(BASE_DIR, "test_pdfs")
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
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                users[row["username"]] = row["password_hash"]
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


def pdf_to_images(filename, username=None, student_name=None, student_number=None, text_name=None):
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

    # キャッシュキーを生成（ユーザー名、生徒名、生徒番号、テキスト名を含む）
    # バージョン4: 黄色背景なし、画面下中央配置、生徒番号0.61、QRコード追加
    cache_key = f"v4_{username or ''}_{student_name or ''}_{student_number or ''}_{text_name or ''}"
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
    for i, img in enumerate(images, start=1):
        # 1枚目でユーザー名または生徒情報が指定されている場合、テキストを描画
        if i == 1 and (username or student_name or student_number):
            try:
                draw = ImageDraw.Draw(img)
                img_width, img_height = img.size
                # フォントサイズを下げる
                font_size = max(20, int(img_width / 60))
                
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
                
                # 生徒番号を画面下中央の0.61の位置に描画
                if student_number:
                    student_number_text = student_number  # 「生徒番号：」を削除
                    bbox = draw.textbbox((0, 0), student_number_text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                    
                    # 画面下中央の0.61の位置（中央揃え）
                    x_pos = (img_width - text_width) / 2
                    y_pos = int(img_height * 0.61) - text_height / 2
                    
                    # テキストを描画（背景なし）
                    draw.text(
                        (x_pos, y_pos),
                        student_number_text,
                        fill=(0, 0, 0, 255),
                        font=font
                    )
                
                # ユーザー名を画面下中央の0.73の位置に描画
                if username:
                    username_text = username  # 「ユーザー：」を削除
                    bbox = draw.textbbox((0, 0), username_text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                    
                    # 画面下中央の0.73の位置（中央揃え）
                    x_pos = (img_width - text_width) / 2
                    y_pos = int(img_height * 0.73) - text_height / 2
                    
                    # テキストを描画（背景なし）
                    draw.text(
                        (x_pos, y_pos),
                        username_text,
                        fill=(0, 0, 0, 255),
                        font=font
                    )
                
                # QRコードを生成して左下に配置
                if student_name and username and text_name:
                    try:
                        # QRコードのデータ: 生徒名,講師名,テキスト名
                        qr_data = f"{student_name},{username},{text_name}"
                        
                        # QRコードを生成
                        qr = qrcode.QRCode(
                            version=1,
                            error_correction=qrcode.constants.ERROR_CORRECT_L,
                            box_size=10,
                            border=4,
                        )
                        qr.add_data(qr_data)
                        qr.make(fit=True)
                        
                        # QRコード画像を生成
                        qr_img = qr.make_image(fill_color="black", back_color="white")
                        
                        # QRコードのサイズを調整（画像サイズの約10%）
                        qr_size = int(min(img_width, img_height) * 0.1)
                        qr_img = qr_img.resize((qr_size, qr_size), Image.Resampling.LANCZOS)
                        
                        # 左下に配置（マージンを考慮）
                        margin = 20
                        qr_x = margin
                        qr_y = img_height - qr_size - margin
                        
                        # QRコードを画像に貼り付け
                        img.paste(qr_img, (qr_x, qr_y))
                        
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
        
        # デバッグ用：最初の数個のキーを出力
        if file == files[0] if files else None:
            sample_keys = list(text_mappings.keys())[:3]
            print(f"DEBUG: Looking for file_path='{file_path}', file='{file}'")
            print(f"DEBUG: Sample saved paths: {sample_keys}")
        
        # 完全一致で検索
        if file_path in text_mappings:
            matched_mappings = text_mappings[file_path]
            print(f"DEBUG: Exact match found for '{file_path}', mappings count: {len(matched_mappings)}")
            print(f"DEBUG: Mappings content: {matched_mappings}")
        else:
            # ファイル名だけで検索（フォルダパスが異なる場合に対応）
            matched_by_filename = False
            for saved_path, mappings_list in text_mappings.items():
                # ファイル名の部分だけを比較
                saved_filename = saved_path.split('/')[-1] if '/' in saved_path else saved_path
                # ファイル名が一致するか確認
                if saved_filename == file:
                    matched_mappings = mappings_list
                    print(f"DEBUG: Matched by filename: saved_path='{saved_path}', saved_filename='{saved_filename}', current_file='{file}', mappings count: {len(matched_mappings)}")
                    print(f"DEBUG: Mappings content: {matched_mappings}")
                    matched_by_filename = True
                    break
            
            # ファイル名でマッチしなかった場合、フォルダ内で唯一のファイルなら引き継ぐ
            if not matched_by_filename:
                folder_path_for_search = decoded_folder_path if decoded_folder_path else ""
                folder_path_for_search = normalize_file_path(folder_path_for_search)
                folder_mappings = find_mappings_by_folder_and_index(folder_path_for_search, 0, text_mappings, files)
                if folder_mappings:
                    matched_mappings = folder_mappings
                    print(f"DEBUG: Matched by folder (single file): folder='{folder_path_for_search}', mappings count: {len(matched_mappings)}")
        
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
            text_name=text_name if selected_student_name else None
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

    return render_template(
        "view.html",
        username=user,
        filename=decoded_filename,
        image_urls=image_urls,
        students=students,
        selected_student_name=selected_student_name,
    )


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
    abort(404)


if __name__ == "__main__":
    # 開発用: IISではwfastcgiを使用
    debug_mode = os.environ.get("FLASK_DEBUG", "False").lower() == "true"
    app.run(host="0.0.0.0", port=5000, debug=debug_mode)
