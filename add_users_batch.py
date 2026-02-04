# -*- coding: utf-8 -*-
"""ユーザー一括追加スクリプト。printviewer のプロジェクトルートで実行: python add_users_batch.py"""
import os
import csv
from pathlib import Path
from werkzeug.security import generate_password_hash

BASE_DIR = Path(__file__).resolve().parent
USERS_FILE = BASE_DIR / "users.csv"

# 追加するユーザー: (username, password, is_admin)
# 〇 がついているのが管理者
NEW_USERS = [
    ("久米", "ikuno", False),
    ("吉野", "iwase", False),
    ("合田", "ogita", False),
    ("勝山", "gouukon", False),
    ("坂井", "kotsuta", False),
    ("桑田", "suganuma", False),
    ("清水", "takeuchi", False),
    ("高田", "takeda", False),
    ("江田", "tazawa", False),
    ("永田", "nagaosa", False),
    ("佐倉", "natsume", False),
    ("亀井", "niikawa", False),
    ("佐々木", "barada", False),
    ("安部", "higashi", True),
    ("新井", "maie", False),
    ("川上", "mae", False),
    ("海田", "matsuura", True),
    ("松西", "matsubishi", False),
    ("松田", "matsumoto", False),
    ("天野", "mizuta", False),
    ("白石", "mineoka", False),
    ("千葉", "mori", False),
    ("吉岡", "yamaguchi", False),
    ("大木", "yamasaki", False),
    ("咲山", "yamazaki", False),
    ("太田", "yokota", False),
    ("広瀬", "ikeda", False),
    ("長門", "takahashi", False),
    ("伊藤", "takeda", False),
    ("小林", "naitou", False),
    ("森", "nakazawa", False),
    ("吉田", "nagase", False),
    ("高倉", "nukanobu", False),
]


def load_users():
    users = {}
    if USERS_FILE.exists():
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if "username" in row and "password_hash" in row:
                    is_admin = row.get("is_admin", "0").strip() == "1"
                    users[row["username"]] = {
                        "password_hash": row["password_hash"],
                        "is_admin": is_admin,
                    }
    return users


def save_all_users(users):
    with open(USERS_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["username", "password_hash", "is_admin"])
        for user, data in users.items():
            writer.writerow([user, data["password_hash"], "1" if data["is_admin"] else "0"])


def main():
    users = load_users()
    added = 0
    skipped = 0
    for username, password, is_admin in NEW_USERS:
        if username in users:
            print(f"スキップ（既存）: {username}")
            skipped += 1
            continue
        users[username] = {
            "password_hash": generate_password_hash(password),
            "is_admin": is_admin,
        }
        admin_mark = " [管理者]" if is_admin else ""
        print(f"追加: {username} {admin_mark}")
        added += 1
    save_all_users(users)
    print(f"\n完了: {added} 件追加, {skipped} 件スキップ（既存）")


if __name__ == "__main__":
    main()
