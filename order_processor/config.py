# 設定ファイル - ここを環境に合わせて変更してください
import os

# このファイル（config.py）のある場所を基準にパスを解決する
# → どのフォルダから実行しても同じDBと出力先を使う
_HERE = os.path.dirname(os.path.abspath(__file__))

# メール設定
EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",      # SMTPサーバー（Gmailの場合）
    "smtp_port": 587,
    "sender_address": "your_email@gmail.com",   # 送信元メールアドレス
    "sender_password": "your_app_password",     # アプリパスワード
    "recipients": {
        "田村さん": "tamura@example.com",        # 田村さんのメールアドレス
        "脇阪さん": "wakasaka@example.com",      # 脇阪さんのメールアドレス
    },
}

# ファイルパス設定
MASTER_EXCEL_PATH = r"C:\Users\YourName\Documents\構成一覧.xlsx"   # 構成一覧Excelのパス
DB_PATH     = os.path.join(_HERE, "data", "master.db")              # SQLiteデータベースのパス
OUTPUT_DIR  = os.path.join(_HERE, "output")                         # 出力ファイルの保存先
