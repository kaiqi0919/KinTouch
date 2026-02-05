# 出退勤管理システム - 定数定義

from datetime import timezone, timedelta

# 日本時間のタイムゾーン設定
JST = timezone(timedelta(hours=9))

# パスワードのハッシュ値（SHA-256）
PASSWORD_HASH = "944381cba581a7ee3f59b5e2a97686c9b48fc0b7da14eb3fceff5f84f62d0ff7"

# ディレクトリ設定
DATA_DIR = "data"
CONFIG_PATH = "reader_config.json"

# ウィンドウ設定
WINDOW_WIDTH = 900
WINDOW_HEIGHT = 600