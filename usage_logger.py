"""
アプリの利用ログを記録するモジュール
記録する内容：使用日時・物件名・担当者名・作成した形式(PDF/Word)・アクセス元IPアドレス
ログは logs/usage_log.csv に追記され、Excelでもそのまま開けるようUTF-8(BOM付き)で保存する。
"""
import os
import csv
from datetime import datetime

LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
LOG_PATH = os.path.join(LOG_DIR, "usage_log.csv")
LOG_HEADER = ["日時", "物件名", "担当者名", "作成した形式", "IPアドレス"]


def get_client_ip():
    """アクセスしているパソコンのIPアドレスを取得する。

    Streamlitが公式に提供している機能ではなく、内部の仕組みを利用しているため、
    将来Streamlitのバージョンが上がった際に取得できなくなる可能性がある。
    取得できない場合はアプリを止めずに「不明」を返す。
    """
    try:
        from streamlit import runtime
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        ctx = get_script_run_ctx()
        if ctx is None:
            return "不明"
        session_info = runtime.get_instance().get_client(ctx.session_id)
        if session_info is None:
            return "不明"
        return session_info.request.remote_ip
    except Exception:
        return "不明"


def log_usage(property_name, contact_name, formats):
    """利用ログを1行追記する
    property_name: 物件名
    contact_name: 担当者名(空文字の場合は「(未入力)」と記録)
    formats: 作成した形式のリスト 例：["PDF"] や ["Word"]
    """
    os.makedirs(LOG_DIR, exist_ok=True)
    is_new = not os.path.exists(LOG_PATH)

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        property_name,
        contact_name.strip() if contact_name.strip() else "(未入力)",
        "・".join(formats),
        get_client_ip(),
    ]

    with open(LOG_PATH, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(LOG_HEADER)
        writer.writerow(row)


def read_log():
    """記録済みの利用ログを読み込む。
    戻り値: (ヘッダーのリスト, データ行のリスト(新しい順)) 。記録が無ければ (None, [])
    """
    if not os.path.exists(LOG_PATH):
        return None, []
    with open(LOG_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return None, []
    header, data = rows[0], rows[1:]
    return header, list(reversed(data))
