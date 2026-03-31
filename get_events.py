import os
import requests
from dotenv import load_dotenv

# .env ファイルから環境変数を読み込む
load_dotenv()

# Azure AD アプリ登録で取得した認証情報
CLIENT_ID = os.getenv("CLIENT_ID")        # アプリケーション (クライアント) ID
TENANT_ID = os.getenv("TENANT_ID")        # ディレクトリ (テナント) ID
CLIENT_SECRET = os.getenv("CLIENT_SECRET") # クライアントシークレット
USER_EMAIL = os.getenv("SENDER_EMAIL")     # カレンダーを取得する対象ユーザーのメールアドレス


def get_access_token() -> str:
    """
    Microsoft Graph API へアクセスするためのトークンを取得する。
    クライアントクレデンシャルフロー（アプリ認証）を使用する。
    """
    # トークン取得エンドポイント（テナント固有）
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",  # Graph API 全般のスコープ
        "grant_type": "client_credentials",               # アプリ認証フロー
    }
    response = requests.post(url, data=data, timeout=30)
    response.raise_for_status()  # 認証失敗時は例外を発生させる
    return response.json()["access_token"]


def get_events(user_email: str, start_datetime: str, end_datetime: str) -> list:
    """
    指定期間のカレンダーイベントを取得する。
    start_datetime, end_datetime は ISO 8601 形式 (例: "2026-03-31T00:00:00")
    繰り返しイベントも含めて取得するため calendarView を使用する。
    """
    # まずアクセストークンを取得
    access_token = get_access_token()

    # calendarView エンドポイント：指定期間内に発生するイベントを返す（繰り返しイベント含む）
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/calendarView"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    params = {
        "startDateTime": start_datetime,          # 取得開始日時
        "endDateTime": end_datetime,              # 取得終了日時
        "$select": "subject,start,end,location,organizer",  # 取得するフィールドを絞る
        "$orderby": "start/dateTime",             # 開始日時の昇順で並べる
        "$top": 50,                               # 1 回のリクエストで最大 50 件取得
    }

    events = []
    # ページネーション対応：全件取得できるまでループ
    while url:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()  # HTTP エラー時は例外を発生させる
        data = response.json()
        events.extend(data.get("value", []))  # 取得したイベントをリストに追加
        # 次ページがあれば続けて取得（@odata.nextLink にはパラメータが含まれるため params は不要）
        url = data.get("@odata.nextLink")
        params = {}  # nextLink 使用時はパラメータを空にする（URL に含まれているため）

    return events


if __name__ == "__main__":
    # 取得したい期間を指定（ISO 8601 形式）
    start = "2026-03-01T00:00:00"
    end = "2026-03-31T23:59:59"

    events = get_events(USER_EMAIL, start, end)

    # 取得結果を一覧表示
    print(f"取得件数: {len(events)} 件")
    for event in events:
        subject = event.get("subject", "(件名なし)")
        start_dt = event.get("start", {}).get("dateTime", "")
        end_dt = event.get("end", {}).get("dateTime", "")
        location = event.get("location", {}).get("displayName", "")
        print(f"  [{start_dt} - {end_dt}] {subject} / 場所: {location}")
