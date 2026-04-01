import os
import requests
from dotenv import load_dotenv

# .env ファイルから環境変数を読み込む
load_dotenv()

# Azure AD アプリ登録で取得した認証情報
CLIENT_ID = os.getenv("CLIENT_ID")        # アプリケーション (クライアント) ID
TENANT_ID = os.getenv("TENANT_ID")        # ディレクトリ (テナント) ID
CLIENT_SECRET = os.getenv("CLIENT_SECRET") # クライアントシークレット
USER_EMAIL = os.getenv("SENDER_EMAIL")     # 予定を追加する対象ユーザーのメールアドレス


def get_access_token() -> str:
    """
    Microsoft Graph API へアクセスするためのトークンを取得する。
    クライアントクレデンシャルフロー（アプリ認証）を使用する。
    """
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    response = requests.post(url, data=data, timeout=30)
    response.raise_for_status()
    return response.json()["access_token"]


def create_event(
    user_email: str,
    subject: str,
    start_datetime: str,
    end_datetime: str,
    timezone: str = "Asia/Tokyo",
    location: str = "",
    body_text: str = "",
    attendees: list[str] | None = None,
) -> dict:
    """
    指定ユーザーのカレンダーに予定を追加する。

    Parameters
    ----------
    user_email    : 予定を追加するユーザーのメールアドレス
    subject       : 件名
    start_datetime: 開始日時（ISO 8601 形式, 例: "2026-04-01T10:00:00"）
    end_datetime  : 終了日時（ISO 8601 形式, 例: "2026-04-01T11:00:00"）
    timezone      : タイムゾーン（デフォルト: "Asia/Tokyo"）
    location      : 場所（省略可）
    body_text     : 本文・メモ（省略可）
    attendees     : 招待するメールアドレスのリスト（省略可）

    Returns
    -------
    作成されたイベントの JSON レスポンス（dict）
    """
    access_token = get_access_token()

    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/events"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    # イベントの基本情報
    payload: dict = {
        "subject": subject,
        "start": {
            "dateTime": start_datetime,
            "timeZone": timezone,
        },
        "end": {
            "dateTime": end_datetime,
            "timeZone": timezone,
        },
    }

    # 場所が指定された場合のみ追加
    if location:
        payload["location"] = {"displayName": location}

    # 本文が指定された場合のみ追加
    if body_text:
        payload["body"] = {"contentType": "Text", "content": body_text}

    # 招待者が指定された場合のみ追加
    if attendees:
        payload["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]

    response = requests.post(url, headers=headers, json=payload, timeout=30)
    print("create_event status:", response.status_code)

    if response.status_code == 201:
        print("予定の追加に成功しました。")
    else:
        print("予定の追加に失敗しました。")
        print("body:", response.text)
        response.raise_for_status()

    return response.json()


if __name__ == "__main__":
    # 追加したい予定の情報を指定
    event = create_event(
        user_email=USER_EMAIL,
        subject="打合せｆｆｆｆ",
        start_datetime="2026-07-10T10:00:00",
        end_datetime="2026-07-10T11:00:00",
        timezone="Asia/Tokyo",
        location="",
        body_text="",
        attendees=[],  # 招待者のメールアドレスを追加する場合はここに記述
    )

    # 作成したイベントの ID と件名を表示
    print(f"作成されたイベント ID: {event.get('id')}")
    print(f"件名: {event.get('subject')}")
    print(f"開始: {event.get('start', {}).get('dateTime')}")
    print(f"終了: {event.get('end', {}).get('dateTime')}")
