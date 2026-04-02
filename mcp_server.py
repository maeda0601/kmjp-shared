# /// script
# requires-python = ">=3.12"
# dependencies = [
#   "mcp[cli]",
#   "requests",
#   "python-dotenv",
# ]
# ///

import os
import requests
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

# .env ファイルから環境変数を読み込む
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
USER_EMAIL = os.getenv("SENDER_EMAIL")

mcp = FastMCP("Outlook Calendar & Mail")


def get_access_token() -> str:
    """Microsoft Graph API へアクセスするためのトークンを取得する。"""
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


@mcp.tool()
def create_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    timezone: str = "Asia/Tokyo",
    location: str = "",
    body_text: str = "",
    attendees: list[str] | None = None,
) -> dict:
    """
    Outlookカレンダーに予定を追加する。

    Parameters:
    - subject: 件名
    - start_datetime: 開始日時（ISO 8601形式, 例: "2026-04-01T10:00:00"）
    - end_datetime: 終了日時（ISO 8601形式, 例: "2026-04-01T11:00:00"）
    - timezone: タイムゾーン（デフォルト: "Asia/Tokyo"）
    - location: 場所（省略可）
    - body_text: 本文・メモ（省略可）
    - attendees: 招待するメールアドレスのリスト（省略可）
    """
    access_token = get_access_token()

    url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/events"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    payload: dict = {
        "subject": subject,
        "start": {"dateTime": start_datetime, "timeZone": timezone},
        "end": {"dateTime": end_datetime, "timeZone": timezone},
    }

    if location:
        payload["location"] = {"displayName": location}
    if body_text:
        payload["body"] = {"contentType": "Text", "content": body_text}
    if attendees:
        payload["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]

    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()

    event = response.json()
    return {
        "id": event.get("id"),
        "subject": event.get("subject"),
        "start": event.get("start", {}).get("dateTime"),
        "end": event.get("end", {}).get("dateTime"),
    }


@mcp.tool()
def get_events(start_datetime: str, end_datetime: str) -> list[dict]:
    """
    指定期間のOutlookカレンダーイベントを取得する。

    Parameters:
    - start_datetime: 取得開始日時（ISO 8601形式, 例: "2026-04-01T00:00:00"）
    - end_datetime: 取得終了日時（ISO 8601形式, 例: "2026-04-30T23:59:59"）
    """
    access_token = get_access_token()

    url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/calendarView"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    params = {
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
        "$select": "subject,start,end,location,organizer",
        "$orderby": "start/dateTime",
        "$top": 50,
    }

    events = []
    while url:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()
        events.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = {}

    return [
        {
            "subject": e.get("subject", ""),
            "start": e.get("start", {}).get("dateTime", ""),
            "end": e.get("end", {}).get("dateTime", ""),
            "location": e.get("location", {}).get("displayName", ""),
        }
        for e in events
    ]


@mcp.tool()
def send_mail(to_email: str, subject: str, body_text: str) -> str:
    """
    指定した宛先にメールを送信する。

    Parameters:
    - to_email: 宛先メールアドレス
    - subject: 件名
    - body_text: 本文（テキスト形式）
    """
    access_token = get_access_token()

    url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": True,
    }

    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    return "メール送信に成功しました。"


if __name__ == "__main__":
    mcp.run()
