# Microsoft Graph API ツール集

Microsoft Graph API を使用して、Azure AD アプリケーションからメール送信・カレンダーイベント取得を行う Python スクリプトです。

## 概要

OAuth2 クライアント資格情報フロー（アプリ認証）で認証を行い、以下の操作が可能です。

- **メール送信** (`app.py`): 指定したユーザーのメールボックスからメールを送信
- **カレンダー取得** (`get_events.py`): 指定したユーザーのカレンダーイベントを期間指定で取得（繰り返しイベント含む）

## 必要要件

- Python 3.x
- 以下のパッケージ:
  - `requests`
  - `python-dotenv`

```bash
pip install requests python-dotenv
```

## セットアップ

### Azure AD アプリの準備

1. Azure ポータルでアプリを登録
2. 以下のアプリケーション権限を付与（委任ではなくアプリケーション権限）
   - `Mail.Send` — メール送信に必要
   - `Calendars.Read` — カレンダー取得に必要
3. 管理者の同意を付与
4. クライアントシークレットを作成

### 環境変数の設定

`.env` ファイルをプロジェクトルートに作成し、以下を記載します:

```env
CLIENT_ID=<Azure ADアプリのクライアントID>
TENANT_ID=<AzureテナントID>
CLIENT_SECRET=<クライアントシークレット>
SENDER_EMAIL=<対象ユーザーのメールアドレス>
```

## 使い方

### メール送信 (`app.py`)

`app.py` の `__main__` ブロックを編集して宛先・件名・本文を指定し、実行します:

```bash
python app.py
```

モジュールとして利用する場合:

```python
from app import send_mail

send_mail(
    to_email="recipient@example.com",
    subject="件名",
    body_text="メール本文",
)
```

### カレンダーイベント取得 (`get_events.py`)

`get_events.py` の `__main__` ブロックで取得期間を指定し、実行します:

```bash
python get_events.py
```

モジュールとして利用する場合:

```python
from get_events import get_events

events = get_events(
    user_email="user@example.com",
    start_datetime="2026-04-01T00:00:00",
    end_datetime="2026-04-30T23:59:59",
)

for event in events:
    print(event["subject"])
```

## 関数仕様

### `app.py`

#### `get_access_token() -> str`

Azure AD からアクセストークンを取得します。

#### `send_mail(to_email, subject, body_text) -> None`

| 引数 | 型 | 説明 |
|---|---|---|
| `to_email` | `str` | 送信先メールアドレス |
| `subject` | `str` | メールの件名 |
| `body_text` | `str` | メールの本文（プレーンテキスト） |

### `get_events.py`

#### `get_access_token() -> str`

Azure AD からアクセストークンを取得します。

#### `get_events(user_email, start_datetime, end_datetime) -> list`

| 引数 | 型 | 説明 |
|---|---|---|
| `user_email` | `str` | カレンダーを取得するユーザーのメールアドレス |
| `start_datetime` | `str` | 取得開始日時（ISO 8601形式 例: `"2026-03-01T00:00:00"`） |
| `end_datetime` | `str` | 取得終了日時（ISO 8601形式 例: `"2026-03-31T23:59:59"`） |

戻り値はイベントオブジェクトのリストです。各イベントには `subject`（件名）、`start`/`end`（開始・終了日時）、`location`（場所）、`organizer`（主催者）が含まれます。

50件を超える場合はページネーションで全件取得します。

## 注意事項

- `.env` ファイルはリポジトリにコミットしないでください
- クライアントシークレットは定期的にローテーションしてください
