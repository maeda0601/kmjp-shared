# Microsoft Graph API ツール集

Microsoft Graph API を使用して、Azure AD アプリケーションからメール送信・カレンダー操作を行う Python スクリプト／MCPサーバーです。

## 概要

OAuth2 クライアント資格情報フロー（アプリ認証）で認証を行い、以下の操作が可能です。

| 機能 | スクリプト | MCPツール名 |
|---|---|---|
| メール送信 | `app.py` | `send_mail` |
| カレンダー予定取得 | `get_events.py` | `get_events` |
| カレンダー予定追加 | `create_event.py` | `create_event` |

---

## MCPサーバーとして使う（推奨）

Claude Desktop から自然言語でカレンダー操作・メール送信ができます。

### 必要要件

- [uv](https://docs.astral.sh/uv/) がインストールされていること

```bash
# Windows
winget install --id=astral-sh.uv -e

# Mac/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### セットアップ

#### 1. リポジトリをクローン

```bash
git clone <リポジトリURL>
cd kmjp-shared
```

#### 2. `.env` ファイルを作成

```env
CLIENT_ID=<Azure ADアプリのクライアントID>
TENANT_ID=<AzureテナントID>
CLIENT_SECRET=<クライアントシークレット>
SENDER_EMAIL=<自分のメールアドレス>
```

> `CLIENT_ID` / `TENANT_ID` / `CLIENT_SECRET` は管理者から共有してもらいます。  
> `SENDER_EMAIL` は自分のメールアドレスを設定します。

#### 3. Claude Desktop に登録

`%APPDATA%\Claude\claude_desktop_config.json` の `mcpServers` に以下を追加します：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": [
        "--directory",
        "<クローンしたパス>",
        "run",
        "mcp_server.py"
      ]
    }
  }
}
```

> Windows で `uv` がフルパス指定の場合は `"command"` を `"C:\\Users\\<ユーザー名>\\.local\\bin\\uv.exe"` に変更してください。

#### 4. Claude Desktop を再起動

登録後に Claude Desktop を再起動すると `create_event`・`get_events`・`send_mail` ツールが使えるようになります。

### 動作確認

```bash
uv run mcp_server.py
```

---

## スクリプトとして使う

### 必要要件

- Python 3.12 以上

```bash
pip install requests python-dotenv
```

`.env` ファイルの作成は上記と同様です。

### メール送信 (`app.py`)

```bash
python app.py
```

```python
from app import send_mail

send_mail(
    to_email="recipient@example.com",
    subject="件名",
    body_text="メール本文",
)
```

### カレンダー予定取得 (`get_events.py`)

```bash
python get_events.py
```

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

### カレンダー予定追加 (`create_event.py`)

```bash
python create_event.py
```

```python
from create_event import create_event

event = create_event(
    user_email="user@example.com",
    subject="打合せ",
    start_datetime="2026-04-01T10:00:00",
    end_datetime="2026-04-01T11:00:00",
    timezone="Asia/Tokyo",
    location="会議室A",
    body_text="議題: プロジェクト進捗確認",
    attendees=["attendee@example.com"],
)

print(event["id"])
```

---

## Azure AD アプリの準備（管理者作業）

1. Azure ポータルでアプリを登録
2. 以下のアプリケーション権限を付与（委任ではなくアプリケーション権限）
   - `Mail.Send` — メール送信に必要
   - `Calendars.ReadWrite` — カレンダー取得・予定追加に必要
3. 管理者の同意を付与
4. クライアントシークレットを作成し、メンバーに共有

---

## 注意事項

- `.env` ファイルはリポジトリにコミットしないでください
- クライアントシークレットは定期的にローテーションしてください
