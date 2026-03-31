# Microsoft Graph API メール送信ツール

Microsoft Graph API を使用して、Azure AD アプリケーションからメールを送信する Python スクリプトです。

## 概要

OAuth2 クライアント資格情報フローで認証を行い、指定したユーザーのメールボックスからメールを送信します。

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
2. `Mail.Send` のアプリケーション権限を付与（委任ではなくアプリケーション権限）
3. 管理者の同意を付与
4. クライアントシークレットを作成

### 環境変数の設定

`.env` ファイルをプロジェクトルートに作成し、以下を記載します:

```env
CLIENT_ID=<Azure ADアプリのクライアントID>
TENANT_ID=<AzureテナントID>
CLIENT_SECRET=<クライアントシークレット>
SENDER_EMAIL=<送信元メールアドレス>
```

## 使い方

### スクリプトとして直接実行

`app.py` の `__main__` ブロックを編集して宛先・件名・本文を指定し、実行します:

```bash
python app.py
```

### モジュールとして利用

```python
from app import send_mail

send_mail(
    to_email="recipient@example.com",
    subject="件名",
    body_text="メール本文",
)
```

## 関数仕様

### `get_access_token() -> str`

Azure AD からアクセストークンを取得します。

### `send_mail(to_email, subject, body_text) -> None`

| 引数 | 型 | 説明 |
|---|---|---|
| `to_email` | `str` | 送信先メールアドレス |
| `subject` | `str` | メールの件名 |
| `body_text` | `str` | メールの本文（プレーンテキスト） |

## 注意事項

- `.env` ファイルはリポジトリにコミットしないでください
- クライアントシークレットは定期的にローテーションしてください
