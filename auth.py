"""
一回限りの OAuth2 認証セットアップスクリプト

使い方:
  docker run --rm -p 8080:8080 \
    -v "C:/Users/kshinjo/work/WINUSET/credentials:/credentials" \
    winuset python auth.py

  1. 表示された URL をブラウザで開き、oikaze.jp アカウントでログイン
  2. 認証が完了すると自動的に token.json が保存される
"""

import json
import os

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    'openid',
    'https://www.googleapis.com/auth/userinfo.email',
]

CLIENT_SECRETS_PATH = '/credentials/client_secrets.json'
TOKEN_PATH = '/credentials/token.json'


def main():
    if not os.path.exists(CLIENT_SECRETS_PATH):
        print(f'エラー: {CLIENT_SECRETS_PATH} が見つかりません。')
        print()
        print('以下の手順で client_secrets.json を用意してください:')
        print('  1. https://console.cloud.google.com/ を開く')
        print('  2. 「APIとサービス」→「認証情報」→「認証情報を作成」')
        print('  3. 「OAuthクライアントID」→ アプリの種類:「デスクトップアプリ」')
        print('  4. JSONをダウンロードして credentials/client_secrets.json として保存')
        return

    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_PATH, SCOPES)

    import builtins
    import re as _re

    url_file = '/credentials/auth_url.txt'
    _original_print = builtins.print

    def _capture_print(*args, **kwargs):
        text = ' '.join(str(a) for a in args)
        urls = _re.findall(r'https://accounts\.google\.com\S+', text)
        if urls:
            with open(url_file, 'w') as f:
                f.write(urls[0] + '\n')
            import sys
            sys.stderr.write(
                f'\n認証URLを {url_file} に保存しました。\n'
                'credentials/auth_url.txt をメモ帳で開き、URLをブラウザにコピーしてください。\n\n'
            )
            sys.stderr.flush()
        _original_print(*args, **kwargs)

    print('ローカルサーバーを起動中 (port 8080)...')

    builtins.print = _capture_print
    try:
        creds = flow.run_local_server(port=8080, open_browser=False, bind_addr='0.0.0.0')
    finally:
        builtins.print = _original_print

    token_data = {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': list(creds.scopes),
        'id_token': getattr(creds, 'id_token', None),
    }

    with open(TOKEN_PATH, 'w') as f:
        json.dump(token_data, f, indent=2)

    print()
    print(f'認証情報を保存しました: {TOKEN_PATH}')
    print('次回からは main.py が自動的にトークンを使用します。')


if __name__ == '__main__':
    main()
