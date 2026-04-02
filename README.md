# WINUSET 月次統計自動記入ツール 説明書

## 概要

毎月の定型作業として、メール取得APIおよびCSVファイルから統計データを取得し、
Excelファイルへ自動的に記入するツールです。
Pythonは不要で、Dockerコンテナ上で動作します。

---

## ディレクトリ構成

```
WINUSET/
├── dockerfile          # Dockerイメージのビルド定義
├── main.py             # メインスクリプト
├── source/             # 入力CSVファイルを置くフォルダ
├── report/             # 出力対象のExcelファイルが入るフォルダ
└── credentials/
    └── config.json     # メール取得APIのシークレットキー
```

---

## 初回セットアップ

### 1. Dockerイメージのビルド（初回のみ）

```bash
docker build -t winuset .
```

### 2. credentials/config.json の作成（初回のみ）

```json
{
  "api_key": "GASに設定したシークレットキー"
}
```

---

## 毎月の作業手順

### 1. CSVファイルを source/ に配置する

前月分（例：2026年2月 → `202602`）の以下のCSVを配置します。

| CSVファイル名 | 対応するExcel |
|---|---|
| 集計_winlogin_user_YYYYMM.csv | 07_Windows利用統計.xlsx |
| 集計_unuseOneDrive_user_YYYYMM.csv | OneDriveログインしない.xlsx |
| 集計_useOneDrive_user_YYYYMM.csv | OneDriveログインする.xlsx |
| 集計_dualOneDrive_user_YYYYMM.csv | OneDriveログイン両方.xlsx |

### 2. report/ 内のExcelファイルをすべて閉じる

Excelを開いたままにするとファイル保存エラーが発生します。

### 3. スクリプトを実行する

```bash
docker run --rm `
  -v "C:/Users/kshinjo/work/WINUSET/source:/source" `
  -v "C:/Users/kshinjo/work/WINUSET/report:/report" `
  -v "C:/Users/kshinjo/work/WINUSET/credentials:/credentials" `
  winuset
```

---

## 実行ログの見方

正常に完了した場合、以下のようなログが表示されます。

```
対象年月: 202602  書き込み先シート: [02]

APIからWindowsクライアント利用回数を取得中...
  -> 取得値: 729  書き込み先: Winログイン!AF13
  -> 完了

APIからWindowsアプリケーション利用回数を取得中...
  -> SPSS(ccmaster): 1
  -> 完了: 30 行を書き込み（0 行はスキップ）  列: 13 (2月)

APIからOneDriveログイン状況を取得中...
  -> 必ず使用:4  両方:65  使用しない:121  行:12
  -> 完了

処理中: 集計_winlogin_user_202602.csv
  -> 07_Windows利用統計.xlsx [sheet: 02]
  -> 完了: 54 行を書き込みました
...
```

- **対象年月**: 実行日から自動計算した前月（例：3月16日実行 → 202602）
- **書き込み先シート**: 月番号2桁（例：[02] = 2月シート）
- **AF列の行番号**: 4月始まり年度順（4月=3行目, 2月=13行目, 3月=14行目）

### エラーメッセージ

| メッセージ | 原因 | 対処 |
|---|---|---|
| `[SKIP] CSV が見つかりません` | source/にCSVがない | CSVファイルを配置する |
| `[SKIP] Excel が見つかりません` | report/にExcelがない | Excelファイルを確認する |
| `[WARN] シートが存在しません` | 月番号のシートがない | Excelのシート名を確認する |
| `[ERROR] API取得失敗: 401` | APIキーが不正 | config.jsonのキーを確認する |
| `PermissionError` | Excelが開いている | Excelをすべて閉じてから再実行 |

---

## メール取得API連携

Google Apps Script（GAS）で作成されたメール取得APIを使用しています。

### APIエンドポイント

```
https://script.google.com/macros/s/AKfycby.../exec?type=<種別>&key=<シークレットキー>
```

### typeパラメータと取得データ

| type | メール件名 | 取得する値 | 書き込み先 |
|---|---|---|---|
| `winclient` | Windowsクライアント月次統計 | `■Windowsクライアント利用回数:` の数値 | 07_Windows利用統計.xlsx「Winログイン」シート AF列 |
| `winapp` | Windowsアプリケーション月次統計 | アプリ別利用回数（30アプリ） | 07_Windows利用統計.xlsx「Winアプリ」シート 各月列 |
| `spss` | SPSS使用者数月次レポート | `ccmasterドメイン（教卓PC，共用PC等）：` の数値 | 07_Windows利用統計.xlsx「Winアプリ」シート SPSS行 |
| `onedrive` | OneDrive 月次ログインレポート | 必ず使用 / 使用するときもある / 使用しない | 月次OneDriveログイン状況.xlsx「OneDriveログイン」シート B/C/D列 |

### セキュリティ

- APIへのアクセスはすべてシークレットキーによる認証を使用
- キーは `credentials/config.json` に保存（Gitには含まれない）

---

## Excel書き込みルール

### 月シートへのCSV貼り付け（4ファイル）

各Excelの月番号シート（例：`02`）に、対応するCSVの内容をそのまま上書きします。
書式（色・罫線等）は保持され、値のみ更新されます。

### 4月始まり年度の列・行の対応

Winアプリシート（列）およびWinログインシート（AF列の行）は、
4月始まりの日本の年度に沿って以下のように対応しています。

| 月 | インデックス |
|---|---|
| 4月 | 3 |
| 5月 | 4 |
| ... | ... |
| 12月 | 11 |
| 1月 | 12 |
| 2月 | 13 |
| 3月 | 14 |

---

## Dockerfileについて

```dockerfile
FROM python:3.12-slim
WORKDIR /app
RUN pip install --no-cache-dir openpyxl requests
COPY main.py auth.py ./
CMD ["python", "main.py"]
```

- ベースイメージ: Python 3.12（軽量版）
- 使用ライブラリ:
  - `openpyxl`: Excelファイルの読み書き
  - `requests`: メール取得APIへのHTTPリクエスト
- イメージの再ビルドが必要なのは `main.py` や `dockerfile` を変更した場合のみ
