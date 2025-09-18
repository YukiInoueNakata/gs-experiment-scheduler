# Google Apps Script 実験スケジューラー

Google Apps Scriptを使用した実験参加者向けスケジュール管理システムです。

## セットアップ手順

### 1. 環境設定

個人情報はGoogle Apps Scriptのスクリプトプロパティで管理します：

1. Google Apps Script エディタを開く
2. 左メニューの「設定」→「スクリプト プロパティ」をクリック
3. 以下のプロパティを追加：

| プロパティ名 | 説明 | 例 |
|-------------|------|-----|
| `SS_ID` | スプレッドシートID（必須） | `1AbC2DeF3GhI4JkL5MnO6PqR7StU8VwX9YzA` |
| `ADMIN_EMAILS` | 管理者メールアドレス（カンマ区切り） | `admin@example.com,manager@example.com` |
| `MAIL_FROM_NAME` | メール送信者名 | `実験担当（自動送信）` |
| `LOCATION` | 実験場所 | `立命館大学 OIC C号館 3F 実験室A` |

### 2. スプレッドシートIDの取得方法

1. Google スプレッドシートを開く
2. URLから以下の部分をコピー：
   ```
   https://docs.google.com/spreadsheets/d/【ここがスプレッドシートID】/edit
   ```

### 3. GitHub投稿時の注意

- `.env`ファイルには実際の値は入力しないでください
- `.gitignore`により環境設定ファイルはGitの管理対象外になります
- Setting.jsから個人情報を削除し、環境変数経由で読み込むよう修正済みです

## ファイル構成

- `Setting.js` - 設定ファイル（個人情報を環境変数化済み）
- `Code.js` - メイン処理
- `Mail.js` - メール送信処理
- その他の`.js`ファイル - 各種機能モジュール
- `.env` - 環境設定のテンプレート（実際の値は記入しない）
- `.gitignore` - Git管理除外設定

## 使用方法

1. 上記の環境設定を完了
2. `setup()`関数を実行してシートとトリガーを初期化
3. Webアプリとしてデプロイ

このREADMEファイルはGitHubへの投稿準備として作成されました。個人情報は含まれていません。