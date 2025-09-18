# Google Apps Script 実験スケジューラー

Google Apps Scriptを使用した実験参加者向けスケジュール管理システムです。

## 📋 はじめに

このシステムは大学の研究実験における参加者のスケジュール管理を自動化するために作られました。
プログラミング初心者でも安心して使えるよう、導入から利用まで丁寧に説明します。

### できること
- 実験の日時枠を自動で生成
- 参加者の申し込みを受け付け
- 最小人数に達した枠を自動で確定
- 確定・リマインドメールの自動送信
- 管理者向けの状況レポート配信

## 🚀 完全ガイド（初心者向け）

### 📝 事前準備

このシステムを使うために以下のアカウントが必要です：
- **Googleアカウント** - Gmail、Google Drive、Google Apps Scriptを使うため
- **GitHubアカウント** - コードをダウンロード・管理するため（無料で作成可能）

---

## 🖥️ STEP 1: 開発環境の準備

### 1-1. コマンドプロンプト（ターミナル）を開く

**Windows の場合:**
1. `Windows + R` キーを押す
2. `cmd` と入力してEnterキーを押す
3. 黒い画面（コマンドプロンプト）が開きます

**Mac の場合:**
1. `Command + Space` キーを押す
2. `ターミナル` と入力してEnterキーを押す

### 1-2. Node.js のインストール

Node.jsは、JavaScriptをコンピューター上で実行するための環境です。

1. [Node.js公式サイト](https://nodejs.org/) にアクセス
2. **LTS版**（推奨版）をダウンロード
3. ダウンロードしたファイルを実行してインストール
4. インストール確認のため、コマンドプロンプトで以下を実行：

```bash
node --version
```

バージョン番号（例：`v18.17.0`）が表示されればOK！

### 1-3. clasp のインストール

claspは、Google Apps Scriptをコマンドラインで操作するためのツールです。

コマンドプロンプトで以下のコマンドを実行：

```bash
npm install -g @google/clasp
```

インストール確認：

```bash
clasp --version
```

### 1-4. clasp でGoogleアカウントにログイン

```bash
clasp login
```

ブラウザが開いて、Googleアカウントでログインを求められます。
ログインが完了すると「Login successful!」と表示されます。

---

## 📦 STEP 2: プロジェクトの取得

### 2-1. GitHubからプロジェクトをダウンロード

1. コマンドプロンプトで、プロジェクトを置きたい場所に移動：

```bash
# デスクトップに移動する例
cd Desktop
```

2. GitHubからプロジェクトをクローン：

```bash
git clone https://github.com/YukiInoueNakata/gs-experiment-scheduler.git
```

3. プロジェクトフォルダに移動：

```bash
cd gs-experiment-scheduler
```

### 2-2. プロジェクト内容の確認

フォルダ内のファイルを確認：

```bash
# Windows の場合
dir

# Mac の場合
ls -la
```

以下のようなファイルが表示されるはずです：
- `Code.js`, `Setting.js`, `Mail.js` など - プログラム本体
- `.env` - 環境設定のテンプレート
- `README.md` - このファイル

---

## 🔧 STEP 3: Google Apps Script の設定

### 3-1. 新しいGoogle Apps Scriptプロジェクトを作成

1. コマンドプロンプトで以下を実行（プロジェクトフォルダ内で）：

```bash
clasp create --title "実験スケジューラー" --type webapp
```

成功すると `Created new script: https://script.google.com/d/xxxxx/edit` のようなメッセージが表示されます。

### 3-2. ローカルのコードをGoogle Apps Scriptにアップロード

```bash
clasp push
```

「Manifest file has been updated. Do you want to push and overwrite?」と聞かれたら `y` を入力してEnterを押します。

### 3-3. スプレッドシートの作成

1. [Google スプレッドシート](https://docs.google.com/spreadsheets/)にアクセス
2. 「空白のスプレッドシート」を作成
3. 適当な名前を付ける（例：「実験スケジューラー_データ」）
4. URLから**スプレッドシートID**をコピー：

```
https://docs.google.com/spreadsheets/d/【ここがスプレッドシートID】/edit
```

例：`1AbC2DeF3GhI4JkL5MnO6PqR7StU8VwX9YzA` の部分

### 3-4. 環境設定（スクリプト プロパティ）

1. 以下のコマンドでGoogle Apps Scriptエディタを開く：

```bash
clasp open
```

2. 左メニューの「設定」→「スクリプト プロパティ」をクリック
3. 「スクリプト プロパティを追加」で以下を追加：

| プロパティ名 | 値（あなたの情報に変更） |
|-------------|---------------------|
| `SS_ID` | コピーしたスプレッドシートID |
| `ADMIN_EMAILS` | あなたのメールアドレス |
| `MAIL_FROM_NAME` | `実験担当（自動送信）` |
| `LOCATION` | `○○大学 ○号館 ○F 実験室` |

### 3-5. 初期セットアップの実行

Google Apps Scriptエディタで：

1. `Code.js`を開く
2. 関数選択で`setup`を選択
3. 「実行」ボタンをクリック
4. 権限を求められたら「権限を確認」→「許可」をクリック

成功すると、スプレッドシートに必要なシートが自動作成されます。

### 3-6. Webアプリとしてデプロイ

1. Google Apps Scriptエディタで「デプロイ」→「新しいデプロイ」
2. 種類の選択で「ウェブアプリ」を選択
3. 設定：
   - 説明：任意（例：`実験スケジューラー v1.0`）
   - 次のユーザーとして実行：「自分」
   - アクセスできるユーザー：「全員」
4. 「デプロイ」をクリック

**重要**: デプロイ後に表示される**ウェブアプリURL**をコピーして保存してください。これが参加者に公開するURLです。

---

## 📊 STEP 4: 設定のカスタマイズ

### 4-1. 基本設定の変更

Google Apps Scriptエディタで`Setting.js`を開き、必要に応じて以下を変更：

```javascript
const CONFIG = {
  title: '実験参加スケジュール',        // Webページのタイトル

  // 枠生成の設定
  startDate: '2025-09-01',            // 実験開始日
  endDate:   '2025-09-30',            // 実験終了日
  timeWindows: [                      // 1日の時間枠
    '11:00-12:00',
    '13:20-14:20',
    '15:00-16:00',
    '16:50-17:50'
  ],

  // 人数設定
  capacity: 2,                        // 1枠の最大人数
  minCapacityToConfirm: 2,            // 確定に必要な最小人数

  // 除外設定
  excludeWeekends: true,              // 土日を除外
  excludeDates: ['2025-09-16','2025-09-23'], // 特定日を除外
};
```

### 4-2. 設定変更後の反映

設定を変更したら：

1. ファイルを保存（`Ctrl + S`）
2. コマンドプロンプトで：

```bash
clasp push
```

---

## 🎯 STEP 5: 運用開始

### 5-1. 参加者への案内

デプロイで取得したWebアプリURLを参加者に案内します：

```
実験参加申し込みURL: https://script.google.com/macros/s/xxxxx/exec

【申し込み方法】
1. 上記URLにアクセス
2. 名前とメールアドレスを入力
3. 参加希望の日時を選択
4. 「申し込み」ボタンをクリック

※確定の可否はメールでお知らせします
```

### 5-2. 管理者の日常業務

- **状況確認**: スプレッドシートで申し込み状況を確認
- **メール送信**: 自動で確定メール・リマインドメールが送信されます
- **緊急時**: `Code.js`の関数を手動実行で対応可能

---

## 🔧 トラブルシューティング

### よくあるエラーと解決方法

#### ❌ 「clasp コマンドが見つかりません」
**原因**: Node.jsまたはclaspがインストールされていない
**解決方法**:
```bash
# Node.jsのバージョン確認
node --version

# claspの再インストール
npm install -g @google/clasp
```

#### ❌ 「権限がありません」エラー
**原因**: Google Apps Scriptの権限設定
**解決方法**:
1. Google Apps Scriptエディタで関数を実行
2. 「権限を確認」→「詳細」→「安全ではないページに移動」
3. すべての権限を許可

#### ❌ 「スプレッドシートが見つかりません」
**原因**: スプレッドシートIDが間違っている
**解決方法**:
1. スプレッドシートのURLから正しいIDをコピー
2. スクリプト プロパティの`SS_ID`を修正

#### ❌ メールが送信されない
**原因**: Gmail APIの制限またはメールアドレスの間違い
**解決方法**:
1. 送信者のメールアドレスがGoogleアカウントと一致しているか確認
2. 1日のメール送信制限を確認（通常100通まで）

#### ❌ Webアプリが表示されない
**原因**: デプロイ設定の問題
**解決方法**:
1. 「アクセスできるユーザー」が「全員」になっているか確認
2. 新しいデプロイを作成してURLを更新

### サポートが必要な場合

1. **エラーメッセージをコピー**して保存
2. **何をしていた時に発生したか**を記録
3. GoogleやStack Overflowで検索

---

## 📚 参考資料

- [Google Apps Script公式ドキュメント](https://developers.google.com/apps-script)
- [clasp公式ドキュメント](https://github.com/google/clasp)
- [Node.js公式サイト](https://nodejs.org/)

---

## 📂 ファイル構成

```
gs-experiment-scheduler/
├── Code.js              # メイン処理
├── Setting.js           # 設定ファイル（環境変数化済み）
├── Mail.js              # メール送信処理
├── Utils.js             # ユーティリティ関数
├── Templetes.js         # メールテンプレート
├── Menu.js              # 管理メニュー
├── BatchProcess.js      # バッチ処理
├── AddSlots.js          # 枠追加処理
├── Cancel.js            # キャンセル処理
├── DataCleanup.js       # データクリーンアップ
├── Index.html           # Webアプリのフロントエンド
├── appsscript.json      # Apps Script設定
├── .clasp.json          # clasp設定
├── .env                 # 環境変数テンプレート
├── .gitignore           # Git除外設定
└── README.md            # このファイル
```

---

## 🔐 セキュリティ注意事項

- **個人情報の取り扱い**: 参加者の個人情報は適切に管理してください
- **スプレッドシートの共有**: 必要最小限の人数にのみ共有
- **メール送信**: テスト送信で動作確認してから本運用
- **バックアップ**: 重要なデータは定期的にバックアップ

このシステムは教育・研究目的で作成されています。商用利用や大規模運用の際は、追加のセキュリティ対策を検討してください。