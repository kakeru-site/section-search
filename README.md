# セクション検索

部品名からセクション・部品コードを引き、型式から車種名を出す、自分専用の店頭業務ツールです。
GitHub Pages（画面）＋ Firebase（ログインとデータ）で動きます。

> **部品データはこのリポジトリに置かないでください。** データはFirebaseに入れ、ログインした本人だけが読めます。

## できること

| 画面 | 内容 |
|---|---|
| セクション検索 | 部品名の一部・ひらがな・部品コード・セクションで検索。結果から Google / メモ へジャンプ。セクション番号を押すと同じセクションの部品を一覧 |
| 型式 → 車種 | 型式（例：DAA-HFC27）を入れると対応表から車種名を判定。未登録ならGoogleで調べてその場で登録 |
| メモ | 業務中の疑問をメモ。部品・型式に紐付け、未解決／解決済みで管理（自動保存） |
| 検索履歴 | いつ何を調べたか。押すと再検索 |
| データ編集 | 部品・型式の追加／修正／削除、Excel・JSON・CSVの取り込み、バックアップ、設定 |

キーボード：`/` 検索欄へ ／ `↑↓` 選択 ／ `Enter` Google ／ `Ctrl+Enter` メモ ／ `S` `K` `M` `H` `E` 画面切り替え

## セットアップ手順

### 1. Firebase（mame-chishiki プロジェクトに相乗り）

1. **Authentication → ログイン方法** で「メール / パスワード」を有効にする
2. **Authentication → ユーザー → ユーザーを追加** で、自分のメールアドレスとパスワードを登録する
   - 追加したユーザーの **ユーザーUID** をコピーしておく
3. **Realtime Database → データベースを作成**
   - ロケーション：`asia-southeast1（シンガポール）`
   - 「ロックモードで開始」
4. **Realtime Database → ルール** に `database.rules.json` の中身を貼り、`ここに自分のUIDを貼る`（2か所）を手順2のUIDに置き換えて「公開」
5. **プロジェクトの設定 → マイアプリ → ウェブアプリを追加**（名前は「セクション検索」など）
   - 表示された `firebaseConfig` の値を `js/config.js` の `FIREBASE` に貼る
   - `databaseURL` が無い場合は、Realtime Database の画面上部に出ているURLを入れる

### 2. GitHub Pages

1. GitHubで新しいリポジトリを作る（例：`section-search`）
2. このフォルダの中身（`index.html` `css/` `js/` など）をアップロード
3. **Settings → Pages** で「Deploy from a branch」→ `main` / `(root)` → Save
4. 数分後に `https://ユーザー名.github.io/section-search/` で開けるようになる

### 3. Firebaseにドメインを許可

**Authentication → 設定 → 承認済みドメイン** に `ユーザー名.github.io` を追加します。
（これをしないとログインできません）

### 4. データを入れる

1. 公開したページにログイン
2. **データ編集 → 取り込み・書き出し**
   - 部品データ：`parts.json`（Excelから変換済み・頻出★つき）を選ぶ
   - 型式対応表：`kata_seed.json`（下書き。すべて「未確認」で入る）を選ぶ
3. Excel（部品マスタシート）を直接取り込むこともできます。ただし★頻出は付きません。

### 5. 生体認証ロック（任意）

**データ編集 → 設定 → 生体認証ロック → この端末に設定する**
iPhoneならFace ID、WindowsならWindows Hello（顔・指紋・PIN）で解除できます。端末ごとに設定してください。
※ 端末のロック解除用です。初回ログイン（メール＋パスワード）の代わりではありません。

### 6. Google検索の埋め込み（任意）

型式が未登録のとき、Google検索の結果をアプリ内に表示できます。

1. https://programmablesearchengine.google.com/ で検索エンジンを作成（「ウェブ全体を検索」が選べればオン。選べない場合は、よく見るサイトを登録）
2. 表示される **検索エンジンID**（cx）をコピー
3. **データ編集 → 設定 → Google検索の埋め込み** に貼って保存

設定しない場合は、Googleが新しいタブで開きます。

## デモモード

`js/config.js` の `apiKey` が空のままだと「デモモード」で動きます（ログイン不要・このブラウザにだけ保存）。
見た目や操作を試すときに使えます。

## ファイル構成

```
index.html
css/style.css
js/config.js     ← Firebaseの設定を貼るところ
js/app.js        画面と操作
js/search.js     あいまい検索・型式判定
js/store.js      Firebase / デモの読み書き
js/faceid.js     生体認証ロック
database.rules.json  Realtime Database のルール（Firebaseコンソールに貼る）
```

## データの置き場所（Realtime Database）

```
sectionApp/
  parts/{id}      sec, code, name, note, hot, updatedAt
  kata/{型式}     name, verified, note, updatedAt
  memos/{id}      title, body, done, partId, partLabel, kataCode, kataName, createdAt, updatedAt
  history/{id}    type(part|kata), q, hits, at
  settings/       googlePrefix, cseId, synonyms
  stats/sections/{sec}  sec, n
```
