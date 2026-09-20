# GAS（会員証のバックエンド）の触り方

スプレッドシート「LINE会員情報」に紐づく Apps Script プロジェクトのソースです。

| ファイル | 中身 |
|---|---|
| **`Code.gs`** | **Apps Script に貼り付けるのはこれ** |
| `Code.test.js` | Node.js 用のテスト。**Apps Script に貼らないこと** |

`Code.test.js` を貼ると `ReferenceError: require is not defined` になります。

---

## 更新のしかた

### 1. コードを貼り替える

script.google.com でプロジェクトを開き、`コード.gs` を **Ctrl+A → 貼り付け → Ctrl+S**。

⚠️ **ブラウザで長いコードを選択すると途中で切れることがあります。** 切れると
`SyntaxError: Unexpected end of input` になります。貼ったあと一番下までスクロールし、
最後が `function authorizeAndTestLine()` で終わっていることを確認してください。

切れる場合はファイルに落としてから開くほうが確実です。

```bash
curl -s https://raw.githubusercontent.com/nanyakanya-ops/nanyakanya-card/main/gas/Code.gs -o /c/Users/hormo/Desktop/Code.gs
wc -l < /c/Users/hormo/Desktop/Code.gs
notepad /c/Users/hormo/Desktop/Code.gs
```

### 2. 外部通信の承認を通す（初回・権限が増えたとき）

関数選択で **`authorizeAndTestLine`** を選び **▷ 実行**。

- 「承認が必要です」→ 権限を確認 → アカウント選択
- 「Googleで確認されていません」→ **詳細** → **「（安全ではないページ）に移動」** → **許可**
- 実行ログに **`HTTP 400`** と出れば成功（`dummy` という偽トークンを送っているので
  400 が返るのが正解。LINEまで通信が届いた証拠）

⚠️ **`doGet` や `doPost` を実行しても承認画面は出ません。** これらは外部通信をしないため、
`UrlFetchApp` の権限（`script.external_request`）が要求されないからです。この状態で公開すると
実機で「**UrlFetchApp.fetch を呼び出す権限がありません**」になります。2026-09-21 にこれで詰まりました。

### 3. デプロイ

**デプロイ → デプロイを管理 → 鉛筆（編集）→ バージョン「新バージョン」→ デプロイ**

⚠️ **「新しいデプロイ」を作らないこと。** URLが変わり、`index.html` の `GAS_URL` も
直すことになります。「デプロイを管理 → 編集」ならURLは変わりません。

「アクセスできるユーザー」は **「全員」**。「Googleアカウントを持つ全員」では、
LINEの内蔵ブラウザ（Googleに未ログイン）から弾かれます。

---

## 動作確認

```bash
curl -s -m 30 -L "https://script.google.com/macros/s/＜デプロイID＞/exec" \
  --data '{"action":"check","idToken":"dummy"}'; echo
```

期待する結果:

```json
{"status":"unauthorized","message":"ログイン情報を確認できませんでした。…","detail":"LINEが拒否 HTTP400 …"}
```

偽トークンなので `unauthorized` が正解です。**JSONが返ること**が要点。

⚠️ **`-X POST` を付けないこと。** `-L` と併用すると、curl がリダイレクト先へ本文を送らず
`411 Length Required` になります。`--data` だけで POST になります。

HTMLが返る場合は、GASが実行されていません（承認未完了、アクセス権、貼り付けミスのいずれか）。

---

## テスト

```bash
node gas/Code.test.js
```

Apps Script の実行環境（`SpreadsheetApp` / `UrlFetchApp` / `LockService` など）を
偽物に差し替えて、**本物のスプレッドシートに触らずに**ロジックだけを動かします。
追加パッケージは不要です。

認証の偽装・名寄せの乗っ取り・会員番号の採番・二重登録・入力チェックなどを見ています。
`Code.gs` を変えたら必ず走らせてください。

---

## 困ったときは

| 症状 | 原因 |
|---|---|
| 実機で「サーバーに接続できませんでした」 | GASが実行されていない。承認・アクセス権・貼り付けミスを確認 |
| 「UrlFetchApp.fetch を呼び出す権限がありません」 | 上の手順2（`authorizeAndTestLine`）が未実施 |
| `ReferenceError: require is not defined` | `Code.test.js` を貼ってしまっている |
| `SyntaxError: Unexpected end of input` | 貼り付けが途中で切れている |
| `411 Length Required` | curl に `-X POST` を付けている |

**元に戻したいとき**は、デプロイを管理 → 鉛筆 → バージョンを1つ前に → デプロイ。
これで即座に以前の動作に戻ります。

---

## 設計のメモ

- **クライアントが送る userId は信用しない。** LINEのIDトークンを `api.line.me` に照会し、
  返ってきた `sub` だけを会員のLINE IDとして使う。以前は `?action=check&uid=...` を
  叩けば誰でも他人の会員番号を引けた
- **名寄せはA列（LINE ID）が空の行だけを対象にする。** 以前は氏名が一致しただけで
  A列を上書きしており、同姓同名の別人が登録すると先に登録していた人のLINE IDが消え、
  別人がその会員番号を受け取っていた
- 電話番号の一致を氏名より先に探す（同姓同名より確実なため）
- 電話番号の突合は全角数字を半角に直してから。シートには全角で入っている
- 同時登録で同じ会員番号を二人に割り当てないよう `LockService` で直列化
