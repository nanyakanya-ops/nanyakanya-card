/**
 * なんやかん屋 LINE会員証 バックエンド
 *
 * 【重要】クライアントから送られてくる userId は信用しない。
 * LINEが発行したIDトークンをLINE本体に照会し、その結果に入っている sub だけを
 * 会員のLINE IDとして扱う。これによりURLを叩くだけで他人の会員番号を引ける状態を防ぐ。
 */

const SHEET_ID = '1hKTDLuxi-NEh4pKrtUwEmHXyHXZasBrOe8wzfjjYc_Q';
const SHEET_NAME = '会員';

// LIFF ID「2009831421-Gb85ZD1F」の「-」より前がチャネルID
const LINE_CHANNEL_ID = '2009831421';

// 列の位置（0始まり）
const COL_UID = 0;        // A: LINE ID
const COL_MEMBER_ID = 1;  // B: 会員番号
const COL_NAME = 2;       // C: 名前
const COL_KANA = 3;       // D: カナ
const COL_ADDRESS = 4;    // E: 住所
const COL_PHONE = 5;      // F: 電話番号
const COL_EMAIL = 6;      // G: メール
const COL_REGISTERED = 7; // H: 登録日時


function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 直前の検証がなぜ失敗したかを覚えておく。原因の切り分けにだけ使う。
// IDトークン本体は絶対に入れない
var lastVerifyError = '';

/**
 * IDトークンをLINEに照会し、確認できたLINE ID(sub)を返す。確認できなければ null。
 */
function verifyIdToken_(idToken) {
  lastVerifyError = '';

  if (!idToken) {
    lastVerifyError = 'トークンが送られてきていない';
    return null;
  }

  // payload をオブジェクトで渡すと、contentType を明示したときの変換のされ方が
  // はっきりしない。フォーム形式の文字列を自分で組み立てて曖昧さをなくす
  var body = 'id_token=' + encodeURIComponent(idToken) +
             '&client_id=' + encodeURIComponent(LINE_CHANNEL_ID);

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: body,
      muteHttpExceptions: true
    });
  } catch (err) {
    lastVerifyError = 'LINEに接続できない: ' + err;
    console.log(lastVerifyError);
    return null;   // LINEに繋がらないときは「確認できなかった」として扱う
  }

  // 失敗した理由は実行ログに残す（トークン本体は出さない）。
  // Apps Script の「実行数」から確認できる
  if (res.getResponseCode() !== 200) {
    lastVerifyError = 'LINEが拒否 HTTP' + res.getResponseCode() + ' ' +
                      res.getContentText().slice(0, 200);
    console.log(lastVerifyError);
    return null;
  }

  var payload;
  try {
    payload = JSON.parse(res.getContentText());
  } catch (err) {
    lastVerifyError = 'LINEの応答を解釈できない';
    console.log(lastVerifyError);
    return null;
  }

  if (!payload.sub) {
    lastVerifyError = 'LINEの応答に sub が無い';
    console.log(lastVerifyError);
    return null;
  }

  // aud（このトークンの宛先）が自分のチャネルであることも念のため確かめる
  if (String(payload.aud) !== LINE_CHANNEL_ID) {
    lastVerifyError = 'チャネルID不一致 期待=' + LINE_CHANNEL_ID + ' 実際=' + payload.aud;
    console.log(lastVerifyError);
    return null;
  }

  return String(payload.sub);
}

/** 全角数字を半角に直したうえで数字だけ取り出す（シートには全角で入っている） */
function normalizePhone_(v) {
  var s = String(v == null ? '' : v).replace(/[０-９]/g, function (m) {
    return String.fromCharCode(m.charCodeAt(0) - 0xFEE0);
  });
  return s.replace(/[^0-9]/g, '');
}

/** 空白（半角・全角）を取り除く */
function normalizeName_(v) {
  return String(v == null ? '' : v).replace(/[\s　]/g, '');
}

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

/** すでにこのLINE IDで登録済みの行番号（0始まりの配列添字）。無ければ -1 */
function findRowByUid_(data, uid) {
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_UID]) === uid) return i;
  }
  return -1;
}

/**
 * 紙の会員など、まだLINEと紐付いていない行を探して名寄せする。
 *
 * A列（LINE ID）が既に入っている行は【対象にしない】。
 * 同姓同名の別人が登録したときに、先に登録していた人のLINE IDを上書きして
 * 会員記録を乗っ取ってしまうため。
 *
 * 電話番号の一致を先に探す。同姓同名より確実なので、氏名の偶然の一致に負けないようにする。
 */
function findMergeRow_(data, phone, name) {
  var i;
  for (i = 1; i < data.length; i++) {
    if (String(data[i][COL_UID]).trim() !== '') continue;
    var rowPhone = normalizePhone_(data[i][COL_PHONE]);
    if (phone !== '' && rowPhone !== '' && rowPhone === phone) return i;
  }
  for (i = 1; i < data.length; i++) {
    if (String(data[i][COL_UID]).trim() !== '') continue;
    var rowName = normalizeName_(data[i][COL_NAME]);
    if (name !== '' && rowName !== '' && rowName === name) return i;
  }
  return -1;
}

/** 使われていない一番小さい会員番号 */
function nextMemberId_(data) {
  var used = {};
  for (var i = 1; i < data.length; i++) {
    var raw = String(data[i][COL_MEMBER_ID]).trim();
    if (raw === '') continue;
    var n = Number(raw);
    if (!isNaN(n) && n > 0) used[n] = true;
  }
  var id = 1;
  while (used[id]) id++;
  return String(id);
}

/** 登録内容の確認。問題があれば日本語のメッセージ、無ければ null */
function validateRegistration_(params) {
  var labels = [[ 'name', '氏名' ], [ 'kana', 'フリガナ' ], [ 'address', '住所' ], [ 'phone', '電話番号' ]];
  for (var i = 0; i < labels.length; i++) {
    if (normalizeName_(params[labels[i][0]]) === '') {
      return labels[i][1] + 'が入力されていません。';
    }
  }
  var digits = normalizePhone_(params.phone);
  if (digits.length < 10 || digits.length > 11) {
    return '電話番号は数字10〜11桁で入力してください。';
  }
  return null;
}


function handleCheck_(uid) {
  var data = getSheet_().getDataRange().getValues();
  var row = findRowByUid_(data, uid);
  if (row === -1) return json_({ status: 'new' });
  return json_({ status: 'found', memberId: String(data[row][COL_MEMBER_ID]) });
}

function handleRegister_(uid, params) {
  var invalid = validateRegistration_(params);
  if (invalid) return json_({ status: 'error', message: invalid });

  // 同時に登録が走ると同じ会員番号を二人に割り当ててしまうため、ここは1件ずつ通す
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ status: 'error', message: '混み合っています。少し待ってからもう一度お試しください。' });
  }

  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();

    // すでに登録済みなら、二重に作らずその会員番号を返す
    var mine = findRowByUid_(data, uid);
    if (mine !== -1) {
      return json_({ status: 'success', memberId: String(data[mine][COL_MEMBER_ID]) });
    }

    // 紙の会員などとの名寄せ
    var merge = findMergeRow_(data, normalizePhone_(params.phone), normalizeName_(params.name));
    if (merge !== -1) {
      sheet.getRange(merge + 1, COL_UID + 1).setValue(uid);
      return json_({ status: 'success', memberId: String(data[merge][COL_MEMBER_ID]) });
    }

    // 完全新規
    var newId = nextMemberId_(data);
    sheet.appendRow([
      uid,
      newId,
      params.name,
      params.kana,
      params.address,
      params.phone,
      params.email,
      Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss')
    ]);
    return json_({ status: 'success', memberId: newId });

  } finally {
    lock.releaseLock();
  }
}


function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ status: 'error', message: 'リクエストを読み取れませんでした。' });
  }

  // ここが要。クライアントの申告ではなく、LINEに確認したIDだけを使う
  var uid = verifyIdToken_(params.idToken);
  if (!uid) {
    // detail は原因の切り分け用。トークン本体は含まない
    return json_({
      status: 'unauthorized',
      message: 'ログイン情報を確認できませんでした。LINEアプリから開き直してください。',
      detail: lastVerifyError
    });
  }

  if (params.action === 'check') return handleCheck_(uid);
  if (params.action === 'register') return handleRegister_(uid, params);
  return json_({ status: 'error', message: '不明な操作です。' });
}

/**
 * 以前はここで ?action=check&uid=... を受けていたが、
 * LINE IDを知っていれば誰でも他人の会員番号を引けたため廃止した。
 * 照会も登録も doPost でIDトークンを検証したうえで行う。
 */
function doGet(e) {
  return json_({ status: 'error', message: 'この方法では利用できません。LINEアプリから会員証を開いてください。' });
}
