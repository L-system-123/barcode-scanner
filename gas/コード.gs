/**
 * パレット積み付けスキャン — 受け口（Apps Script ウェブアプリ）
 *
 * スキャン画面（GitHub Pages・公開）はデータを何も持たない殻にしてある。
 * 仕分け表・運送会社名・ユーザーとパスワードはすべてこちら側に置き、
 * ログインが通った端末にだけ仕分け表を返す。
 *
 * ■ 置き場所
 *   スプレッドシートの 拡張機能 → Apps Script に貼る（シート専用のスクリプトにする）
 *
 * ■ 初回だけやること
 *   1. 関数「初期設定」を選んで実行（権限の確認が出たら許可）
 *        → タブ「スキャン記録」「マスタ」「ユーザー」ができ、メニュー「スキャン管理」が出る
 *   2. マスタ タブに仕分け表を貼る（マスタ.csv の中身）
 *   3. ユーザー タブに使う人を1行ずつ書く（A ユーザーID / C 有効 / D パスワード）
 *      D列はパスワードをそのまま書いてよい。本人が初めてログインに成功した時点で
 *      暗号化した文字列に自動で書き換わる。変える時も新しいパスワードを書き直すだけ
 *   4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分
 *        アクセスできるユーザー: 全員
 *      出てきた URL（…/exec）をスキャン画面の初回設定に貼る
 *
 * ■ コードを直したら
 *   デプロイを管理 → 編集（鉛筆）→ バージョン「新バージョン」で更新する。
 *   「新しいデプロイ」にすると URL が変わり、全端末で貼り直しになる。
 *
 * ■ パスワードの持ち方
 *   シートに手で書いた平文は、初回ログイン成功時に元に戻せない形（HMAC-SHA256）へ置き換える。
 *   平文のままだとシートの閲覧者に見えるため。書いてから初回ログインまでの間だけ平文が残る。
 *   計算に使う秘密の鍵は スクリプト プロパティ の PEPPER にある（初期設定で自動生成）。
 *   PEPPER を消したり変えたりすると全員のパスワードが通らなくなるので触らない。
 */

var SHEET_LOG = 'スキャン記録';
var SHEET_MASTER = 'マスタ';
var SHEET_USERS = 'ユーザー';
var LOG_HEADER = ['パレット番号', '運送会社名', 'カートン管理番号', 'スキャン日時', '配送CD', 'ユーザーID'];
var MASTER_HEADER = ['上4桁', '中4桁から', '中4桁まで', '配送CD', '名称', 'CD要', '色'];
var USERS_HEADER = ['ユーザーID', '名前', '有効', 'パスワード（初回ログインで暗号化）', '連続ミス', '最終ログイン', '備考'];
/* ユーザー タブの列番号（1始まり） */
var U_ID = 1, U_NAME = 2, U_ON = 3, U_HASH = 4, U_MISS = 5, U_LAST = 6, U_NOTE = 7;

var TZ = 'Asia/Tokyo';
var TOKEN_HOURS = 12;          // ログインの有効時間。1シフトより長めに
var MISS_LIMIT = 5;            // 同じIDで続けて間違えたら、シート上で無効にする回数
var FAIL_LIMIT_ALL = 30;       // 全体で間違いがこれだけ続いたら一時的に全員弾く（存在しないIDの総当たり対策）
var FAIL_WINDOW_SEC = 15 * 60;

/* ===== 入口 =====
   画面側は Content-Type: text/plain で JSON を送ってくる。
   application/json にするとブラウザが事前確認(preflight)を飛ばし、GAS はそれに応答できない */
function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return reply({ ok: false, error: 'bad_request' });
  }
  try {
    switch (req.action) {
      case 'login':  return reply(login(req));
      case 'master': return reply(withUser(req, function () { return { ok: true, master: readMaster() }; }));
      case 'commit': return reply(withUser(req, function (user) { return commit(req, user); }));
      case 'today':  return reply(withUser(req, function (user) { return today(req, user); }));
      case 'ping':   return reply(withUser(req, function () { return { ok: true }; }));
      default:       return reply({ ok: false, error: 'bad_action' });
    }
  } catch (err) {
    console.error(err);
    return reply({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

/* URLをブラウザで直接開いた時用。動いているかだけ分かればよい */
function doGet() {
  return reply({ ok: true, service: 'scan' });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ===== ユーザー ===== */

/* ユーザーIDは英数字だけ。パレット番号に入るのと、
   = や + で始まる値をシートに書くと数式として解釈されるのを防ぐため */
function normUser(v) {
  var s = String(v || '').trim().toUpperCase();
  return /^[A-Z0-9]{1,10}$/.test(s) ? s : '';
}

function usersSheet() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  if (!sh) throw new Error('ユーザー タブが無い。初期設定を実行すること');
  return sh;
}

/* 見つかれば { row: シート上の行番号, values: [...] } */
function findUser(sh, user) {
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, USERS_HEADER.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (normUser(vals[i][U_ID - 1]) === user) return { row: i + 2, values: vals[i] };
  }
  return null;
}

function isOn(v) {
  return v === true || /^(true|1|○|有効)$/i.test(String(v).trim());
}

function pepper() {
  var props = PropertiesService.getScriptProperties();
  var p = props.getProperty('PEPPER');
  if (!p) {
    p = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('PEPPER', p);
  }
  return p;
}

function hashPass(salt, pass) {
  var bytes = Utilities.computeHmacSha256Signature(salt + ':' + pass, pepper());
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* 比較にかかる時間から一致した桁数を推測されないよう、最後まで比べる */
function sameText(a, b) {
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function isHashed(stored) {
  return /^[0-9a-f]{32}\$[0-9a-f]{64}$/.test(String(stored || ''));
}

/* 暗号化済みならそれと照合、手で書いた平文ならそのまま照合する */
function checkPass(stored, pass) {
  var s = String(stored === null || stored === undefined ? '' : stored);
  if (!s || !pass) return false;
  if (!isHashed(s)) return sameText(s, pass);
  var parts = s.split('$');
  return sameText(hashPass(parts[0], pass), parts[1]);
}

function makeHash(pass) {
  var salt = Utilities.getUuid().replace(/-/g, '');
  return salt + '$' + hashPass(salt, pass);
}

/* ===== ログイン ===== */
function login(req) {
  var user = normUser(req.user);
  var pass = String(req.pass || '');
  if (!user) return { ok: false, error: 'bad_user' };

  var cache = CacheService.getScriptCache();
  var failAll = +(cache.get('fail:*') || 0);
  if (failAll >= FAIL_LIMIT_ALL) return { ok: false, error: 'busy', minutes: FAIL_WINDOW_SEC / 60 };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = usersSheet();
    var u = findUser(sh, user);

    if (!u) {
      cache.put('fail:*', String(failAll + 1), FAIL_WINDOW_SEC);
      Utilities.sleep(800);
      /* 存在しないIDとパスワード違いを区別しない。IDの一覧を探られないように */
      return { ok: false, error: 'bad_login' };
    }
    if (!isOn(u.values[U_ON - 1])) return { ok: false, error: 'disabled' };

    if (!checkPass(u.values[U_HASH - 1], pass)) {
      var miss = (+u.values[U_MISS - 1] || 0) + 1;
      sh.getRange(u.row, U_MISS).setValue(miss);
      cache.put('fail:*', String(failAll + 1), FAIL_WINDOW_SEC);
      if (miss >= MISS_LIMIT) {
        sh.getRange(u.row, U_ON).setValue(false);
        sh.getRange(u.row, U_NOTE).setValue(MISS_LIMIT + '回ミスで停止 ' + Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'));
        dropTokens(user);
        return { ok: false, error: 'disabled_now' };
      }
      Utilities.sleep(800);
      return { ok: false, error: 'bad_login', left: MISS_LIMIT - miss };
    }

    /* 照合は大文字小文字を区別しないが、パレット番号や記録にはシートA列の表記をそのまま使う */
    user = String(u.values[U_ID - 1]).trim();
    sh.getRange(u.row, U_MISS).setValue(0);
    sh.getRange(u.row, U_LAST).setValue(new Date());
    /* 手で書かれた平文は、ここで暗号化した文字列に置き換える */
    if (!isHashed(u.values[U_HASH - 1])) sh.getRange(u.row, U_HASH).setValue(makeHash(pass));
    /* パスワード欄を文字列扱いにする。数字だけのパスワードの先頭の0が消えないように */
    sh.getRange('D:D').setNumberFormat('@');

    /* トークンはスクリプト プロパティに置く。キャッシュは最長6時間で消えるためシフト中に切れる。
       期限切れは発行のたびに掃除するので溜まらない */
    var props = PropertiesService.getScriptProperties();
    var token = Utilities.getUuid() + Utilities.getUuid();
    purgeTokens(props);
    kickTokens(user);   // 同じIDで同時に使えるのは1台だけ。前の端末はここでログアウトさせる
    props.setProperty('tok:' + token, JSON.stringify({ user: user, exp: Date.now() + TOKEN_HOURS * 3600 * 1000 }));
    return { ok: true, token: token, user: user, hours: TOKEN_HOURS, master: readMaster() };
  } finally {
    lock.releaseLock();
  }
}

/* トークンが有効で、かつそのユーザーがシート上でまだ有効な時だけ通す。
   チェックを外した瞬間から、ログイン中の端末も止まる */
function withUser(req, fn) {
  var raw = PropertiesService.getScriptProperties().getProperty('tok:' + String(req.token || ''));
  if (!raw) {
    /* 別の端末のログインで追い出されたのか、期限切れなのかを画面に伝え分ける */
    var kicked = CacheService.getScriptCache().get('kicked:' + String(req.token || ''));
    return { ok: false, error: kicked ? 'kicked' : 'auth' };
  }
  var t = JSON.parse(raw);
  if (t.exp < Date.now()) return { ok: false, error: 'auth' };
  var u = findUser(usersSheet(), normUser(t.user));
  if (!u || !isOn(u.values[U_ON - 1])) return { ok: false, error: 'auth' };
  return fn(t.user);
}

function purgeTokens(props) {
  var all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('tok:') !== 0) return;
    try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); }
    catch (e) { props.deleteProperty(k); }
  });
}

/* 同じユーザーの既存トークンを消し、「別の端末でログインされた」と分かる印を6時間残す */
function kickTokens(user) {
  var props = PropertiesService.getScriptProperties();
  var cache = CacheService.getScriptCache();
  var all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('tok:') !== 0) return;
    try {
      if (normUser(JSON.parse(all[k]).user) !== normUser(user)) return;
    } catch (e) { /* 壊れたものは消すだけ */ }
    props.deleteProperty(k);
    cache.put('kicked:' + k.slice(4), '1', 21600);
  });
}

function dropTokens(user) {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('tok:') !== 0) return;
    try { if (!user || normUser(JSON.parse(all[k]).user) === normUser(user)) props.deleteProperty(k); }
    catch (e) { props.deleteProperty(k); }
  });
}

/* ===== マスタ =====
   上4桁はシートで数値に化けやすい（0201 → 201）ので、読む時に4桁へ戻す */
function readMaster() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MASTER);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, MASTER_HEADER.length).getValues();
  var out = [];
  rows.forEach(function (r) {
    var head = String(r[0]).replace(/\D/g, '');
    if (!head) return;
    var rec = { head: ('0000' + head).slice(-4), name: String(r[4]).trim() };
    var cd = String(r[3]).replace(/\D/g, '');
    if (cd) rec.cd = ('00' + cd).slice(-2);
    if (r[1] !== '' && r[2] !== '') rec.mid = [+r[1], +r[2]];
    if (r[5] === true || /^(1|true|要|必要|○|有)$/i.test(String(r[5]).trim())) rec.check = true;
    if (String(r[6]).trim()) rec.color = String(r[6]).trim();
    out.push(rec);
  });
  return out;
}

/* ===== パレット確定 =====
   連番は画面側では数えない。リロードや端末の持ち替えで重複するため、
   シート上の「その日・そのユーザー」の最大番号 + 1 をここで振る。

   通信が切れて画面側が再送してきた時に二重に書かないよう、
   画面が振った clientId で「書き込み済みか」を見る */
function commit(req, user) {
  var clientId = String(req.clientId || '');
  if (!/^[A-Za-z0-9-]{8,64}$/.test(clientId)) return { ok: false, error: 'bad_client_id' };

  var items = Array.isArray(req.items) ? req.items : [];
  if (!items.length) return { ok: false, error: 'empty' };

  /* 運送会社名は画面から受け取らず、マスタから引き直す。画面の値は信用しない。
     配送CDを持たない欄（返品・仕入・その他）は名称で引く */
  var cd = String(req.cd || '').replace(/\D/g, '');
  var master = readMaster();
  var carrier = null;
  for (var i = 0; i < master.length; i++) {
    if (cd ? master[i].cd === cd : (!master[i].cd && master[i].name === String(req.name || ''))) { carrier = master[i]; break; }
  }
  if (!carrier) return { ok: false, error: 'unknown_carrier' };

  var rows = [];
  for (var j = 0; j < items.length; j++) {
    var code = String(items[j].code || '');
    if (!/^\d{12}$/.test(code)) return { ok: false, error: 'bad_code', code: code };
    var at = new Date(+items[j].at);
    if (isNaN(at.getTime())) at = new Date();
    rows.push([code, at]);
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var props = PropertiesService.getScriptProperties();
    var done = props.getProperty('done:' + clientId);
    if (done) return { ok: true, pallet: done, count: rows.length, duplicate: true };

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
    var prefix = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-' + user + '-';
    var seq = 0;
    var last = sh.getLastRow();
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
        var p = String(r[0]);
        if (p.indexOf(prefix) === 0) seq = Math.max(seq, +p.slice(prefix.length) || 0);
      });
    }
    var palletNo = prefix + ('000' + (seq + 1)).slice(-3);

    var out = rows.map(function (r) {
      return [palletNo, carrier.name, r[0], r[1], carrier.cd || '', user];
    });
    sh.getRange(last + 1, 1, out.length, LOG_HEADER.length).setValues(out);
    SpreadsheetApp.flush();

    props.setProperty('done:' + clientId, palletNo);
    purgeDone(props);
    return { ok: true, pallet: palletNo, count: out.length };
  } finally {
    lock.releaseLock();
  }
}

/* ===== 今日のパレット =====
   スキャン記録は確定順に下へ追記されるので、今日の分は必ず末尾にまとまっている。
   シート全体は読まず、末尾から最大 TODAY_MAX_ROWS 行だけ見る */
var TODAY_MAX_ROWS = 5000;

function today(req, user) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, pallets: [] };
  var n = Math.min(last - 1, TODAY_MAX_ROWS);
  var vals = sh.getRange(last - n + 1, 1, n, LOG_HEADER.length).getValues();

  var prefix = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-';
  var me = normUser(user);
  var byNo = {}, order = [];
  vals.forEach(function (r) {
    var no = String(r[0]);
    if (no.indexOf(prefix) !== 0) return;
    if (!req.all && normUser(r[5]) !== me) return;
    var p = byNo[no];
    if (!p) {
      p = byNo[no] = { pallet: no, carrier: String(r[1]), cd: String(r[4]), user: String(r[5]), items: [] };
      order.push(no);
    }
    var at = r[3] instanceof Date ? r[3].getTime() : null;
    p.items.push({ code: String(r[2]), at: at });
  });
  /* 新しく確定したものを上に */
  return { ok: true, pallets: order.reverse().map(function (no) { return byNo[no]; }) };
}

/* 再送判定用の記録は直近300件だけ残す。プロパティには容量上限がある */
function purgeDone(props) {
  var keys = Object.keys(props.getProperties()).filter(function (k) { return k.indexOf('done:') === 0; });
  if (keys.length <= 300) return;
  keys.slice(0, keys.length - 300).forEach(function (k) { props.deleteProperty(k); });
}

/* ===== シートのメニュー ===== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('スキャン管理')
    .addItem('パスワード設定（登録・変更・停止解除）', 'パスワード設定')
    .addSeparator()
    .addItem('全員ログアウト', '全員ログアウト')
    .addItem('初期設定', '初期設定')
    .addToUi();
}

/* パスワードは画面に出ない入力欄で受け取る（ui.prompt だと平文で見える） */
function パスワード設定() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font:14px sans-serif}label{display:block;margin:10px 0 4px}input{width:100%;padding:6px;box-sizing:border-box}' +
    'button{margin-top:14px;padding:8px 18px}#msg{margin-top:10px;color:#b00}</style>' +
    '<label>ユーザーID（英数字10文字まで）</label><input id="id" autocomplete="off">' +
    '<label>名前（任意・新規の時だけ使う）</label><input id="nm" autocomplete="off">' +
    '<label>パスワード（8文字以上）</label><input id="pw" type="password" autocomplete="new-password">' +
    '<label>もう一度</label><input id="pw2" type="password" autocomplete="new-password">' +
    '<button id="go">設定</button><div id="msg"></div>' +
    '<script>' +
    'document.getElementById("go").onclick=function(){' +
    ' var id=document.getElementById("id").value,nm=document.getElementById("nm").value,' +
    '     a=document.getElementById("pw").value,b=document.getElementById("pw2").value,m=document.getElementById("msg");' +
    ' if(a!==b){m.textContent="パスワードが一致しません";return;}' +
    ' m.style.color="#555";m.textContent="設定中…";' +
    ' google.script.run.withSuccessHandler(function(r){m.style.color="#070";m.textContent=r;' +
    '   document.getElementById("pw").value="";document.getElementById("pw2").value="";})' +
    ' .withFailureHandler(function(e){m.style.color="#b00";m.textContent=e.message;})' +
    ' .setPasswordFromDialog(id,nm,a);};' +
    '</script>'
  ).setWidth(360).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'パスワード設定');
}

/* ダイアログから呼ばれる。登録・変更・停止解除を兼ねる */
function setPasswordFromDialog(id, name, pass) {
  var user = normUser(id);
  if (!user) throw new Error('ユーザーIDは英数字10文字までです');
  if (String(pass).length < 8) throw new Error('パスワードは8文字以上にしてください');

  var stored = makeHash(pass);

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = usersSheet();
    var u = findUser(sh, user);
    var nm = String(name || '').replace(/^[=+\-@]+/, '').trim();
    if (u) {
      sh.getRange(u.row, U_HASH).setValue(stored);
      sh.getRange(u.row, U_ON).setValue(true);
      sh.getRange(u.row, U_MISS).setValue(0);
      sh.getRange(u.row, U_NOTE).setValue('');
      if (nm) sh.getRange(u.row, U_NAME).setValue(nm);
      dropTokens(user);   // 古いパスワードでログイン中の端末は切る
      return user + ' のパスワードを変更しました（有効に戻しています）';
    }
    var row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, USERS_HEADER.length).setValues([[user, nm, true, stored, 0, '', '']]);
    sh.getRange(row, U_ON).insertCheckboxes();
    return user + ' を登録しました';
  } finally {
    lock.releaseLock();
  }
}

/* ===== 初回だけ手で実行する ===== */
function 初期設定() {
  var ss = SpreadsheetApp.getActive();

  var log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) log.appendRow(LOG_HEADER);
  log.setFrozenRows(1);
  log.getRange('A:A').setNumberFormat('@');
  log.getRange('C:C').setNumberFormat('@');                   // 管理番号の先頭0を守る
  log.getRange('D:D').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  log.getRange('E:F').setNumberFormat('@');

  var m = ss.getSheetByName(SHEET_MASTER) || ss.insertSheet(SHEET_MASTER);
  if (m.getLastRow() === 0) m.appendRow(MASTER_HEADER);
  m.setFrozenRows(1);
  m.getRange('A:A').setNumberFormat('@');
  m.getRange('D:D').setNumberFormat('@');

  var u = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  if (u.getLastRow() === 0) u.appendRow(USERS_HEADER);
  u.setFrozenRows(1);
  u.getRange('A:A').setNumberFormat('@');
  u.getRange('D:D').setNumberFormat('@');
  u.getRange('F:F').setNumberFormat('yyyy/mm/dd hh:mm');

  pepper();   // パスワード用の秘密鍵を先に作っておく
  onOpen();
  SpreadsheetApp.getUi().alert('タブを用意しました。\n\n次は マスタ タブに仕分け表を貼り、\nメニュー「スキャン管理 → パスワード設定」で使う人を登録してください。');
}

/* パスワードの漏えいが疑われる時などに、ログイン中の全端末を切る */
function 全員ログアウト() {
  dropTokens('');
  SpreadsheetApp.getUi().alert('全員ログアウトしました');
}
