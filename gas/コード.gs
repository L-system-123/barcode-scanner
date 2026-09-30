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
 *      D列はパスワードをそのまま書く（暗号化はしない）。変える時も書き直すだけ
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
 *   シートに書いた平文のまま照合する（運用の手間を優先して暗号化はやめた）。
 *   以前の版が暗号化した値（32桁$64桁の英数字）が残っていても、そのまま通る。
 *   読めるようにしたければ、D列を平文のパスワードで書き直せばよい。
 *   暗号化済みの値の照合には スクリプト プロパティ の PEPPER を使うので、それが残っている間は消さない。
 *
 * ■ スキャン記録の置き場所
 *   このシート（パスワードがある方）には置かない。現場管理者に見せる別のスプレッドシート
 *   （実績シート）に書く。実績シートのIDは スクリプト プロパティ の LOG_SHEET_ID に入れる
 *   （公開リポジトリにIDを載せないため、コードには書かない）。
 *   IDを入れたら関数「実績シート準備」を1回実行する → 実績シートに「日別」「明細」「スキャン記録」
 *   ができ、このシートに残っている スキャン記録 タブがあれば中身を移してタブを消す。
 */

var SHEET_LOG = 'スキャン記録';
var SHEET_MASTER = 'マスタ';
var SHEET_USERS = 'ユーザー';
var SHEET_DAILY = '日別';
var SHEET_DETAIL = '明細';
var PROP_LOG_ID = 'LOG_SHEET_ID';
var LOG_HEADER = ['パレット番号', '運送会社名', 'カートン管理番号', 'スキャン日時', '配送CD', 'ユーザーID'];
var MASTER_HEADER = ['上4桁', '中4桁から', '中4桁まで', '配送CD', '名称', 'CD要', '色'];
var USERS_HEADER = ['ユーザーID', '名前', '有効', 'パスワード', '連続ミス', '最終ログイン', '備考'];
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
      case 'find':   return reply(withUser(req, function () { return find(req); }));
      case 'done':   return reply(withUser(req, function () { return doneToday(); }));
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
    /* パスワード欄を文字列扱いにする。数字だけのパスワードの先頭の0が消えないように */
    sh.getRange('D:D').setNumberFormat('@');

    /* トークンはスクリプト プロパティに置く。キャッシュは最長6時間で消えるためシフト中に切れる。
       期限切れは発行のたびに掃除するので溜まらない */
    var props = PropertiesService.getScriptProperties();
    var token = Utilities.getUuid() + Utilities.getUuid();
    purgeTokens(props);
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
  if (!raw) return { ok: false, error: 'auth' };
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

/* ===== 実績シート（スキャン記録の置き場所） ===== */
function logBook() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_LOG_ID);
  if (!id) throw new Error('スクリプト プロパティ ' + PROP_LOG_ID + ' に実績シートのIDを入れること');
  return SpreadsheetApp.openById(id);
}

function logSheet() {
  var sh = logBook().getSheetByName(SHEET_LOG);
  if (!sh) throw new Error('実績シートに スキャン記録 タブが無い。実績シート準備を実行すること');
  return sh;
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
    if (done) return { ok: true, pallet: done, count: rows.length, duplicate: true, codes: codeList(todayCodeMap(logSheet())) };

    var sh = logSheet();
    var prefix = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-' + user + '-';
    var seq = 0;
    var last = sh.getLastRow();
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
        var p = String(r[0]);
        if (p.indexOf(prefix) === 0) seq = Math.max(seq, +p.slice(prefix.length) || 0);
      });
    }
    /* 今日すでに別のパレットで確定済みの箱は書かない（先に確定した方を正とする）。
       画面側でも読んだ時点で止めるが、2人が同時に同じ箱を読んだ場合はここでしか止められない */
    var already = todayCodeMap(sh);
    var skipped = [];
    rows = rows.filter(function (r) {
      var hit = already[r[0]];
      if (hit) { skipped.push({ code: r[0], pallet: hit.pallet, user: hit.user }); return false; }
      return true;
    });
    /* 全部確定済みだった時はパレット番号を振らない。再送されても同じ判定になるので done: も残さない */
    if (!rows.length) return { ok: true, pallet: '', count: 0, skipped: skipped, codes: codeList(already) };

    var palletNo = prefix + ('000' + (seq + 1)).slice(-3);

    var out = rows.map(function (r) {
      return [palletNo, carrier.name, r[0], r[1], carrier.cd || '', user];
    });
    sh.getRange(last + 1, 1, out.length, LOG_HEADER.length).setValues(out);
    SpreadsheetApp.flush();

    props.setProperty('done:' + clientId, palletNo);
    purgeDone(props);
    /* 画面の「スキャン済み」一覧を、別に取りに来させずこの返事で更新する（呼び出し回数を増やさないため） */
    out.forEach(function (r) { if (!already[r[2]]) already[r[2]] = { pallet: palletNo, user: user }; });
    return { ok: true, pallet: palletNo, count: out.length, skipped: skipped, codes: codeList(already) };
  } finally {
    lock.releaseLock();
  }
}

/* ===== 今日のパレット =====
   スキャン記録は確定順に下へ追記されるので、今日の分は必ず末尾にまとまっている。
   シート全体は読まず、末尾から最大 TODAY_MAX_ROWS 行だけ見る */
var TODAY_MAX_ROWS = 5000;

/* 今日確定済みの箱: { 管理番号: { pallet, user } }。同じ番号が複数あれば先に確定した方 */
function todayCodeMap(sh) {
  var last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  var n = Math.min(last - 1, TODAY_MAX_ROWS);
  var prefix = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-';
  sh.getRange(last - n + 1, 1, n, LOG_HEADER.length).getValues().forEach(function (r) {
    var no = String(r[0]), code = String(r[2]);
    if (no.indexOf(prefix) === 0 && !map[code]) map[code] = { pallet: no, user: String(r[5]) };
  });
  return map;
}

/* 画面側の「スキャン済み」判定用。全員分を [管理番号, パレット番号, ユーザー] の並びで返す */
function doneToday() {
  return { ok: true, codes: codeList(todayCodeMap(logSheet())) };
}
function codeList(map) {
  return Object.keys(map).map(function (c) { return [c, map[c].pallet, map[c].user]; });
}

/* ===== 管理番号から探す =====
   過去の分も含めて、スキャン記録の C列 を完全一致で探す。
   同じ番号が何度も出ることはまず無いが、念のため新しい方から最大20件 */
function find(req) {
  var code = String(req.code || '');
  if (!/^\d{12}$/.test(code)) return { ok: false, error: 'bad_code' };
  var sh = logSheet();
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, hits: [] };
  var cells = sh.getRange(2, 3, last - 1, 1).createTextFinder(code).matchEntireCell(true).findAll();
  var hits = cells.slice(-20).reverse().map(function (c) {
    var r = sh.getRange(c.getRow(), 1, 1, LOG_HEADER.length).getValues()[0];
    return { pallet: String(r[0]), carrier: String(r[1]), cd: String(r[4]), user: String(r[5]),
             at: r[3] instanceof Date ? r[3].getTime() : null };
  });
  return { ok: true, hits: hits };
}

function today(req, user) {
  var sh = logSheet();
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

  var stored = pass;   // 平文のまま（冒頭の「パスワードの持ち方」参照）

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

  var m =ss.getSheetByName(SHEET_MASTER) || ss.insertSheet(SHEET_MASTER);
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

/* 実績シートを整える。何度実行してもよい（タブの作成・書式・数式は上書き、記録は重複させない）。
   1. スキャン記録 タブを用意する
   2. このシート（パスワードがある方）に スキャン記録 タブが残っていれば、実績シートへ移して消す
   3. 現場管理者が見る「日別」「明細」タブを作る */
function 実績シート準備() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = logBook();
    ss.setSpreadsheetTimeZone(TZ);

    var log = ss.getSheetByName(SHEET_LOG);
    if (!log) {
      /* 新規作成したスプレッドシートの空の「シート1」はそのまま使う */
      var first = ss.getSheets()[0];
      log = (first && first.getLastRow() === 0 && ss.getSheets().length === 1)
        ? first.setName(SHEET_LOG) : ss.insertSheet(SHEET_LOG);
    }
    if (log.getLastRow() === 0) log.appendRow(LOG_HEADER);
    log.setFrozenRows(1);
    log.getRange('A:A').setNumberFormat('@');
    log.getRange('C:C').setNumberFormat('@');                   // 管理番号の先頭0を守る
    log.getRange('D:D').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    log.getRange('E:F').setNumberFormat('@');

    var moved = moveOldLog(log);
    var dups = removeDupLog(log);

    buildDaily(ss.getSheetByName(SHEET_DAILY) || ss.insertSheet(SHEET_DAILY));
    buildDetail(ss.getSheetByName(SHEET_DETAIL) || ss.insertSheet(SHEET_DETAIL));
    ss.setActiveSheet(ss.getSheetByName(SHEET_DAILY));
    ss.moveActiveSheet(1);
    ss.setActiveSheet(ss.getSheetByName(SHEET_DETAIL));
    ss.moveActiveSheet(2);
    [ss.getSheetByName(SHEET_DAILY), ss.getSheetByName(SHEET_DETAIL), log].forEach(protectForViewers);

    var msg = '実績シートを準備しました。' + (moved === null ? '' : '\n移した記録: ' + moved + ' 行（元のタブは消しました）') +
      (dups ? '\n重複していた行を消しました: ' + dups + ' 行' : '');
    console.log(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

/* このシートに残っている スキャン記録 を実績シートへ移す。
   同じ パレット番号＋管理番号 が既にあれば書かない。移した分は日付順になるよう先頭側に差し込む
   （今日のパレットの読み取りは「記録は下ほど新しい」前提のため）。
   全部そろったのを確かめてから元のタブを消す。元のタブが無ければ null */
function moveOldLog(log) {
  var admin = SpreadsheetApp.getActive();
  var old = admin.getSheetByName(SHEET_LOG);
  if (!old) return null;

  var have = {};
  var last = log.getLastRow();
  if (last >= 2) {
    log.getRange(2, 1, last - 1, LOG_HEADER.length).getValues().forEach(function (r) {
      have[String(r[0]) + '|' + String(r[2])] = true;
    });
  }
  var rows = old.getLastRow() >= 2 ? old.getRange(2, 1, old.getLastRow() - 1, LOG_HEADER.length).getValues() : [];
  var add = rows.filter(function (r) {
    var k = String(r[0]) + '|' + String(r[2]);
    if (!String(r[0]) || have[k]) return false;
    have[k] = true;
    return true;
  });
  if (add.length) {
    log.insertRowsBefore(2, add.length);
    log.getRange(2, 1, add.length, LOG_HEADER.length).setValues(add);
    SpreadsheetApp.flush();
  }

  /* 書いたあとで読み直して、元の行が全部あるか確かめる */
  var check = {};
  log.getRange(2, 1, log.getLastRow() - 1, LOG_HEADER.length).getValues().forEach(function (r) {
    check[String(r[0]) + '|' + String(r[2])] = true;
  });
  var missing = rows.filter(function (r) { return String(r[0]) && !check[String(r[0]) + '|' + String(r[2])]; });
  if (missing.length) throw new Error('移しきれなかった行が ' + missing.length + ' 行ある。元のタブは消していない');
  admin.deleteSheet(old);
  return add.length;
}

/* 同じ日に同じ管理番号が2行以上あれば、先に確定した行（上の行）だけ残す。
   確定時の重複チェックを入れる前に、別パレットへ二重に確定された分の掃除用。消した行数を返す */
function removeDupLog(log) {
  var last = log.getLastRow();
  if (last < 3) return 0;
  var rows = log.getRange(2, 1, last - 1, LOG_HEADER.length).getValues();
  var seen = {};
  var keep = rows.filter(function (r) {
    var k = String(r[0]).slice(0, 8) + '|' + String(r[2]);
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
  var gone = rows.length - keep.length;
  if (!gone) return 0;
  log.getRange(2, 1, keep.length, LOG_HEADER.length).setValues(keep);
  log.deleteRows(2 + keep.length, gone);
  SpreadsheetApp.flush();
  return gone;
}

/* 日別: 1日1行、新しい日が上。日付はパレット番号の先頭8桁（確定した日）で数える */
function buildDaily(sh) {
  sh.clear();
  sh.getRange('A1').setValue('日別の実績（新しい日が上・自動で増えます）').setFontWeight('bold');
  sh.getRange('A2:F2').setValues([['日付', 'パレット数', '箱数', '人数', '最初の確定', '最後の確定']])
    .setFontWeight('bold').setBackground('#e8eef7');
  var L = "'" + SHEET_LOG + "'!A2:A", C = "'" + SHEET_LOG + "'!C2:C", D = "'" + SHEET_LOG + "'!D2:D", U = "'" + SHEET_LOG + "'!F2:F";
  sh.getRange('A3').setFormula(
    '=ARRAYFORMULA(IFERROR(LET(k, SORT(UNIQUE(FILTER(LEFT(' + L + ',8), ' + L + '<>"")),1,FALSE), p, k&"-*",' +
    ' HSTACK(DATE(LEFT(k,4),MID(k,5,2),RIGHT(k,2)),' +
    ' MAP(p, LAMBDA(x, COUNTUNIQUEIFS(' + L + ', ' + L + ', x))),' +
    ' MAP(p, LAMBDA(x, COUNTUNIQUEIFS(' + C + ', ' + L + ', x))),' +
    ' MAP(p, LAMBDA(x, COUNTUNIQUEIFS(' + U + ', ' + L + ', x))),' +
    ' MAP(p, LAMBDA(x, MINIFS(' + D + ', ' + L + ', x))),' +
    ' MAP(p, LAMBDA(x, MAXIFS(' + D + ', ' + L + ', x))))), "まだ記録がありません"))');
  sh.getRange('A3:A').setNumberFormat('yyyy/mm/dd (ddd)');
  sh.getRange('B3:D').setNumberFormat('#,##0');
  sh.getRange('E3:F').setNumberFormat('hh:mm');
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 130);
  for (var c = 2; c <= 6; c++) sh.setColumnWidth(c, 95);
}

/* 明細: B1 で日付を選ぶと、その日のパレット一覧・担当者別・運送会社別が出る */
function buildDetail(sh) {
  sh.clear();
  sh.getRange('A1').setValue('日付').setFontWeight('bold');
  var b1 = sh.getRange('B1');
  b1.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getParent().getSheetByName(SHEET_DAILY).getRange('A3:A'), true)
    .setAllowInvalid(false).build());
  b1.setFormula("='" + SHEET_DAILY + "'!A3").setNumberFormat('yyyy/mm/dd (ddd)')
    .setBackground('#fff8d6').setFontWeight('bold');
  sh.getRange('C1').setValue('← ここで日付を選ぶ（▼）').setFontColor('#666');

  var R = "'" + SHEET_LOG + "'!A2:F";
  var day = '"&TEXT($B$1,"yyyymmdd")&"-';
  var none = '"この日の記録はありません"';
  /* パレットごとに1行にまとめたもの。担当者別・運送会社別はこれをさらにまとめる */
  var perPallet = 'QUERY(' + R + ', "select A, F, B, count(C), min(D), max(D) where A starts with \'' + day + '\' group by A, F, B", 0)';

  sh.getRange('A3').setValue('パレット一覧（確定順）').setFontWeight('bold');
  sh.getRange('A4').setFormula('=IF($B$1="", "", IFERROR(QUERY(' + R + ', "select A, B, E, F, count(C), min(D), max(D) where A starts with \'' + day + '\'' +
    ' group by A, B, E, F order by max(D)' +
    ' label A \'パレット番号\', B \'運送会社\', E \'配送CD\', F \'担当\', count(C) \'箱数\', min(D) \'最初のスキャン\', max(D) \'最後のスキャン\'", 0), ' + none + '))');

  sh.getRange('I3').setValue('担当者別').setFontWeight('bold');
  sh.getRange('I4').setFormula('=IF($B$1="", "", IFERROR(QUERY(' + perPallet + ', "select Col2, count(Col1), sum(Col4), min(Col5), max(Col6)' +
    ' group by Col2 order by sum(Col4) desc' +
    ' label Col2 \'担当\', count(Col1) \'パレット\', sum(Col4) \'箱数\', min(Col5) \'最初\', max(Col6) \'最後\'", 1), ' + none + '))');

  sh.getRange('O3').setValue('運送会社別').setFontWeight('bold');
  sh.getRange('O4').setFormula('=IF($B$1="", "", IFERROR(QUERY(' + perPallet + ', "select Col3, count(Col1), sum(Col4)' +
    ' group by Col3 order by sum(Col4) desc' +
    ' label Col3 \'運送会社\', count(Col1) \'パレット\', sum(Col4) \'箱数\'", 1), ' + none + '))');

  ['A4:G4', 'I4:M4', 'O4:Q4'].forEach(function (a) { sh.getRange(a).setFontWeight('bold').setBackground('#e8eef7'); });
  ['F5:G', 'L5:M'].forEach(function (a) { sh.getRange(a).setNumberFormat('hh:mm'); });
  sh.setFrozenRows(4);
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(8, 24);
  sh.setColumnWidth(14, 24);
  sh.setColumnWidth(15, 150);
}

/* 現場管理者は編集者として共有しても、明細の日付欄（B1）以外は書き換えられないようにする。
   保護は所有者（このスクリプトを実行した人）だけが外せる */
function protectForViewers(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  var p = sh.protect().setDescription('現場管理者は閲覧のみ');
  if (sh.getName() === SHEET_DETAIL) p.setUnprotectedRanges([sh.getRange('B1')]);
  p.removeEditors(p.getEditors());
  if (p.canDomainEdit()) p.setDomainEdit(false);
}

/* パスワードの漏えいが疑われる時などに、ログイン中の全端末を切る */
function 全員ログアウト() {
  dropTokens('');
  SpreadsheetApp.getUi().alert('全員ログアウトしました');
}
