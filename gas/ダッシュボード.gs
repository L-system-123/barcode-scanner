/**
 * パレットスキャン実績 — ダッシュボード（Apps Script ウェブアプリ）
 *
 * 実績シート（スキャン記録があるスプレッドシート）にバインドする、スキャン受け口とは別のスクリプト。
 * パスワードのある管理者シートには触らない。ログインは無し（URLを知っていれば見られる）。
 *
 * ■ 置き場所
 *   実績シートの 拡張機能 → Apps Script に、このファイル1つだけを貼る
 *
 * ■ デプロイ
 *   デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *     次のユーザーとして実行: 自分
 *     アクセスできるユーザー: 全員
 *   直した時は デプロイを管理 → 編集（鉛筆）→「新バージョン」（URLを変えないため）
 *
 * ■ 作り
 *   ページを開くと、画面から google.script.run で getStats を呼んで数字だけ受け取り、描き直す。
 *   ページごと読み直すと GAS の入れ子 iframe が白画面になるので、自動更新も数字の取り直しで行う。
 *   同じ日の集計は30秒キャッシュする（何台で開いてもシートを読む回数を増やさないため）。
 */

var SHEET_LOG = 'スキャン記録';
var TZ = 'Asia/Tokyo';
var CACHE_SEC = 30;
var DATE_LIST_MAX = 60;

function doGet() {
  return HtmlService.createHtmlOutput(PAGE)
    .setTitle('パレットスキャン実績')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* day: 'yyyyMMdd'。空なら今日 */
function getStats(day) {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  day = /^\d{8}$/.test(String(day || '')) ? String(day) : today;

  var cache = CacheService.getScriptCache();
  var hit = cache.get('st:' + day);
  if (hit) return JSON.parse(hit);

  var st = buildStats(readLog(), day, today);
  try { cache.put('st:' + day, JSON.stringify(st), CACHE_SEC); } catch (e) { /* 大きすぎる時は諦める */ }
  return st;
}

function readLog() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
}

/* rows: [パレット番号, 運送会社名, 管理番号, スキャン日時, 配送CD, ユーザーID] */
function buildStats(rows, day, today) {
  var dates = {}, seen = {}, pallets = {}, order = [];
  var hours = {}, users = {}, carriers = {};
  var first = 0, last = 0, boxes = 0;

  rows.forEach(function (r) {
    var no = String(r[0]);
    if (!/^\d{8}-/.test(no)) return;
    var d = no.slice(0, 8);
    dates[d] = true;
    if (d !== day) return;

    var code = String(r[2]);
    if (seen[code]) return;              // 同じ日の同じ箱は1箱として数える
    seen[code] = true;

    var at = r[3] instanceof Date ? r[3].getTime() : 0;
    var user = String(r[5]), carrier = String(r[1]);
    boxes++;
    if (at) {
      if (!first || at < first) first = at;
      if (at > last) last = at;
      var h = Math.floor(((at / 3600000) + 9) % 24);   // 日本時間の時
      hours[h] = (hours[h] || 0) + 1;
    }

    var p = pallets[no];
    if (!p) {
      p = pallets[no] = { pallet: no, carrier: carrier, cd: String(r[4]), user: user, boxes: 0, first: 0, last: 0 };
      order.push(no);
    }
    p.boxes++;
    if (at && (!p.first || at < p.first)) p.first = at;
    if (at > p.last) p.last = at;

    var u = users[user] || (users[user] = { user: user, pallets: {}, boxes: 0, first: 0, last: 0 });
    u.boxes++;
    u.pallets[no] = true;
    if (at && (!u.first || at < u.first)) u.first = at;
    if (at > u.last) u.last = at;

    var c = carriers[carrier] || (carriers[carrier] = { name: carrier, pallets: {}, boxes: 0 });
    c.boxes++;
    c.pallets[no] = true;
  });

  /* パレットの確定時刻は、最後に読んだ箱の時刻で代用する（シートに確定時刻の列が無いため） */
  var palletList = order.map(function (no) { return pallets[no]; });
  var hourly = [];
  var hs = Object.keys(hours).map(Number);
  var h0 = Math.min.apply(null, hs.concat([8])), h1 = Math.max.apply(null, hs.concat([17]));
  var palletsByHour = {};
  palletList.forEach(function (p) {
    if (!p.last) return;
    var h = Math.floor(((p.last / 3600000) + 9) % 24);
    palletsByHour[h] = (palletsByHour[h] || 0) + 1;
  });
  for (var h = h0; h <= h1; h++) hourly.push({ h: h, boxes: hours[h] || 0, pallets: palletsByHour[h] || 0 });

  var userList = Object.keys(users).map(function (k) {
    var u = users[k];
    var span = (u.last - u.first) / 3600000;
    return { user: u.user, pallets: Object.keys(u.pallets).length, boxes: u.boxes, first: u.first, last: u.last,
             pace: u.boxes > 1 && span > 0 ? Math.round(u.boxes / Math.max(span, 0.25)) : null };
  }).sort(function (a, b) { return b.boxes - a.boxes; });

  var carrierList = Object.keys(carriers).map(function (k) {
    var c = carriers[k];
    return { name: c.name, pallets: Object.keys(c.pallets).length, boxes: c.boxes };
  }).sort(function (a, b) { return b.boxes - a.boxes; });

  var recent = palletList.slice().sort(function (a, b) { return b.last - a.last; }).slice(0, 10);

  return {
    day: day,
    isToday: day === today,
    dates: Object.keys(dates).sort().reverse().slice(0, DATE_LIST_MAX),
    totals: { pallets: palletList.length, boxes: boxes, people: userList.length, first: first, last: last },
    hourly: hourly,
    users: userList,
    carriers: carrierList,
    recent: recent,
    updated: Date.now()
  };
}

/* ===== 画面 =====
   このファイル1つで済むように HTML を文字列で持つ。中の JS ではバッククォートと ${ を使わないこと */
var PAGE = `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<style>
  :root {
    color-scheme: light;
    --paper: #f3f5f7; --panel: #ffffff; --ink: #151a1f; --ink-2: #5b6672; --ink-3: #8a949e;
    --line: #dde3e8; --grid: #eef1f4; --bar: #2a78d6; --bar-track: #e6edf6;
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      color-scheme: dark;
      --paper: #0e1216; --panel: #161b21; --ink: #e5eaef; --ink-2: #9aa5b1; --ink-3: #6d7884;
      --line: #28313a; --grid: #1f262d; --bar: #3987e5; --bar-track: #1d2a3a;
    }
  }
  :root[data-theme="dark"] {
    color-scheme: dark;
    --paper: #0e1216; --panel: #161b21; --ink: #e5eaef; --ink-2: #9aa5b1; --ink-3: #6d7884;
    --line: #28313a; --grid: #1f262d; --bar: #3987e5; --bar-track: #1d2a3a;
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--paper); color: var(--ink);
         font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic UI", "Meiryo", system-ui, sans-serif; line-height: 1.5; }
  .num { font-variant-numeric: tabular-nums; }
  header { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 16px; padding: 14px 16px; }
  h1 { font-size: 19px; margin: 0; margin-right: auto; }
  select { font: inherit; font-size: 15px; padding: 6px 10px; border-radius: 7px; border: 1px solid var(--line);
           background: var(--panel); color: var(--ink); }
  .upd { font-size: 12.5px; color: var(--ink-2); }
  .upd.bad { color: #d03a3a; font-weight: 700; }
  main { padding: 0 16px 24px; display: grid; gap: 14px; grid-template-columns: 1fr; max-width: 1400px; margin: 0 auto; }
  @media (min-width: 980px) { main { grid-template-columns: 1fr 1fr; } .wide { grid-column: 1 / -1; } }
  .card { background: var(--panel); border: 1px solid var(--line); border-radius: 10px; padding: 14px 16px; min-width: 0; }
  .card h2 { font-size: 14px; margin: 0 0 10px; color: var(--ink-2); font-weight: 700; }
  .tiles { display: grid; grid-template-columns: repeat(2, 1fr); gap: 14px; }
  @media (min-width: 700px) { .tiles { grid-template-columns: repeat(4, 1fr); } }
  .tile .v { font-size: 30px; font-weight: 700; line-height: 1.1; white-space: nowrap; }
  @media (min-width: 700px) { .tile .v { font-size: 38px; } }
  .tile .v small { font-size: 15px; font-weight: 400; color: var(--ink-2); margin-left: 3px; }
  .tile .k { font-size: 13px; color: var(--ink-2); margin-top: 4px; }
  .chart { position: relative; }
  .chart svg { display: block; width: 100%; height: 240px; overflow: visible; }
  .chart .ax { font-size: 11px; fill: var(--ink-3); }
  .chart .lbl { font-size: 12px; fill: var(--ink); font-weight: 700; }
  .chart .gl { stroke: var(--grid); stroke-width: 1; }
  .chart .b { fill: var(--bar); }
  .chart .hit { fill: transparent; cursor: default; }
  .chart .hit:hover + .b, .chart .b.on { opacity: .75; }
  .tip { position: absolute; pointer-events: none; background: var(--ink); color: var(--panel); font-size: 12.5px;
         padding: 5px 8px; border-radius: 6px; white-space: nowrap; transform: translate(-50%, -100%); margin-top: -8px; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th { text-align: left; font-size: 12px; color: var(--ink-2); font-weight: 700; padding: 4px 8px 6px; border-bottom: 1px solid var(--line); white-space: nowrap; }
  td { padding: 7px 8px; border-bottom: 1px solid var(--grid); white-space: nowrap; }
  th.r, td.r { text-align: right; }
  td.name { white-space: normal; }
  .scroll { overflow-x: auto; }
  .inbar { display: flex; align-items: center; gap: 8px; min-width: 130px; }
  .inbar i { display: block; height: 10px; border-radius: 0 4px 4px 0; background: var(--bar); min-width: 2px; }
  .inbar span { min-width: 3.2em; text-align: right; }
  .empty { color: var(--ink-3); font-size: 14px; padding: 18px 0; text-align: center; }
  .mono { font-family: ui-monospace, "Cascadia Mono", Consolas, monospace; font-size: 13px; }
</style>
</head>
<body>
<header>
  <h1>パレットスキャン実績</h1>
  <select id="day" aria-label="日付"></select>
  <span class="upd" id="upd">読み込み中…</span>
</header>
<main>
  <section class="card wide">
    <div class="tiles">
      <div class="tile"><div class="v num" id="tPal">—</div><div class="k">パレット</div></div>
      <div class="tile"><div class="v num" id="tBox">—</div><div class="k">箱</div></div>
      <div class="tile"><div class="v num" id="tPpl">—</div><div class="k">人</div></div>
      <div class="tile"><div class="v num" id="tSpan">—</div><div class="k">最初〜最後のスキャン</div></div>
    </div>
  </section>
  <section class="card wide">
    <h2>時間帯別の箱数</h2>
    <div class="chart" id="chart"></div>
  </section>
  <section class="card">
    <h2>担当者別</h2>
    <div class="scroll" id="users"></div>
  </section>
  <section class="card">
    <h2>運送会社別</h2>
    <div class="scroll" id="carriers"></div>
  </section>
  <section class="card wide">
    <h2>直近のパレット</h2>
    <div class="scroll" id="recent"></div>
  </section>
</main>
<script>
var REFRESH_MS = 60000;
var $ = function (id) { return document.getElementById(id); };
var current = '', loading = false, lastOk = 0, lastSt = null;

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function hm(ms) {
  if (!ms) return '—';
  var t = new Date(ms);
  return ('0' + t.getHours()).slice(-2) + ':' + ('0' + t.getMinutes()).slice(-2);
}
function fmtDay(d) {
  var t = new Date(+d.slice(0, 4), +d.slice(4, 6) - 1, +d.slice(6, 8));
  return d.slice(0, 4) + '/' + d.slice(4, 6) + '/' + d.slice(6, 8) + ' (' + '日月火水木金土'.charAt(t.getDay()) + ')';
}
function n(v) { return Number(v || 0).toLocaleString('ja-JP'); }

function load() {
  if (loading) return;
  loading = true;
  google.script.run
    .withSuccessHandler(function (st) { loading = false; lastOk = Date.now(); render(st); })
    .withFailureHandler(function (e) {
      loading = false;
      $('upd').className = 'upd bad';
      $('upd').textContent = '更新できません（' + (e && e.message || e) + '）' + (lastOk ? ' 表示は ' + hm(lastOk) + ' 時点' : '');
    })
    .getStats(current);
}

function render(st) {
  lastSt = st;
  current = st.day;
  var sel = $('day');
  var opts = st.dates.indexOf(st.day) < 0 ? [st.day].concat(st.dates) : st.dates;
  sel.innerHTML = opts.map(function (d) {
    return '<option value="' + d + '"' + (d === st.day ? ' selected' : '') + '>' + fmtDay(d) + (d === opts[0] && st.isToday && d === st.day ? '（今日）' : '') + '</option>';
  }).join('');

  $('upd').className = 'upd';
  $('upd').textContent = hm(st.updated) + ' 更新' + (st.isToday ? '・1分ごとに自動更新' : '');

  var t = st.totals;
  $('tPal').textContent = n(t.pallets);
  $('tBox').textContent = n(t.boxes);
  $('tPpl').textContent = n(t.people);
  $('tSpan').textContent = t.first ? hm(t.first) + '〜' + hm(t.last) : '—';

  drawChart(st.hourly);

  var maxU = Math.max.apply(null, st.users.map(function (u) { return u.boxes; }).concat([1]));
  $('users').innerHTML = st.users.length ? '<table><thead><tr><th>担当</th><th class="r">パレット</th><th>箱</th><th class="r">箱/時</th><th class="r">最初</th><th class="r">最後</th></tr></thead><tbody>' +
    st.users.map(function (u) {
      return '<tr><td>' + esc(u.user) + '</td><td class="r num">' + n(u.pallets) + '</td>' +
        '<td>' + inbar(u.boxes, maxU) + '</td><td class="r num">' + (u.pace == null ? '—' : n(u.pace)) + '</td>' +
        '<td class="r num">' + hm(u.first) + '</td><td class="r num">' + hm(u.last) + '</td></tr>';
    }).join('') + '</tbody></table>' : '<div class="empty">この日の記録はありません</div>';

  var maxC = Math.max.apply(null, st.carriers.map(function (c) { return c.boxes; }).concat([1]));
  $('carriers').innerHTML = st.carriers.length ? '<table><thead><tr><th>運送会社</th><th class="r">パレット</th><th>箱</th></tr></thead><tbody>' +
    st.carriers.map(function (c) {
      return '<tr><td class="name">' + esc(c.name) + '</td><td class="r num">' + n(c.pallets) + '</td><td>' + inbar(c.boxes, maxC) + '</td></tr>';
    }).join('') + '</tbody></table>' : '<div class="empty">この日の記録はありません</div>';

  $('recent').innerHTML = st.recent.length ? '<table><thead><tr><th>最後のスキャン</th><th>パレット番号</th><th>運送会社</th><th>担当</th><th class="r">箱</th></tr></thead><tbody>' +
    st.recent.map(function (p) {
      return '<tr><td class="num">' + hm(p.last) + '</td><td class="mono">' + esc(p.pallet) + '</td><td class="name">' + esc(p.carrier) + '</td>' +
        '<td>' + esc(p.user) + '</td><td class="r num">' + n(p.boxes) + '</td></tr>';
    }).join('') + '</tbody></table>' : '<div class="empty">この日の記録はありません</div>';
}

function inbar(v, max) {
  return '<div class="inbar"><span class="num">' + n(v) + '</span><i style="width:' + Math.max(2, Math.round(v / max * 100)) + '%"></i></div>';
}

/* 1系列の棒グラフ。値の数字は一番多い時間帯にだけ付け、ほかはマウス・タップで出す */
function drawChart(hourly) {
  var host = $('chart');
  var W = Math.max(host.clientWidth, 300), H = 240, L = 34, R = 6, T = 20, B = 24;
  var max = Math.max.apply(null, hourly.map(function (x) { return x.boxes; }).concat([1]));
  var step = niceStep(max), top = Math.ceil(max / step) * step;
  var cw = (W - L - R) / hourly.length, bw = Math.min(56, Math.max(4, cw - 2));   /* 隣の棒との間は最低2px。太すぎると塗りの塊に見えるので上限あり */
  var y = function (v) { return T + (H - T - B) * (1 - v / top); };
  var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="時間帯別の箱数">';
  for (var g = 0; g <= top; g += step) {
    s += '<line class="gl" x1="' + L + '" x2="' + (W - R) + '" y1="' + y(g) + '" y2="' + y(g) + '"/>' +
         '<text class="ax num" x="' + (L - 6) + '" y="' + (y(g) + 4) + '" text-anchor="end">' + g + '</text>';
  }
  var maxI = -1;
  hourly.forEach(function (x, i) { if (x.boxes === max && max > 0 && maxI < 0) maxI = i; });
  hourly.forEach(function (x, i) {
    var x0 = L + i * cw + (cw - bw) / 2, h = (H - T - B) * (x.boxes / top), yy = T + (H - T - B) - h;
    s += '<rect class="hit" x="' + (L + i * cw) + '" y="' + T + '" width="' + cw + '" height="' + (H - T - B) + '" data-i="' + i + '"/>';
    if (x.boxes > 0) s += '<path class="b" d="' + barPath(x0, yy, bw, h) + '"/>';
    if (i === maxI) s += '<text class="lbl num" x="' + (x0 + bw / 2) + '" y="' + (yy - 5) + '" text-anchor="middle">' + x.boxes + '</text>';
    if (hourly.length <= 16 || i % 2 === 0) s += '<text class="ax num" x="' + (x0 + bw / 2) + '" y="' + (H - 6) + '" text-anchor="middle">' + x.h + '時</text>';
  });
  s += '</svg><div class="tip" id="tip" hidden></div>';
  host.innerHTML = s;

  var tip = $('tip');
  Array.prototype.forEach.call(host.querySelectorAll('.hit'), function (el) {
    var show = function () {
      var x = hourly[+el.dataset.i];
      var r = el.getBoundingClientRect(), hr = host.getBoundingClientRect();
      tip.textContent = x.h + '時台  ' + x.boxes + '箱・' + x.pallets + 'パレット確定';
      tip.style.left = (r.left - hr.left + r.width / 2) + 'px';
      tip.style.top = Math.max(18, y(x.boxes) * hr.height / H) + 'px';
      tip.hidden = false;
    };
    el.addEventListener('mouseenter', show);
    el.addEventListener('click', show);
    el.addEventListener('mouseleave', function () { tip.hidden = true; });
  });
}
/* 上だけ角丸（半径4px）、下は基準線にそろえる */
function barPath(x, y, w, h) {
  var r = Math.min(4, w / 2, h);
  return 'M' + x + ',' + (y + h) + 'V' + (y + r) + 'Q' + x + ',' + y + ' ' + (x + r) + ',' + y +
         'H' + (x + w - r) + 'Q' + (x + w) + ',' + y + ' ' + (x + w) + ',' + (y + r) + 'V' + (y + h) + 'Z';
}
function niceStep(max) {
  var raw = max / 4, p = Math.pow(10, Math.floor(Math.log10(raw || 1))), m = raw / p;
  return Math.max(1, (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p);
}

$('day').onchange = function () { current = this.value; lastSt = null; load(); };
/* 自動更新は今日を見ている時だけ。過去の日は数字が変わらないので取りに行かない */
setInterval(function () {
  if (!document.hidden && (!lastSt || lastSt.isToday)) load();
}, REFRESH_MS);
document.addEventListener('visibilitychange', function () { if (!document.hidden) load(); });
var resizeT;
window.addEventListener('resize', function () {
  clearTimeout(resizeT);
  resizeT = setTimeout(function () { if (lastSt) drawChart(lastSt.hourly); }, 200);
});
load();
</script>
</body>
</html>`;
