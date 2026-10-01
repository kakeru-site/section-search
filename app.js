// セクション検索 — 画面と操作
import { createStore, DEMO, localKey } from "./store.js";
import * as S from "./search.js";
import * as FID from "./faceid.js";

/* ============================================================
   小道具
   ============================================================ */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const vals = (o) => Object.entries(o || {}).map(([id, v]) => ({ id, ...v }));
const pad = (n) => String(n).padStart(2, "0");
const WD = ["日", "月", "火", "水", "木", "金", "土"];
const fmtHM = (t) => { const d = new Date(t); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const fmtDay = (t) => { const d = new Date(t); return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())}（${WD[d.getDay()]}）`; };
const dayKey = (t) => { const d = new Date(t); return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`; };
function fmtWhen(t) {
  if (!t) return "";
  const now = new Date(), d = new Date(t);
  if (dayKey(t) === dayKey(now)) return fmtHM(t);
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (dayKey(t) === dayKey(y)) return "昨日 " + fmtHM(t);
  return `${d.getMonth() + 1}/${d.getDate()} ${fmtHM(t)}`;
}
const safeKey = (s) => String(s).replace(/[.#$\[\]\/]/g, "_");
function debounce(fn, ms) { let t; const f = (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; f.flush = (...a) => { clearTimeout(t); fn(...a); }; f.cancel = () => clearTimeout(t); return f; }

const IC = {
  search: '<svg viewBox="0 0 24 24" class="ic"><circle cx="10.5" cy="10.5" r="6.5" fill="none" stroke="currentColor" stroke-width="2.2"/><path d="M15.5 15.5L21 21" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/></svg>',
  note: '<svg viewBox="0 0 24 24" class="ic"><path d="M5 3h10l4 4v14H5z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><path d="M8 11h8M8 15h8M8 7h5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>',
  face: '<svg viewBox="0 0 24 24" class="ic"><path d="M3 8V5a2 2 0 0 1 2-2h3M16 3h3a2 2 0 0 1 2 2v3M21 16v3a2 2 0 0 1-2 2h-3M8 21H5a2 2 0 0 1-2-2v-3" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M9 9v1.5M15 9v1.5M12 9v4h-1M9 16c1.8 1.4 4.2 1.4 6 0" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
  logo: '<svg viewBox="0 0 24 24" class="ic"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>',
  star: '<svg viewBox="0 0 24 24" class="ic" style="width:12px;height:12px"><path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z" fill="currentColor"/></svg>',
  ext: '<svg viewBox="0 0 24 24" class="ic" style="width:12px;height:12px"><path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
};

const HELLO = ["締め付けトルクは守ろう", "品番は声に出してもう一度", "迷ったら型式から確認", "左右・前後の聞き忘れに注意", "今日も1件ずつ、正確に"];

/* ============================================================
   状態
   ============================================================ */
let store;
const st = {
  user: null, locked: false, booted: false,
  view: "sec",
  parts: {}, partList: [], kata: {}, memos: {}, history: {}, settings: {}, stats: {},
  loaded: {}, syn: [],
  // セクション検索
  q: "", group: "all", sec: "", sel: 0, limit: 60, results: [],
  // 型式
  kq: "",
  // メモ
  memoFilter: "open", memoQ: "", memoSel: null, memoPart: null, memoKata: null,
  // 履歴
  histFilter: "all", histQ: "",
  // 編集
  editTab: "parts", editQ: "", kataQ: "",
  hello: HELLO[Math.floor(Math.random() * HELLO.length)],
  lastHist: null, lastActive: Date.now(), hiddenAt: 0,
};
const setting = (k, d) => (st.settings && st.settings[k] != null && st.settings[k] !== "" ? st.settings[k] : d);

/* ============================================================
   トースト・モーダル
   ============================================================ */
let toastT;
function toast(msg, ng = false) {
  $(".toast")?.remove();
  const el = document.createElement("div");
  el.className = "toast" + (ng ? " ng" : "");
  el.textContent = msg;
  document.body.appendChild(el);
  clearTimeout(toastT);
  toastT = setTimeout(() => el.remove(), ng ? 4200 : 2200);
}
function modal(html, { onOpen } = {}) {
  return new Promise((resolve) => {
    const root = $("#modal-root");
    root.innerHTML = `<div class="modal-bg"><div class="modal" role="dialog" aria-modal="true">${html}</div></div>`;
    const bg = $(".modal-bg", root);
    const close = (v) => { root.innerHTML = ""; resolve(v); };
    bg.addEventListener("mousedown", (e) => { if (e.target === bg) close(null); });
    bg.addEventListener("click", (e) => {
      const b = e.target.closest("[data-m]");
      if (!b) return;
      const act = b.dataset.m;
      if (act === "cancel") return close(null);
      const form = $("form", bg);
      if (form && act === "ok" && !form.reportValidity()) return;
      const data = form ? Object.fromEntries(new FormData(form)) : {};
      close({ act, data });
    });
    bg.addEventListener("submit", (e) => { e.preventDefault(); $('[data-m="ok"]', bg)?.click(); });
    bg.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.stopPropagation(); close(null); } });
    onOpen && onOpen(bg);
    ($("[autofocus]", bg) || $("input,textarea,button", bg))?.focus();
  });
}
async function confirmBox(msg, okLabel = "OK", danger = false) {
  const r = await modal(`<h3>確認</h3><p style="line-height:1.7;white-space:pre-line;margin:0">${esc(msg)}</p>
    <div class="modal-foot"><button class="btn" data-m="cancel">やめる</button><button class="btn ${danger ? "danger" : "pri"}" data-m="ok" autofocus>${esc(okLabel)}</button></div>`);
  return !!r;
}
const modalOpen = () => !!$("#modal-root").children.length;

/* ============================================================
   起動・ログイン・ロック
   ============================================================ */
async function boot() {
  try { store = await createStore(); }
  catch (e) { showGateError("Firebaseの読み込みに失敗しました。ネットワークか js/config.js を確認してください。\n" + e.message); return; }
  store.onAuth((u) => {
    const was = st.user;
    st.user = u;
    if (!u) { teardown(); showLogin(); return; }
    if (!was) {
      st.locked = !!FID.registered();
      startData();
    }
    route();
  });
}
function showGateError(msg) {
  const g = $("#gate"); g.classList.remove("hidden");
  g.innerHTML = `<div class="stripe"></div><div class="gate-body"><div class="gate-card"><div class="gate-logo">${IC.logo}<span>SECTION<b>検索</b></span></div><p class="gate-err" style="white-space:pre-line">${esc(msg)}</p></div></div>`;
}
function route() {
  $("#gate").classList.toggle("hidden", !!st.user);
  $("#lock").classList.toggle("hidden", !st.user || !st.locked);
  $("#app").classList.toggle("hidden", !st.user || st.locked);
  if (st.user && st.locked) showLock();
  if (st.user && !st.locked) { if (!st.booted) { st.booted = true; renderShell(); } }
}
function showLogin() {
  const g = $("#gate");
  g.innerHTML = `<div class="stripe"></div><div class="gate-body"><form class="gate-card" id="login-form" autocomplete="on">
    <div class="gate-logo">${IC.logo}<span>SECTION<b>検索</b></span></div>
    <div class="gate-sub">部品名からセクション・部品コードを引く店頭業務ツール</div>
    ${DEMO ? `<div class="demo-badge" style="margin-bottom:14px">デモモード：Firebase未設定のため、何を入れてもログインできます。データはこのブラウザにだけ保存されます。</div>` : ""}
    <div class="field"><label for="lg-email">メールアドレス</label><input class="inp" id="lg-email" type="email" name="email" autocomplete="username" required ${DEMO ? 'value="demo@example.com"' : ""}></div>
    <div class="field"><label for="lg-pw">パスワード</label><input class="inp" id="lg-pw" type="password" name="pw" autocomplete="current-password" ${DEMO ? "" : "required"}></div>
    <div class="gate-err" id="lg-err"></div>
    <button class="btn pri wide" type="submit">ログイン</button>
    <div style="margin-top:14px;text-align:right"><button class="linkbtn" type="button" id="lg-reset">パスワードを忘れた</button></div>
  </form></div>`;
  route();
  const f = $("#login-form");
  f.addEventListener("submit", async (e) => {
    e.preventDefault();
    const btn = $("button[type=submit]", f); btn.disabled = true; $("#lg-err").textContent = "";
    try { await store.login(f.email.value.trim(), f.pw.value); }
    catch (err) { $("#lg-err").textContent = err.message; }
    finally { btn.disabled = false; }
  });
  $("#lg-reset").addEventListener("click", async () => {
    const em = f.email.value.trim();
    if (!em) { $("#lg-err").textContent = "先にメールアドレスを入力してください"; return; }
    try { await store.resetPassword(em); toast("再設定メールを送りました"); } catch (err) { $("#lg-err").textContent = err.message; }
  });
  setTimeout(() => $("#lg-email")?.focus(), 30);
}
function showLock() {
  const l = $("#lock");
  if ($("#lock-card", l)) return;
  l.innerHTML = `<div class="stripe"></div><div class="gate-body"><div class="gate-card center" id="lock-card">
    <div class="gate-logo" style="justify-content:center">${IC.logo}<span>SECTION<b>検索</b></span></div>
    <div class="gate-sub">ロック中です</div>
    <svg viewBox="0 0 24 24" class="lock-face"><path d="M3 8V5a2 2 0 0 1 2-2h3M16 3h3a2 2 0 0 1 2 2v3M21 16v3a2 2 0 0 1-2 2h-3M8 21H5a2 2 0 0 1-2-2v-3" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/><path d="M9 9v1.5M15 9v1.5M12 9v4h-1M9 16c1.8 1.4 4.2 1.4 6 0" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
    <button class="btn pri wide" id="unlock">生体認証で解除</button>
    <div class="gate-err" id="lock-err" style="margin-top:10px"></div>
    <button class="linkbtn" id="lock-pw">パスワードでログインし直す</button>
  </div></div>`;
  const go = async () => {
    $("#lock-err").textContent = "";
    try { await FID.verify(); st.locked = false; st.lastActive = Date.now(); l.innerHTML = ""; route(); toast("ロックを解除しました"); }
    catch (e) { $("#lock-err").textContent = e.name === "NotAllowedError" ? "キャンセルされたか、時間切れになりました" : "解除できませんでした：" + e.message; }
  };
  $("#unlock").addEventListener("click", go);
  $("#lock-pw").addEventListener("click", async () => { l.innerHTML = ""; st.locked = false; await store.logout(); });
  setTimeout(() => $("#unlock")?.focus(), 30);
}
function lockNow() {
  if (!FID.registered() || !st.user || st.locked) return;
  st.locked = true; route();
}
// 自動ロック（一定時間操作なし / 一定時間画面を離れていた）
["pointerdown", "keydown", "wheel"].forEach((ev) => addEventListener(ev, () => { st.lastActive = Date.now(); }, { passive: true }));
document.addEventListener("visibilitychange", () => {
  if (document.hidden) st.hiddenAt = Date.now();
  else if (st.hiddenAt && Date.now() - st.hiddenAt > FID.settings().lockMin * 60000) lockNow();
});
setInterval(() => { if (Date.now() - st.lastActive > FID.settings().lockMin * 60000) lockNow(); }, 30000);

/* ============================================================
   データ購読
   ============================================================ */
let unsubs = [];
function teardown() { unsubs.forEach((u) => typeof u === "function" && u()); unsubs = []; st.booted = false; st.loaded = {}; $("#app").innerHTML = ""; }
function startData() {
  const onErr = (e) => toast(e.message, true);
  const sub = (path, key, after) => unsubs.push(store.sub(path, (v) => { st[key] = v || {}; st.loaded[key] = true; after && after(); onData(key); }, onErr));
  sub("parts", "parts", () => { st.partList = vals(st.parts).map(S.indexPart); });
  sub("kata", "kata");
  sub("memos", "memos");
  sub("history", "history");
  sub("settings", "settings", () => { st.syn = S.parseSynonyms(setting("synonyms", S.DEFAULT_SYNONYMS)); });
  sub("stats", "stats");
}
function onData(key) {
  if (!st.booted) return;
  renderNav();
  renderRight();
  const v = st.view;
  if (v === "sec" && ["parts", "memos", "settings", "stats", "history"].includes(key)) renderResults();
  else if (v === "kata" && ["kata", "memos"].includes(key)) renderKataResult();
  else if (v === "memo" && ["memos", "parts"].includes(key)) renderMemo(false);
  else if (v === "hist" && key === "history") renderHist();
  else if (v === "edit") {
    if (st.editTab === "parts" && key === "parts" && $("#etbl")) editPartsTable();
    else if (st.editTab === "kata" && key === "kata" && $("#ktbl")) editKataTable();
    else if (st.editTab === "io" && ["parts", "kata"].includes(key) && !modalOpen()) editIO();
  }
}

/* ============================================================
   画面の骨組み
   ============================================================ */
const NAV = [
  ["sec", "セクション検索", "S"],
  ["kata", "型式 → 車種", "K"],
  ["memo", "メモ", "M"],
  ["hist", "検索履歴", "H"],
  ["edit", "データ編集", "E"],
];
function renderShell() {
  $("#app").innerHTML = `<div class="stripe"></div><div class="shell">
    <aside class="side"><div class="logo">${IC.logo}<span>SECTION<b>検索</b></span></div><nav id="nav"></nav>
      <div class="side-foot">${DEMO ? '<div class="demo-badge">デモモード（このブラウザにだけ保存）</div>' : ""}<button class="faceid" id="fid-card" data-go="edit:settings"></button>
      <div class="who"><span id="who"></span><button class="linkbtn" id="logout">ログアウト</button></div></div></aside>
    <main id="main"></main>
    <aside class="right" id="right"></aside></div>`;
  $("#who").textContent = st.user?.email || "";
  $("#logout").addEventListener("click", async () => { if (await confirmBox("ログアウトしますか？", "ログアウト")) store.logout(); });
  renderNav(); renderFid(); go(st.view);
}
function renderNav() {
  const open = vals(st.memos).filter((m) => !m.done).length;
  $("#nav").innerHTML = NAV.map(([k, l, key]) => `<button class="nav ${st.view === k ? "on" : ""}" data-view="${k}"><span>${l}${k === "memo" && open ? `<span class="badge">${open}</span>` : ""}</span><kbd>${key}</kbd></button>`).join("");
}
async function renderFid() {
  const c = $("#fid-card"); if (!c) return;
  const reg = FID.registered();
  if (reg) { c.className = "faceid on"; c.innerHTML = `${IC.face}<div><b>生体認証ロック ON</b>${FID.settings().lockMin}分操作なしでロック</div>`; }
  else { const av = await FID.platformAvailable(); c.className = "faceid"; c.innerHTML = `${IC.face}<div><b>生体認証ロック OFF</b>${av ? "押して設定する" : "この端末では使えません"}</div>`; }
}
function go(view, opt = {}) {
  st.view = view;
  if (opt.editTab) st.editTab = opt.editTab;
  renderNav();
  const m = $("#main"); m.scrollTop = 0;
  if (view === "sec") renderSec();
  else if (view === "kata") renderKata();
  else if (view === "memo") renderMemo(true);
  else if (view === "hist") renderHist();
  else if (view === "edit") renderEdit();
  renderRight();
}

/* ============================================================
   右パネル
   ============================================================ */
function renderRight() {
  const r = $("#right"); if (!r) return;
  if (st.view === "kata") {
    const list = vals(st.kata).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)).slice(0, 10);
    r.innerHTML = `<div class="box"><div class="bt">型式の対応表 <span class="num">${Object.keys(st.kata).length}件</span><button class="more" data-go="edit:kata">編集 →</button></div>
      <ul class="list">${list.map((k) => `<li><button data-kq="${esc(k.id)}"><span class="kc">${esc(k.id)}</span><span>${esc(k.name)}${k.verified ? "" : ' <small style="color:var(--ac2)">未確認</small>'}</span></button></li>`).join("") || '<li class="empty">まだ登録がありません</li>'}</ul></div>
      <div class="box"><div class="bt">しくみ</div><p style="font-size:12px;line-height:1.7;color:var(--mu);margin:0">ハイフンより後ろが、登録した型式コードで<b style="color:var(--tx)">終わっているか</b>を判定します。<br>例：GC27・HFC27・C27 → すべて「C27」<br><br>車種名（例：セレナ）を入れると、型式を逆引きできます。</p></div>`;
    return;
  }
  const hist = vals(st.history).sort((a, b) => b.at - a.at).slice(0, 8);
  const open = vals(st.memos).filter((m) => !m.done).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  r.innerHTML = `<div class="box"><div class="bt">検索履歴<button class="more" data-view="hist">すべて →</button></div>
    <ul class="list">${hist.map((h) => `<li><button data-rerun="${esc(h.id)}"><time>${fmtWhen(h.at)}</time><span>${h.type === "kata" ? '<small style="color:var(--ac2)">型式 </small>' : ""}${esc(h.q)}</span></button></li>`).join("") || '<li class="empty">まだ履歴がありません</li>'}</ul></div>
    <div class="box"><div class="bt">未解決メモ ${open.length ? `<span class="num">${open.length}</span>` : ""}<button class="more" data-view="memo">一覧 →</button></div>
    <ul class="list">${open.slice(0, 6).map((m) => `<li><button data-memo="${esc(m.id)}"><span class="dot"></span><span>${esc(m.title || "（無題）")}</span></button></li>`).join("") || '<li class="empty">未解決のメモはありません</li>'}</ul>
    <button class="btn sm" style="margin-top:8px" data-newmemo="1">＋ メモを書く</button></div>`;
}

/* ============================================================
   セクション検索
   ============================================================ */
function renderSec() {
  const hot = st.partList.filter((p) => p.hot);
  const tip = hot.length && Math.random() < 0.6 ? (() => { const p = hot[Math.floor(Math.random() * hot.length)]; return `${p.name}は「${p.hot}」`; })() : st.hello;
  $("#main").innerHTML = `<div class="view-head"><div><h1 class="h1">セクション検索</h1>
    <div class="tip">部品名の一部・ひらがなでもOK ／ <kbd>/</kbd> 検索欄へ ／ <kbd>↑↓</kbd> 選択 ／ <kbd>Enter</kbd> Google ／ <kbd>Ctrl</kbd>+<kbd>Enter</kbd> メモ</div></div></div>
    <p class="hello">今日のひとこと：${esc(tip)}</p>
    <label class="search">${IC.search}<input id="q" type="search" autocomplete="off" spellcheck="false" placeholder="部品名・部品コード・セクションで検索（例：べると、15208）" value="${esc(st.q)}" data-live="1"><span class="cnt" id="qcnt"></span><button class="clear" id="qclear" title="消す（Esc）" aria-label="検索語を消す">×</button></label>
    <div class="chips" id="chips"></div>
    <div id="results"></div>`;
  const q = $("#q");
  q.addEventListener("input", () => { st.q = q.value; st.sel = 0; st.limit = 60; renderResults(); histTyping(); });
  q.addEventListener("keydown", secKeys);
  $("#qclear").addEventListener("click", (e) => { e.preventDefault(); clearSec(); q.focus(); });
  renderResults();
  setTimeout(() => q.focus(), 0);
}
function clearSec() { st.q = ""; st.sec = ""; st.group = "all"; st.sel = 0; const q = $("#q"); if (q) q.value = ""; renderResults(); }
function renderChips() {
  const el = $("#chips"); if (!el) return;
  el.innerHTML = S.GROUPS.map((g) => `<button class="chip ${st.group === g.id ? "on" : ""}" data-group="${g.id}">${esc(g.label)}</button>`).join("") +
    (st.sec ? `<button class="chip secf" data-secoff="1" title="セクションの絞り込みを解除">SEC ${esc(st.sec)} ×</button>` : "");
}
function highlight(text, ranges) {
  if (!ranges || !ranges.length) return esc(text);
  let out = "", i = 0;
  for (const [s, e] of ranges) { out += esc(text.slice(i, s)) + "<mark>" + esc(text.slice(s, e)) + "</mark>"; i = e; }
  return out + esc(text.slice(i));
}
const memosForPart = (id) => vals(st.memos).filter((m) => m.partId === id);
function renderResults() {
  renderChips();
  const box = $("#results"); if (!box) return;
  const cnt = $("#qcnt");
  if (!st.loaded.parts) { box.innerHTML = '<p class="empty">読み込み中…</p>'; return; }
  if (!st.partList.length) {
    cnt.textContent = "";
    box.innerHTML = `<div class="panel"><h3>部品データがまだありません</h3><p style="margin:0 0 12px;font-size:13px;color:var(--mu);line-height:1.7">「データ編集 → 取り込み・書き出し」から、parts.json か Excel（部品マスタ）を取り込んでください。</p><button class="btn pri" data-go="edit:io">取り込み画面へ</button></div>`;
    return;
  }
  const home = !st.q.trim() && st.group === "all" && !st.sec;
  if (home) { cnt.textContent = `全${st.partList.length}件`; cnt.className = "cnt"; box.innerHTML = homePanels(); st.results = []; return; }
  st.results = S.searchParts(st.partList, st.q, st.syn, { group: st.group, sec: st.sec });
  const n = st.results.length;
  cnt.textContent = n ? `${n}件ヒット` : "見つかりません";
  cnt.className = "cnt " + (n ? "" : "ng");
  if (!n) {
    const q = st.q.trim();
    box.innerHTML = `<div class="panel"><h3>「${esc(q)}」は見つかりませんでした</h3>
      <p style="margin:0 0 12px;font-size:12.5px;color:var(--mu);line-height:1.7">言い方を変える（例：フィルター → エレメント）か、短くしてみてください。よく使う言い換えは「データ編集 → 設定 → 言い換えリスト」に登録できます。</p>
      <div class="flex-wrap"><button class="btn" data-google="${esc(q)}">${IC.search} Googleで「${esc(setting("googlePrefix", "日産"))} ${esc(q)}」</button><button class="btn" data-addpart="${esc(q)}">＋ 部品データに追加</button></div></div>`;
    return;
  }
  if (st.sel >= Math.min(n, st.limit)) st.sel = 0;
  const rows = st.results.slice(0, st.limit).map((r, i) => {
    const p = r.p, mc = memosForPart(p.id).length;
    return `<div class="row ${i === st.sel ? "sel" : ""}" data-pid="${esc(p.id)}" data-i="${i}">
      <div><div class="pname">${highlight(p.name, r.ranges)}</div>${p.note ? `<div class="pnote">${esc(p.note)}</div>` : ""}${p.hot ? `<div class="phot">${IC.star}頻出：${esc(p.hot)}</div>` : ""}</div>
      <div><button class="sec" data-sec="${esc(p.sec)}" title="このセクションの部品を一覧">SEC<b>${esc(p.sec)}</b></button></div>
      <div class="pcode">${esc(p.code)}</div>
      <div class="acts"><button class="btn" data-act="google" title="Googleで検索（Enter）">${IC.search}Google</button><button class="btn ${mc ? "has" : ""}" data-act="memo" title="メモ（Ctrl+Enter）">${IC.note}メモ${mc ? `<span class="cnt">${mc}</span>` : ""}</button></div></div>`;
  }).join("");
  box.innerHTML = `<div class="thead"><span>部品名</span><span>セクション</span><span>部品コード</span><span>ジャンプ</span></div><div class="rows">${rows}</div>` +
    (n > st.limit ? `<div class="more-row"><button class="btn" data-more="1">さらに表示（残り${n - st.limit}件）</button></div>` : "");
}
function homePanels() {
  const stats = vals(st.stats.sections).sort((a, b) => (b.n || 0) - (a.n || 0)).slice(0, 10);
  const hot = st.partList.filter((p) => p.hot).sort((a, b) => S.cmpSec(a.sec, b.sec));
  const recent = [], seen = new Set();
  for (const h of vals(st.history).filter((h) => h.type !== "kata").sort((a, b) => b.at - a.at)) { if (!seen.has(h.q)) { seen.add(h.q); recent.push(h.q); } if (recent.length >= 10) break; }
  return `${stats.length ? `<div class="panel"><h3>よく引くセクション</h3><div class="flex-wrap">${stats.map((s) => `<button class="qchip" data-sec="${esc(s.sec)}"><b>${esc(s.sec)}</b><small>${s.n}回</small></button>`).join("")}</div></div>` : ""}
    ${recent.length ? `<div class="panel"><h3>最近の検索</h3><div class="flex-wrap">${recent.map((q) => `<button class="qchip" data-q="${esc(q)}">${esc(q)}</button>`).join("")}</div></div>` : ""}
    ${hot.length ? `<div class="panel"><h3>${IC.star} 頻出パーツ（上位20％）</h3><div class="flex-wrap">${hot.map((p) => `<button class="qchip" data-q="${esc(p.name)}" title="${esc(p.hot)}"><b>${esc(p.sec)}</b>${esc(p.name)}</button>`).join("")}</div></div>` : ""}
    <p class="sample">登録部品 ${st.partList.length}件 ／ 上のチップを押すか、検索欄に入力してください。</p>`;
}
function secKeys(e) {
  const n = Math.min(st.results.length, st.limit);
  if (e.key === "ArrowDown" && n) { e.preventDefault(); st.sel = (st.sel + 1) % n; markSel(); }
  else if (e.key === "ArrowUp" && n) { e.preventDefault(); st.sel = (st.sel - 1 + n) % n; markSel(); }
  else if (e.key === "Enter" && n) {
    e.preventDefault();
    const p = st.results[st.sel]?.p; if (!p) return;
    if (e.ctrlKey || e.metaKey || e.shiftKey) openMemoForPart(p); else googlePart(p);
  } else if (e.key === "Escape") { e.preventDefault(); clearSec(); }
}
function markSel() {
  $$(".row").forEach((r) => r.classList.toggle("sel", +r.dataset.i === st.sel));
  $(`.row[data-i="${st.sel}"]`)?.scrollIntoView({ block: "nearest" });
}
function googleOpen(q) { window.open("https://www.google.com/search?q=" + encodeURIComponent(q), "_blank", "noopener"); }
function googlePart(p) {
  histTyping.flush();
  googleOpen(`${setting("googlePrefix", "日産")} ${p.name}`.trim());
  bumpSec(p.sec);
}
function bumpSec(sec) {
  if (!sec) return;
  const k = safeKey(sec);
  store.set(`stats/sections/${k}/sec`, String(sec)).then(() => store.inc(`stats/sections/${k}/n`)).catch(() => {});
}

/* ---------- 検索履歴の記録 ---------- */
async function recordHistory(type, q, hits) {
  q = String(q || "").trim();
  if (!q) return;
  const now = Date.now(), L = st.lastHist;
  try {
    if (L && L.type === type && (L.q === q || q.startsWith(L.q) || L.q.startsWith(q)) && now - L.at < 90 * 1000) {
      // 入力途中の連続記録はまとめる
      await store.update(`history/${L.key}`, { q, at: now, hits });
      st.lastHist = { ...L, q, at: now };
    } else if (L && L.type === type && L.q === q && now - L.at < 10 * 60 * 1000) {
      await store.update(`history/${L.key}`, { at: now, hits });
      st.lastHist = { ...L, at: now };
    } else {
      const key = await store.push("history", { type, q, hits, at: now });
      st.lastHist = { key, type, q, at: now };
      pruneHistory();
    }
  } catch (e) { console.warn(e); }
}
function pruneHistory() {
  const list = vals(st.history).sort((a, b) => b.at - a.at);
  if (list.length <= 1000) return;
  const del = {}; list.slice(1000).forEach((h) => (del[h.id] = null));
  store.update("history", del).catch(() => {});
}
const histTyping = debounce(() => { if (st.view === "sec" && st.q.trim()) recordHistory("part", st.q, st.results.length); }, 1500);
const histKata = debounce(() => { if (st.view === "kata" && S.normKata(st.kq).length >= 3) { const r = S.parseKata(st.kq, st.kata); recordHistory("kata", st.kq.trim(), r?.found ? 1 : 0); } }, 1500);

/* ============================================================
   型式 → 車種
   ============================================================ */
function renderKata() {
  $("#main").innerHTML = `<div class="view-head"><div><h1 class="h1">型式 → 車種</h1>
    <div class="tip">車検証の型式をそのまま入力（例：DAA-HFC27）／ ハイフンより前の記号は自動で無視 ／ 車種名を入れると逆引き</div></div></div>
    <p class="hello">&nbsp;</p>
    <label class="search">${IC.search}<input id="kq" type="search" autocomplete="off" spellcheck="false" placeholder="型式（例：DBA-C26、5AA-GFC27）または車種名" value="${esc(st.kq)}" data-live="1"><span class="cnt" id="kcnt"></span><button class="clear" id="kclear" aria-label="消す">×</button></label>
    <div id="kres"></div>`;
  const k = $("#kq");
  k.addEventListener("input", () => { st.kq = k.value; renderKataResult(); histKata(); });
  k.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); histKata.flush(); }
    if (e.key === "Escape") { e.preventDefault(); st.kq = ""; k.value = ""; renderKataResult(); }
  });
  $("#kclear").addEventListener("click", (e) => { e.preventDefault(); st.kq = ""; k.value = ""; renderKataResult(); k.focus(); });
  renderKataResult();
  setTimeout(() => k.focus(), 0);
}
const memosForKata = (code) => vals(st.memos).filter((m) => m.kataCode === code);
let lastCseQ = "";
function renderKataResult() {
  const box = $("#kres"), cnt = $("#kcnt"); if (!box) return;
  const raw = st.kq.trim();
  if (!raw) {
    cnt.textContent = ""; cnt.className = "cnt";
    box.innerHTML = `<div class="panel"><h3>使い方</h3><p style="margin:0;font-size:12.5px;color:var(--mu);line-height:1.8">
      ① 車検証の「型式」欄をそのまま入力 → 対応表にあれば車種名がすぐ出ます。<br>
      ② 対応表にない型式は、Google検索の結果を見て、その場で登録できます。<br>
      ③ 登録した型式は次回から即表示。使うほど対応表が育ちます。</p></div>`;
    return;
  }
  // 車種名で逆引き（日本語を含むとき）
  if (/[^\x00-\x7F]/.test(raw.normalize("NFKC"))) {
    const list = S.reverseKata(raw, st.kata);
    cnt.textContent = list.length ? `${list.length}件` : "見つかりません"; cnt.className = "cnt " + (list.length ? "ok" : "ng");
    box.innerHTML = list.length
      ? `<div class="panel"><h3>「${esc(raw)}」の型式</h3><div class="flex-wrap">${list.map((k) => `<button class="qchip" data-kq="${esc(k.code)}"><b>${esc(k.code)}</b>${esc(k.name)}${k.verified ? "" : " <small>未確認</small>"}</button>`).join("")}</div></div>`
      : `<div class="panel"><h3>対応表に「${esc(raw)}」はありません</h3><button class="btn" data-google="${esc(raw + " 型式")}">${IC.search} Googleで調べる</button></div>`;
    return;
  }
  const r = S.parseKata(raw, st.kata);
  const gq = `${setting("googlePrefix", "日産")} ${r.input}`.trim();
  const segs = `<div class="split">
      <div class="seg dim"><b>${esc(r.prefix || "—")}</b><small>規制などの記号<br>（判定に使わない）</small></div>
      <div class="seg"><b>${esc(r.mid || "—")}</b><small>エンジン・駆動などの<br>違い</small></div>
      <div class="seg key"><b>${esc(r.key || r.guess || "?")}</b><small>車種・世代${r.found ? "<br>→ " + esc(r.found.name) : "<br>（推定）"}</small></div></div>`;
  if (r.found) {
    const f = r.found, mc = memosForKata(r.key).length;
    cnt.textContent = "対応表にあり"; cnt.className = "cnt ok";
    box.innerHTML = `<div class="hit">
      <div class="car"><small>判定結果</small><div class="cn">${esc(f.name)}</div><div class="cg">${esc(r.key)}型</div>
        <span class="vtag ${f.verified ? "ok" : "ng"}">${f.verified ? "確認済み" : "未確認"}</span>
        ${f.note ? `<div style="font-size:11.5px;color:var(--mu);margin-top:8px;white-space:pre-line">${esc(f.note)}</div>` : ""}</div>
      ${segs}
      <div class="hacts"><button class="btn" data-google="${esc(gq)}">${IC.search} Googleで「${esc(gq)}」</button>
        <button class="btn ${mc ? "has" : ""}" data-katamemo="${esc(r.key)}">${IC.note} この車種のメモ${mc ? `<span class="cnt">${mc}</span>` : ""}</button>
        ${f.verified ? "" : `<button class="btn" data-verify="${esc(r.key)}">✓ 確認済みにする</button>`}
        <button class="btn" data-editkata="${esc(r.key)}">編集</button></div></div>`;
    return;
  }
  cnt.textContent = "対応表に未登録"; cnt.className = "cnt ng";
  const cx = setting("cseId", "");
  box.innerHTML = `<div class="hit" style="border-left-color:var(--ac2)"><div class="car"><small>判定結果</small><div class="cn" style="color:var(--ac2);font-size:24px">未登録</div><div class="cg">${esc(r.guess || "")}${r.guess ? "（推定）" : ""}</div></div>${segs}
      <div class="hacts"><button class="btn" data-google="${esc(gq)}">${IC.search} Googleを新しいタブで開く ${IC.ext}</button></div></div>
    <div class="unreg"><div class="ut"><span class="warn">未登録</span><b class="mono">${esc(r.input)}</b> を調べて、対応表に登録しましょう</div>
    <div class="ugrid">
      ${cx ? `<div class="gbox" id="gbox"></div>` : `<div class="gbox off"><p>Google検索の結果をここに表示するには、<br>「データ編集 → 設定 → Google検索の埋め込み」で<br>検索エンジンIDを登録してください。</p><button class="btn" data-google="${esc(gq)}">${IC.search} 新しいタブで検索</button></div>`}
      <form class="reg" id="kreg"><div class="rl">結果を見て、対応表に登録</div>
        <div class="field"><label>型式コード（車種・世代の部分）</label><input class="inp mono" name="code" required value="${esc(r.guess || "")}" pattern="[0-9A-Za-z]+" title="英数字のみ"></div>
        <div class="field"><label>車種名</label><input class="inp" name="name" required placeholder="例：ノート"></div>
        <label class="switch" style="margin:2px 0 8px"><input type="checkbox" name="verified"> 車検証・EPCで確認済み</label>
        <button class="btn pri" type="submit">登録する</button><small>次回からはすぐ車種名が表示されます</small></form>
    </div></div>`;
  $("#kreg").addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = e.target, code = S.normKata(f.code.value).replace(/[^0-9A-Z]/g, "");
    if (!code) return;
    if (st.kata[code] && !(await confirmBox(`${code} は「${st.kata[code].name}」で登録済みです。上書きしますか？`, "上書き"))) return;
    await saveKata(code, { name: f.name.value.trim(), verified: f.verified.checked, note: st.kata[code]?.note || "" });
    toast(`${code} を「${f.name.value.trim()}」で登録しました`);
  });
  if (cx) mountCse(cx, gq);
}
async function saveKata(code, data) {
  await store.set(`kata/${code}`, { name: data.name, verified: !!data.verified, note: data.note || "", updatedAt: Date.now() });
}

/* ---------- Google 検索の埋め込み（プログラム可能な検索エンジン） ---------- */
let csePromise = null, cseHost = null;
function loadCse(cx) {
  if (csePromise) return csePromise;
  csePromise = new Promise((res, rej) => {
    window.__gcse = { parsetags: "explicit", callback: res };
    const s = document.createElement("script");
    s.src = "https://cse.google.com/cse.js?cx=" + encodeURIComponent(cx);
    s.async = true; s.onerror = () => { csePromise = null; rej(new Error("Google検索の読み込みに失敗")); };
    document.head.appendChild(s);
  });
  return csePromise;
}
const runCse = debounce(async (cx, q) => {
  try {
    await loadCse(cx);
    const g = window.google?.search?.cse?.element;
    if (!g) throw new Error("検索エンジンを初期化できませんでした");
    if (!g.getElement("kata")) g.render({ div: "cse-target", tag: "searchresults-only", gname: "kata", attributes: { linkTarget: "_blank" } });
    if (q !== lastCseQ) { g.getElement("kata").execute(q); lastCseQ = q; }
  } catch (e) {
    const gb = $("#gbox"); if (gb) gb.innerHTML = `<p style="color:#b33">${esc(e.message)}（新しいタブで検索してください）</p>`;
  }
}, 700);
function mountCse(cx, q) {
  if (!cseHost) { cseHost = document.createElement("div"); cseHost.innerHTML = '<div id="cse-target"></div>'; }
  $("#gbox").appendChild(cseHost);
  runCse(cx, q);
}

/* ============================================================
   メモ
   ============================================================ */
function memoMatches(m) {
  if (st.memoFilter === "open" && m.done) return false;
  if (st.memoFilter === "done" && !m.done) return false;
  if (st.memoPart && m.partId !== st.memoPart) return false;
  if (st.memoKata && m.kataCode !== st.memoKata) return false;
  const q = S.norm(st.memoQ);
  if (q && !S.norm(`${m.title} ${m.body} ${m.partLabel || ""} ${m.kataCode || ""} ${m.kataName || ""}`).includes(q)) return false;
  return true;
}
function renderMemo(full) {
  const main = $("#main");
  if (full || !$("#memo-list")) {
    main.innerHTML = `<div class="view-head"><div><h1 class="h1">メモ</h1><div class="tip">業務中の疑問や気づきを書き留める ／ 部品・型式に紐付けられます ／ 書いた内容は自動で保存</div></div>
      <button class="btn pri" data-newmemo="1">＋ 新しいメモ</button></div>
      <div class="memo-grid"><div><div class="toolbar"><input class="inp" id="mq" placeholder="メモを検索" value="${esc(st.memoQ)}" data-live="1"></div>
        <div class="chips" id="mchips" style="margin-top:0"></div><div class="memo-list" id="memo-list"></div></div>
        <div id="memo-ed"></div></div>`;
    $("#mq").addEventListener("input", (e) => { st.memoQ = e.target.value; renderMemoList(); });
  }
  renderMemoList();
  const ed = $("#memo-ed");
  const typing = ed.contains(document.activeElement) && ed.dataset.id === st.memoSel;
  if (!typing) renderMemoEditor();
}
function renderMemoList() {
  const all = vals(st.memos);
  const c = { open: all.filter((m) => !m.done).length, done: all.filter((m) => m.done).length, all: all.length };
  $("#mchips").innerHTML = [["open", "未解決"], ["done", "解決済み"], ["all", "すべて"]].map(([k, l]) => `<button class="chip ${st.memoFilter === k ? "on" : ""}" data-mf="${k}">${l} ${c[k]}</button>`).join("") +
    (st.memoPart ? `<button class="chip secf" data-mpoff="1">部品で絞り込み中 ×</button>` : "") + (st.memoKata ? `<button class="chip secf" data-mkoff="1">${esc(st.memoKata)} ×</button>` : "");
  const list = all.filter(memoMatches).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  $("#memo-list").innerHTML = list.map((m) => `<button class="mitem ${m.id === st.memoSel ? "on" : ""} ${m.done ? "done" : ""}" data-memo="${esc(m.id)}">
      <div class="mt"><span class="dot"></span><span>${esc(m.title || "（無題）")}</span></div>
      <div class="mm"><span>${fmtWhen(m.updatedAt)}</span>${m.partLabel ? `<span>🔧 ${esc(m.partLabel)}</span>` : ""}${m.kataCode ? `<span>🚗 ${esc(m.kataCode)} ${esc(m.kataName || "")}</span>` : ""}</div></button>`).join("") ||
    `<p class="empty">${st.memoFilter === "open" ? "未解決のメモはありません 👍" : "メモはありません"}</p>`;
}
function renderMemoEditor() {
  const ed = $("#memo-ed"); if (!ed) return;
  const m = st.memoSel && st.memos[st.memoSel] ? { id: st.memoSel, ...st.memos[st.memoSel] } : null;
  ed.dataset.id = m ? m.id : "";
  if (!m) { ed.innerHTML = `<div class="editor"><p class="empty" style="margin:0">左のメモを選ぶか、「＋ 新しいメモ」で書き始めてください。</p></div>`; return; }
  const part = m.partId && st.parts[m.partId] ? { id: m.partId, ...st.parts[m.partId] } : null;
  ed.innerHTML = `<div class="editor">
    <input class="inp title" id="me-title" placeholder="タイトル（例：C27のオイルエレメント、寒冷地仕様は？）" value="${esc(m.title)}">
    <div class="links">紐付け：
      ${m.partId ? `<button class="ltag" data-q="${esc(part ? part.name : "")}" title="セクション検索で開く">🔧 <b>${esc(part ? part.sec : "")}</b>${esc(part ? `${part.code} ${part.name}` : m.partLabel || "（削除された部品）")}</button><button class="x" data-unlink="part" title="紐付けを外す">×</button>` : ""}
      ${m.kataCode ? `<button class="ltag" data-kq="${esc(m.kataCode)}">🚗 <b>${esc(m.kataCode)}</b>${esc(st.kata[m.kataCode]?.name || m.kataName || "")}</button><button class="x" data-unlink="kata" title="紐付けを外す">×</button>` : ""}
      ${!m.partId && !m.kataCode ? '<span style="color:var(--mu2)">なし（部品の「メモ」ボタンや型式画面から作ると自動で紐付きます）</span>' : ""}
    </div>
    <textarea class="inp" id="me-body" rows="14" placeholder="疑問点・調べたこと・答え など">${esc(m.body)}</textarea>
    <div class="ed-foot"><span class="meta" id="me-meta">作成 ${fmtWhen(m.createdAt)} ／ 更新 ${fmtWhen(m.updatedAt)}</span>
      <label class="switch"><input type="checkbox" id="me-done" ${m.done ? "checked" : ""}> 解決済み</label>
      <button class="btn danger sm" data-delmemo="${esc(m.id)}">削除</button></div></div>`;
  const save = debounce(async (fields) => {
    try { await store.update(`memos/${m.id}`, { ...fields, updatedAt: Date.now() }); $("#me-meta") && ($("#me-meta").textContent = "保存しました ✓"); }
    catch (e) { toast(e.message, true); }
  }, 600);
  const t = $("#me-title"), b = $("#me-body");
  const fire = () => { $("#me-meta").textContent = "保存中…"; save({ title: t.value, body: b.value }); };
  t.addEventListener("input", fire); b.addEventListener("input", fire);
  t.addEventListener("blur", () => save.flush({ title: t.value, body: b.value }));
  b.addEventListener("blur", () => save.flush({ title: t.value, body: b.value }));
  $("#me-done").addEventListener("change", async (e) => {
    await store.update(`memos/${m.id}`, { done: e.target.checked, updatedAt: Date.now() });
    toast(e.target.checked ? "解決済みにしました" : "未解決に戻しました");
  });
}
async function newMemo(fields = {}) {
  const now = Date.now();
  const id = await store.push("memos", { title: "", body: "", done: false, createdAt: now, updatedAt: now, ...fields });
  st.memoSel = id; st.memoFilter = fields.done ? "all" : "open"; st.memoQ = "";
  if (st.view !== "memo") go("memo"); else renderMemo(true);
  setTimeout(() => { const t = $("#me-title"); if (t) { t.focus(); t.setSelectionRange(t.value.length, t.value.length); } }, 50);
}
async function openMemoForPart(p) {
  histTyping.flush();
  const ms = memosForPart(p.id).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  if (!ms.length) return newMemo({ title: `${p.name}について`, partId: p.id, partLabel: `${p.sec} ${p.code} ${p.name}` });
  st.memoSel = ms[0].id; st.memoFilter = "all"; st.memoPart = ms.length > 1 ? p.id : null; st.memoKata = null; st.memoQ = "";
  go("memo");
}
async function openMemoForKata(code) {
  histKata.flush();
  const ms = memosForKata(code).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  const name = st.kata[code]?.name || "";
  if (!ms.length) return newMemo({ title: `${code}（${name}）について`, kataCode: code, kataName: name });
  st.memoSel = ms[0].id; st.memoFilter = "all"; st.memoKata = ms.length > 1 ? code : null; st.memoPart = null; st.memoQ = "";
  go("memo");
}

/* ============================================================
   検索履歴
   ============================================================ */
function renderHist() {
  const main = $("#main");
  if (!$("#hlist")) {
    main.innerHTML = `<div class="view-head"><div><h1 class="h1">検索履歴</h1><div class="tip">いつ・何を調べたか ／ 押すともう一度検索します</div></div>
      <button class="btn danger sm" id="hclear">履歴をすべて消す</button></div>
      <div class="toolbar"><input class="inp" id="hq" placeholder="履歴を検索" value="${esc(st.histQ)}" data-live="1"><span class="sp"></span><div class="chips" id="hchips" style="margin:0"></div></div>
      <div id="hlist"></div>`;
    $("#hq").addEventListener("input", (e) => { st.histQ = e.target.value; renderHist(); });
    $("#hclear").addEventListener("click", async () => {
      if (!(await confirmBox("検索履歴をすべて削除します。元に戻せません。", "すべて削除", true))) return;
      await store.remove("history"); st.lastHist = null; toast("履歴を削除しました");
    });
  }
  $("#hchips").innerHTML = [["all", "すべて"], ["part", "部品"], ["kata", "型式"]].map(([k, l]) => `<button class="chip ${st.histFilter === k ? "on" : ""}" data-hf="${k}">${l}</button>`).join("");
  const q = S.norm(st.histQ);
  const list = vals(st.history)
    .filter((h) => (st.histFilter === "all" || (h.type || "part") === st.histFilter) && (!q || S.norm(h.q).includes(q)))
    .sort((a, b) => b.at - a.at).slice(0, 400);
  let html = "", day = "";
  for (const h of list) {
    const d = dayKey(h.at);
    if (d !== day) { day = d; html += `<div class="hday">${fmtDay(h.at)}</div>`; }
    const k = h.type === "kata";
    html += `<div class="hrow"><time>${fmtHM(h.at)}</time><span class="htype ${k ? "k" : ""}">${k ? "型式" : "部品"}</span>
      <button class="hq" data-rerun="${esc(h.id)}" title="もう一度検索">${esc(h.q)}</button><span class="hhits">${k ? (h.hits ? "登録あり" : "未登録") : `${h.hits ?? "-"}件`}</span>
      <button class="x" data-delhist="${esc(h.id)}" title="この履歴を消す">×</button></div>`;
  }
  $("#hlist").innerHTML = html || '<p class="empty">履歴はありません</p>';
}
function rerun(id) {
  const h = st.history[id]; if (!h) return;
  if (h.type === "kata") { st.kq = h.q; go("kata"); }
  else { st.q = h.q; st.sec = ""; st.group = "all"; st.sel = 0; go("sec"); }
}

/* ============================================================
   データ編集
   ============================================================ */
const ETABS = [["parts", "部品データ"], ["kata", "型式対応表"], ["io", "取り込み・書き出し"], ["settings", "設定"]];
function renderEdit() {
  $("#main").innerHTML = `<div class="view-head"><div><h1 class="h1">データ編集</h1><div class="tip">部品データ・型式対応表の追加・修正・削除、取り込み、各種設定</div></div></div>
    <div class="subtabs">${ETABS.map(([k, l]) => `<button class="subtab ${st.editTab === k ? "on" : ""}" data-etab="${k}">${l}</button>`).join("")}</div><div id="ebody"></div>`;
  ({ parts: editParts, kata: editKata, io: editIO, settings: editSettings })[st.editTab]();
}
function editParts() {
  $("#ebody").innerHTML = `<div class="toolbar"><input class="inp" id="eq" placeholder="部品名・コード・セクションで絞り込み" value="${esc(st.editQ)}" data-live="1"><span class="stat" id="estat"></span><span class="sp"></span><button class="btn pri" data-addpart="">＋ 部品を追加</button></div><div id="etbl"></div>`;
  $("#eq").addEventListener("input", (e) => { st.editQ = e.target.value; editPartsTable(); });
  editPartsTable();
}
function editPartsTable() {
  const rs = st.editQ.trim() ? S.searchParts(st.partList, st.editQ, st.syn).map((r) => r.p)
    : [...st.partList].sort((a, b) => S.cmpSec(a.sec, b.sec) || String(a.code).localeCompare(String(b.code)));
  $("#estat").textContent = `${rs.length} / ${st.partList.length}件`;
  $("#etbl").innerHTML = `<table class="tbl"><thead><tr><th>SEC</th><th>部品コード</th><th>部品名</th><th>備考</th><th>頻出</th><th></th></tr></thead><tbody>${rs.slice(0, 600).map((p) => `<tr>
    <td class="c-sec">${esc(p.sec)}</td><td class="c-code">${esc(p.code)}</td><td>${esc(p.name)}</td><td class="c-note">${esc(p.note)}</td><td>${p.hot ? `<span title="${esc(p.hot)}" style="color:var(--ac2)">★</span>` : ""}</td>
    <td class="c-act"><button class="btn sm" data-editpart="${esc(p.id)}">編集</button> <button class="btn sm danger" data-delpart="${esc(p.id)}">削除</button></td></tr>`).join("")}</tbody></table>` +
    (rs.length > 600 ? '<p class="stat">先頭600件を表示中。絞り込んでください。</p>' : "") + (rs.length ? "" : '<p class="empty">該当なし</p>');
}
async function partForm(p = {}, title = "部品を追加") {
  const r = await modal(`<h3>${esc(title)}</h3><form>
    <div class="grid2"><div class="field"><label>セクション</label><input class="inp mono" name="sec" required value="${esc(p.sec)}" placeholder="例：150"></div>
    <div class="field"><label>部品コード</label><input class="inp mono" name="code" required value="${esc(p.code)}" placeholder="例：15208"></div></div>
    <div class="field"><label>部品名（検索される名前。言い換えをスペース区切りで足してもOK）</label><input class="inp" name="name" required value="${esc(p.name)}" ${p.id ? "" : "autofocus"}></div>
    <div class="field"><label>備考（注意点など）</label><textarea class="inp" name="note" rows="3">${esc(p.note)}</textarea></div>
    <div class="field"><label>頻出コメント（入れると★頻出になります）</label><input class="inp" name="hot" value="${esc(p.hot)}" placeholder="例：車検の主役"></div>
    </form><div class="modal-foot">${p.id ? `<button class="btn danger left" data-m="delete">削除</button>` : ""}<button class="btn" data-m="cancel">キャンセル</button><button class="btn pri" data-m="ok">保存</button></div>`);
  if (!r) return;
  if (r.act === "delete") return delPart(p.id);
  const d = r.data, norm = (s) => String(s || "").normalize("NFKC").trim();
  const rec = { sec: norm(d.sec), code: norm(d.code), name: norm(d.name).replace(/\s+/g, " "), note: norm(d.note), updatedAt: Date.now() };
  if (norm(d.hot)) rec.hot = norm(d.hot);
  try {
    if (p.id) await store.set(`parts/${p.id}`, rec); else await store.push("parts", rec);
    toast(p.id ? "保存しました" : `「${rec.name}」を追加しました`);
  } catch (e) { toast(e.message, true); }
}
async function delPart(id) {
  const p = st.parts[id]; if (!p) return;
  if (!(await confirmBox(`「${p.name}」（SEC ${p.sec} / ${p.code}）を削除しますか？`, "削除", true))) return;
  await store.remove(`parts/${id}`); toast("削除しました");
}
function editKata() {
  $("#ebody").innerHTML = `<div class="toolbar"><input class="inp" id="kfq" placeholder="型式・車種名で絞り込み" value="${esc(st.kataQ)}" data-live="1"><span class="stat" id="kstat"></span><span class="sp"></span><button class="btn pri" data-editkata="">＋ 型式を追加</button></div><div id="ktbl"></div>`;
  $("#kfq").addEventListener("input", (e) => { st.kataQ = e.target.value; editKataTable(); });
  editKataTable();
}
function editKataTable() {
  const q = S.norm(st.kataQ);
  const list = vals(st.kata).filter((k) => !q || S.norm(k.id + k.name).includes(q)).sort((a, b) => a.name.localeCompare(b.name, "ja") || a.id.localeCompare(b.id));
  const unv = vals(st.kata).filter((k) => !k.verified).length;
  $("#kstat").textContent = `${list.length}件${unv ? `（未確認 ${unv}件）` : ""}`;
  $("#ktbl").innerHTML = `<table class="tbl"><thead><tr><th>型式コード</th><th>車種名</th><th>確認</th><th>メモ</th><th>更新</th><th></th></tr></thead><tbody>${list.map((k) => `<tr>
    <td class="c-code" style="color:var(--ac)">${esc(k.id)}</td><td>${esc(k.name)}</td>
    <td><label class="switch"><input type="checkbox" data-kver="${esc(k.id)}" ${k.verified ? "checked" : ""}> ${k.verified ? "確認済み" : '<span style="color:var(--ac2)">未確認</span>'}</label></td>
    <td class="c-note">${esc(k.note)}</td><td class="stat">${fmtWhen(k.updatedAt)}</td>
    <td class="c-act"><button class="btn sm" data-editkata="${esc(k.id)}">編集</button> <button class="btn sm danger" data-delkata="${esc(k.id)}">削除</button></td></tr>`).join("")}</tbody></table>${list.length ? "" : '<p class="empty">登録がありません。「取り込み・書き出し」から下書き（kata_seed.json）も入れられます。</p>'}`;
}
async function kataForm(code) {
  const k = code ? { id: code, ...st.kata[code] } : {};
  const r = await modal(`<h3>${code ? "型式を編集" : "型式を追加"}</h3><form>
    <div class="grid2"><div class="field"><label>型式コード（例：C27）</label><input class="inp mono" name="code" required pattern="[0-9A-Za-z]+" value="${esc(k.id)}" ${code ? "readonly" : "autofocus"}></div>
    <div class="field"><label>車種名</label><input class="inp" name="name" required value="${esc(k.name)}" ${code ? "autofocus" : ""}></div></div>
    <div class="field"><label>メモ（年式・仕様など）</label><textarea class="inp" name="note" rows="2">${esc(k.note)}</textarea></div>
    <label class="switch"><input type="checkbox" name="verified" ${k.verified ? "checked" : ""}> 車検証・EPCで確認済み</label></form>
    <div class="modal-foot">${code ? '<button class="btn danger left" data-m="delete">削除</button>' : ""}<button class="btn" data-m="cancel">キャンセル</button><button class="btn pri" data-m="ok">保存</button></div>`);
  if (!r) return;
  if (r.act === "delete") return delKata(code);
  const c = S.normKata(r.data.code).replace(/[^0-9A-Z]/g, "");
  if (!code && st.kata[c] && !(await confirmBox(`${c} は登録済みです（${st.kata[c].name}）。上書きしますか？`, "上書き"))) return;
  await saveKata(c, { name: r.data.name.trim(), note: (r.data.note || "").trim(), verified: r.data.verified === "on" });
  toast("保存しました");
}
async function delKata(code) {
  if (!(await confirmBox(`${code}（${st.kata[code]?.name}）を削除しますか？`, "削除", true))) return;
  await store.remove(`kata/${code}`); toast("削除しました");
}

/* ---------- 取り込み・書き出し ---------- */
function editIO() {
  $("#ebody").innerHTML = `<div class="card-grid">
    <div class="card"><h4>部品データの取り込み</h4><p>parts.json（変換済みデータ）か、Excelの「部品マスタ」シート（.xlsx）、CSVを読み込みます。<br>列名「セクション」「部品コード」「部品正式名称（部品名）」を自動で探します。「***」以降は備考になります。</p>
      <div class="radio"><label><input type="radio" name="pmode" value="add" checked> 追加（同じものは飛ばす）</label><label><input type="radio" name="pmode" value="replace"> 全部置き換え</label></div>
      <div class="row2"><input type="file" id="fparts" accept=".json,.xlsx,.xlsm,.xls,.csv" hidden><button class="btn pri" data-pick="fparts">ファイルを選ぶ</button><span class="stat">現在 ${st.partList.length}件</span></div></div>
    <div class="card"><h4>型式対応表の取り込み</h4><p>kata_seed.json（下書き）や、「型式コード,車種名」のCSVを読み込みます。登録済みの型式は上書きしません。</p>
      <div class="row2"><input type="file" id="fkata" accept=".json,.csv" hidden><button class="btn pri" data-pick="fkata">ファイルを選ぶ</button><span class="stat">現在 ${Object.keys(st.kata).length}件</span></div></div>
    <div class="card"><h4>バックアップ</h4><p>部品データ・型式対応表・メモ・設定をまとめて1つのファイルに保存します。定期的に取っておくと安心です。（検索履歴は含みません）</p>
      <div class="row2"><button class="btn" id="bk-save">バックアップを保存</button><input type="file" id="fbk" accept=".json" hidden><button class="btn" data-pick="fbk">バックアップから復元</button></div></div>
    <div class="card"><h4>Excel用に書き出し</h4><p>部品データをCSV（Excelで開ける形式）で保存します。</p><div class="row2"><button class="btn" id="csv-save">CSVで保存</button></div></div>
  </div>`;
  $("#fparts").addEventListener("change", (e) => importParts(e.target.files[0], $('input[name="pmode"]:checked').value).finally(() => (e.target.value = "")));
  $("#fkata").addEventListener("change", (e) => importKata(e.target.files[0]).finally(() => (e.target.value = "")));
  $("#fbk").addEventListener("change", (e) => restoreBackup(e.target.files[0]).finally(() => (e.target.value = "")));
  $("#bk-save").addEventListener("click", saveBackup);
  $("#csv-save").addEventListener("click", saveCsv);
}
function download(name, text, type = "application/json") {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([text], { type }));
  a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
const stamp = () => { const d = new Date(); return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`; };
function strip(o) { const r = {}; for (const [k, v] of Object.entries(o || {})) { const { id, _name, _note, _code, _sec, ...rest } = v; r[k] = rest; } return r; }
function saveBackup() {
  download(`section-search-backup-${stamp()}.json`, JSON.stringify({ format: "section-search/backup@1", exportedAt: new Date().toISOString(), parts: strip(st.parts), kata: st.kata, memos: st.memos, settings: st.settings }, null, 1));
  toast("バックアップを保存しました");
}
function saveCsv() {
  const q = (s) => `"${String(s ?? "").replace(/"/g, '""')}"`;
  const rows = [...st.partList].sort((a, b) => S.cmpSec(a.sec, b.sec)).map((p) => [p.sec, p.code, p.name, p.note, p.hot || ""].map(q).join(","));
  download(`parts-${stamp()}.csv`, "﻿" + ["セクション,部品コード,部品名,備考,頻出"].concat(rows).join("\r\n"), "text/csv");
}
async function restoreBackup(file) {
  if (!file) return;
  try {
    const j = JSON.parse(await file.text());
    if (j.format !== "section-search/backup@1") throw new Error("このアプリのバックアップファイルではありません");
    const n = Object.keys(j.parts || {}).length, k = Object.keys(j.kata || {}).length, m = Object.keys(j.memos || {}).length;
    if (!(await confirmBox(`バックアップ（${j.exportedAt || "日時不明"}）で、今のデータを置き換えます。\n部品 ${n}件 ／ 型式 ${k}件 ／ メモ ${m}件\n\n今のデータは消えます。よろしいですか？`, "復元する", true))) return;
    await store.update("", { parts: j.parts || null, kata: j.kata || null, memos: j.memos || null, settings: j.settings || null });
    toast("復元しました");
  } catch (e) { toast("復元できませんでした：" + e.message, true); }
}

// ファイル → 部品レコードの配列
function cleanPart(sec, code, raw, noteCol = "", hot = "") {
  const N = (s) => String(s ?? "").normalize("NFKC").trim();
  if (typeof sec === "number" && !Number.isInteger(sec)) sec = String(+sec.toFixed(4));
  let name = N(raw), note = "";
  if (!name) return null;
  if (name.includes("***")) [name, note] = [name.slice(0, name.indexOf("***")), name.slice(name.indexOf("***") + 3)];
  else if (name.includes("☞")) [name, note] = [name.slice(0, name.indexOf("☞")), name.slice(name.indexOf("☞") + 1)];
  const lines = name.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  name = (lines.shift() || "").replace(/\s+/g, " ").trim();
  note = [lines.join("\n"), note.trim(), N(noteCol)].filter(Boolean).join("\n");
  if (!name) return null;
  const p = { sec: N(sec), code: N(code), name, note };
  if (N(hot)) p.hot = N(hot);
  return p;
}
function rowsToParts(rows) {
  // 見出し行を探す
  for (let h = 0; h < Math.min(rows.length, 15); h++) {
    const hd = (rows[h] || []).map((c) => String(c ?? "").normalize("NFKC").replace(/\s/g, ""));
    const iSec = hd.findIndex((c) => /^(セクション|SEC|sec|Sec)$/.test(c) || c === "セクション番号");
    const iCode = hd.findIndex((c) => c.includes("部品コード") || c === "コード");
    const iName = hd.findIndex((c, i) => i !== iCode && (c.includes("正式名称") || c === "部品名" || c.startsWith("部品名")));
    if (iSec < 0 || iCode < 0 || iName < 0) continue;
    const iNote = hd.findIndex((c) => c === "備考");
    const iHot = hd.findIndex((c) => c === "頻出");
    const out = [];
    for (const r of rows.slice(h + 1)) {
      const p = cleanPart(r[iSec], r[iCode], r[iName], iNote >= 0 ? r[iNote] : "", iHot >= 0 ? r[iHot] : "");
      if (p && p.sec && p.code) out.push(p);
    }
    return out;
  }
  throw new Error("「セクション」「部品コード」「部品名」の列が見つかりませんでした");
}
function parseCsv(text) {
  const rows = []; let row = [], cur = "", q = false;
  text = text.replace(/^﻿/, "");
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
    else if (c === '"') q = true;
    else if (c === ",") { row.push(cur); cur = ""; }
    else if (c === "\n" || c === "\r") { if (c === "\r" && text[i + 1] === "\n") i++; row.push(cur); rows.push(row); row = []; cur = ""; }
    else cur += c;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}
let xlsxPromise = null;
function loadXlsx() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (!xlsxPromise) xlsxPromise = new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = () => res(window.XLSX); s.onerror = () => { xlsxPromise = null; rej(new Error("Excel読み込み部品のダウンロードに失敗しました")); };
    document.head.appendChild(s);
  });
  return xlsxPromise;
}
async function readPartsFile(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  if (ext === "json") {
    const j = JSON.parse(await file.text());
    if (j.format === "section-search/backup@1") throw new Error("バックアップファイルは「バックアップから復元」で読み込んでください");
    const arr = Array.isArray(j) ? j : j.parts;
    if (!Array.isArray(arr)) throw new Error("部品データの形式ではありません");
    return arr.map((p) => cleanPart(p.sec, p.code, p.name, p.note, p.hot)).filter((p) => p && p.sec && p.code);
  }
  if (ext === "csv") return rowsToParts(parseCsv(await file.text()));
  const X = await loadXlsx();
  const wb = X.read(await file.arrayBuffer(), { type: "array" });
  const names = [...wb.SheetNames].sort((a, b) => (b === "部品マスタ") - (a === "部品マスタ"));
  let lastErr;
  for (const n of names) {
    try { const rows = X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" }); const ps = rowsToParts(rows); if (ps.length) return ps; }
    catch (e) { lastErr = e; }
  }
  throw lastErr || new Error("部品データが見つかりませんでした");
}
async function importParts(file, mode) {
  if (!file) return;
  try {
    const ps = await readPartsFile(file);
    if (!ps.length) throw new Error("取り込める行がありませんでした");
    const sig = (p) => `${S.norm(p.sec)}|${S.norm(p.code)}|${S.norm(p.name)}`;
    const have = new Set(mode === "add" ? st.partList.map(sig) : []);
    const seen = new Set(); const fresh = [];
    for (const p of ps) { const s = sig(p); if (have.has(s) || seen.has(s)) continue; seen.add(s); fresh.push(p); }
    const msg = mode === "replace"
      ? `今の部品データ ${st.partList.length}件 を消して、${fresh.length}件 に置き換えます。\n（メモの紐付けは外れます）`
      : `${ps.length}件 のうち、新しい ${fresh.length}件 を追加します。${ps.length - fresh.length ? `\n（同じものが ${ps.length - fresh.length}件 あったので飛ばします）` : ""}`;
    if (!fresh.length) { toast("新しく追加する部品はありませんでした"); return; }
    if (!(await confirmBox(msg, mode === "replace" ? "置き換える" : "追加する", mode === "replace"))) return;
    const obj = {}; const now = Date.now();
    for (const p of fresh) obj[localKey()] = { ...p, updatedAt: now };
    if (mode === "replace") await store.set("parts", obj); else await store.update("parts", obj);
    toast(`${fresh.length}件 取り込みました`);
  } catch (e) { toast("取り込めませんでした：" + e.message, true); }
}
async function importKata(file) {
  if (!file) return;
  try {
    let list;
    if (file.name.toLowerCase().endsWith(".json")) {
      const j = JSON.parse(await file.text()); list = Array.isArray(j) ? j : j.kata;
    } else {
      list = parseCsv(await file.text()).filter((r) => r.length >= 2).map((r) => ({ code: r[0], name: r[1], note: r[2] || "" }));
    }
    if (!Array.isArray(list)) throw new Error("型式データの形式ではありません");
    const obj = {}; let skip = 0; const now = Date.now();
    for (const k of list) {
      const code = S.normKata(k.code).replace(/[^0-9A-Z]/g, ""), name = String(k.name || "").normalize("NFKC").trim();
      if (!code || !name || /^(型式|CODE)/i.test(code)) continue;
      if (st.kata[code] || obj[code]) { skip++; continue; }
      obj[code] = { name, verified: !!k.verified, note: String(k.note || ""), updatedAt: now };
    }
    const n = Object.keys(obj).length;
    if (!n) { toast(`新しい型式はありませんでした（登録済み ${skip}件）`); return; }
    if (!(await confirmBox(`${n}件 の型式を追加します。${skip ? `\n（登録済みの ${skip}件 は飛ばします）` : ""}${list.some((k) => !k.verified) ? "\n\n※「未確認」として入ります。確かめたら確認済みにしてください。" : ""}`, "追加する"))) return;
    await store.update("kata", obj); toast(`${n}件 追加しました`);
  } catch (e) { toast("取り込めませんでした：" + e.message, true); }
}

/* ---------- 設定 ---------- */
async function editSettings() {
  const reg = FID.registered(), av = await FID.platformAvailable();
  $("#ebody").innerHTML = `<div class="card-grid">
    <div class="card"><h4>生体認証ロック（Face ID / Windows Hello）</h4>
      <p>この端末の生体認証（顔・指紋・PIN）でアプリのロックを解除します。端末ごとに設定します。<br>${reg ? `<b style="color:var(--ok)">この端末は設定済み</b>（${fmtDay(reg.at)}）` : av ? "この端末で使えます。" : '<span style="color:var(--ac2)">この端末・ブラウザでは使えません（httpsで開いているか確認してください）</span>'}</p>
      <div class="row2">${reg ? `<button class="btn" id="fid-test">ロックを試す</button><button class="btn danger" id="fid-off">解除する</button>` : `<button class="btn pri" id="fid-on" ${av ? "" : "disabled"}>この端末に設定する</button>`}</div>
      ${reg ? `<div class="field" style="margin-top:12px"><label>自動ロックまでの時間（操作なし・画面を離れていたとき）</label><select class="inp" id="fid-min" style="max-width:200px">${[5, 15, 30, 60, 180].map((m) => `<option value="${m}" ${FID.settings().lockMin === m ? "selected" : ""}>${m}分</option>`).join("")}</select></div>` : ""}</div>
    <div class="card"><h4>Google検索</h4><p>部品名の前に付ける言葉です（例：日産 → 「日産 オイルエレメント」で検索）。</p>
      <div class="row2"><input class="inp" id="g-prefix" style="max-width:200px" value="${esc(setting("googlePrefix", "日産"))}"><button class="btn" id="g-prefix-save">保存</button></div></div>
    <div class="card"><h4>Google検索の埋め込み（型式画面）</h4><p>Googleの「プログラム可能な検索エンジン」で作った<b>検索エンジンID</b>を入れると、未登録の型式を調べるとき結果がアプリ内に表示されます。空なら新しいタブで開きます。</p>
      <div class="row2"><input class="inp mono" id="cse-id" placeholder="例：a1b2c3d4e5f6g7h8i" value="${esc(setting("cseId", ""))}" style="max-width:260px"><button class="btn" id="cse-save">保存</button></div></div>
    <div class="card" style="grid-column:1/-1"><h4>言い換えリスト</h4><p>1行に1組、「=」でつなぎます。片方で検索すると、もう片方の名前もヒットします。（例：「オイルフィルター」で「オイルエレメント」が出る）</p>
      <textarea class="inp mono" id="syn" rows="8">${esc(setting("synonyms", S.DEFAULT_SYNONYMS))}</textarea>
      <div class="row2" style="margin-top:8px"><button class="btn pri" id="syn-save">保存</button><button class="btn" id="syn-reset">初期値に戻す</button></div></div>
    <div class="card"><h4>アカウント</h4><p>ログイン中：<b>${esc(st.user?.email)}</b>${DEMO ? "<br>デモモードで動いています。js/config.js にFirebaseの設定を入れると本番モードになります。" : ""}</p>
      <div class="row2">${DEMO ? "" : '<button class="btn" id="pw-reset">パスワード再設定メールを送る</button>'}<button class="btn danger" id="logout2">ログアウト</button></div></div>
  </div>`;
  const sset = async (k, v, msg = "保存しました") => { try { await store.set(`settings/${k}`, v); toast(msg); } catch (e) { toast(e.message, true); } };
  $("#fid-on")?.addEventListener("click", async () => {
    try { await FID.register(st.user?.email || "セクション検索"); toast("生体認証ロックを設定しました"); renderFid(); editSettings(); }
    catch (e) { toast(e.name === "NotAllowedError" ? "キャンセルされました" : "設定できませんでした：" + e.message, true); }
  });
  $("#fid-off")?.addEventListener("click", async () => { if (await confirmBox("この端末の生体認証ロックを解除しますか？", "解除する")) { FID.unregister(); renderFid(); editSettings(); toast("解除しました"); } });
  $("#fid-test")?.addEventListener("click", lockNow);
  $("#fid-min")?.addEventListener("change", (e) => { FID.setLockMin(+e.target.value); renderFid(); toast("保存しました"); });
  $("#g-prefix-save").addEventListener("click", () => sset("googlePrefix", $("#g-prefix").value.trim()));
  $("#cse-save").addEventListener("click", () => { const v = $("#cse-id").value.trim(); if (v && !/^[\w:-]+$/.test(v)) return toast("IDの形式が正しくありません", true); sset("cseId", v); if (csePromise) toast("反映するには一度ページを再読み込みしてください"); });
  $("#syn-save").addEventListener("click", () => sset("synonyms", $("#syn").value));
  $("#syn-reset").addEventListener("click", () => { $("#syn").value = S.DEFAULT_SYNONYMS; });
  $("#pw-reset")?.addEventListener("click", async () => { try { await store.resetPassword(st.user.email); toast("再設定メールを送りました"); } catch (e) { toast(e.message, true); } });
  $("#logout2").addEventListener("click", async () => { if (await confirmBox("ログアウトしますか？", "ログアウト")) store.logout(); });
}

/* ============================================================
   クリック（まとめて受け取る）
   ============================================================ */
document.addEventListener("click", async (e) => {
  const t = e.target.closest("button,[data-pid]");
  if (!t || !st.user || st.locked) return;
  const d = t.dataset;
  const row = t.closest(".row");
  if (d.view) return go(d.view);
  if (d.go) { const [v, tab] = d.go.split(":"); return go(v, { editTab: tab }); }
  if (d.act && row) {
    const p = st.results.find((r) => r.p.id === row.dataset.pid)?.p; if (!p) return;
    st.sel = +row.dataset.i; markSel();
    return d.act === "google" ? googlePart(p) : openMemoForPart(p);
  }
  if (d.sec !== undefined) { histTyping.flush(); st.sec = d.sec; st.q = ""; st.group = "all"; st.sel = 0; if (st.view !== "sec") go("sec"); else { $("#q").value = ""; renderResults(); } bumpSec(d.sec); return; }
  if (d.secoff) { st.sec = ""; return renderResults(); }
  if (d.group) { st.group = d.group; st.sel = 0; st.limit = 60; renderResults(); return $("#q")?.focus(); }
  if (d.q !== undefined) { st.q = d.q; st.sec = ""; st.group = "all"; st.sel = 0; if (st.view !== "sec") go("sec"); else { $("#q").value = d.q; renderResults(); $("#q").focus(); } histTyping(); return; }
  if (d.more) { st.limit += 100; return renderResults(); }
  if (d.google) { googleOpen(d.google); if (st.view === "kata") histKata.flush(); return; }
  if (d.addpart !== undefined) { const isCode = /^[0-9A-Za-z]{4,}$/.test(d.addpart); return partForm(isCode ? { code: d.addpart } : { name: d.addpart }); }
  if (d.editpart) return partForm({ id: d.editpart, ...st.parts[d.editpart] }, "部品を編集");
  if (d.delpart) return delPart(d.delpart);
  if (d.kq !== undefined) { st.kq = d.kq; if (st.view !== "kata") go("kata"); else { $("#kq").value = d.kq; renderKataResult(); } return; }
  if (d.verify) { const k = st.kata[d.verify]; await saveKata(d.verify, { ...k, verified: true }); return toast("確認済みにしました"); }
  if (d.editkata !== undefined) return kataForm(d.editkata);
  if (d.delkata) return delKata(d.delkata);
  if (d.katamemo) return openMemoForKata(d.katamemo);
  if (d.memo) { st.memoSel = d.memo; if (st.view !== "memo") { st.memoFilter = st.memos[d.memo]?.done ? "all" : st.memoFilter; go("memo"); } else renderMemo(false); return; }
  if (d.newmemo) return newMemo();
  if (d.mf) { st.memoFilter = d.mf; return renderMemoList(); }
  if (d.mpoff) { st.memoPart = null; return renderMemoList(); }
  if (d.mkoff) { st.memoKata = null; return renderMemoList(); }
  if (d.delmemo) { if (await confirmBox("このメモを削除しますか？", "削除", true)) { await store.remove(`memos/${d.delmemo}`); st.memoSel = null; renderMemo(false); renderMemoEditor(); toast("削除しました"); } return; }
  if (d.unlink) { const id = st.memoSel; if (!id) return; await store.update(`memos/${id}`, d.unlink === "part" ? { partId: null, partLabel: null } : { kataCode: null, kataName: null }); return renderMemoEditor(); }
  if (d.rerun) return rerun(d.rerun);
  if (d.delhist) { await store.remove(`history/${d.delhist}`); if (st.lastHist?.key === d.delhist) st.lastHist = null; return; }
  if (d.hf) { st.histFilter = d.hf; return renderHist(); }
  if (d.etab) { st.editTab = d.etab; return renderEdit(); }
  if (d.pick) return $("#" + d.pick)?.click();
});
document.addEventListener("change", async (e) => {
  const v = e.target.dataset.kver;
  if (v) { const k = st.kata[v]; await saveKata(v, { ...k, verified: e.target.checked }); }
});

/* ============================================================
   キーボード
   ============================================================ */
document.addEventListener("keydown", (e) => {
  if (!st.user || st.locked || modalOpen() || e.isComposing) return;
  const tag = document.activeElement?.tagName;
  const typing = /INPUT|TEXTAREA|SELECT/.test(tag) || document.activeElement?.isContentEditable;
  if (typing || e.ctrlKey || e.metaKey || e.altKey) return;
  const k = e.key.toLowerCase();
  if (e.key === "/") { e.preventDefault(); if (st.view !== "sec" && st.view !== "kata") go("sec"); ($("#q") || $("#kq"))?.focus(); return; }
  const map = { s: "sec", k: "kata", m: "memo", h: "hist", e: "edit" };
  if (map[k]) { e.preventDefault(); go(map[k]); return; }
  const onBody = !document.activeElement || document.activeElement === document.body;
  if (onBody && st.view === "sec" && ["ArrowDown", "ArrowUp", "Enter"].includes(e.key)) { secKeys(e); }
});

boot();
