// 店頭ツール：売価計算・原価計算・パッドローター・バルブ検索
// 計算は全部整数でやる（Excelの小数のずれを出さないため）
import * as S from "./search.js?v=20261003b";

let C = null; // app.js から道具を受け取る
export function setup(ctx) {
  C = ctx;
}

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (s) =>
  String(s ?? "").replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
  );
const yen = (n) =>
  n == null || isNaN(n) ? "" : (n < 0 ? "-¥" : "¥") + Math.abs(Math.round(n)).toLocaleString();
const pct = (n, d = 1) => (n == null || !isFinite(n) ? "" : `${n.toFixed(d)}%`);

// 全角数字・カンマ・¥ を許して数字に
export function toNum(s) {
  const t = String(s ?? "")
    .normalize("NFKC")
    .replace(/[,¥￥\s円]/g, "")
    .replace(/[^\d.\-]/g, "");
  if (t === "" || t === "-" || t === ".") return null;
  const n = Number(t);
  return isFinite(n) ? n : null;
}

// num/den を unit 円単位で丸める（dir: up / down）
export function roundTo(num, den, unit, dir) {
  const q = num / (den * unit);
  return (dir === "up" ? Math.ceil(q) : Math.floor(q)) * unit;
}
export const ROUNDS = [
  { id: "up1", label: "1円↑", unit: 1, dir: "up" },
  { id: "down1", label: "1円↓", unit: 1, dir: "down" },
  { id: "up5", label: "5円↑", unit: 5, dir: "up" },
  { id: "down5", label: "5円↓", unit: 5, dir: "down" },
  { id: "up10", label: "10円↑", unit: 10, dir: "up" },
  { id: "down10", label: "10円↓", unit: 10, dir: "down" },
];
const roundOf = (id) => ROUNDS.find((r) => r.id === id) || ROUNDS[2];

// 定価 × 掛け率% → 売価
export function salePrice(list, ratePct, roundId) {
  const r = roundOf(roundId);
  const rate100 = Math.round(ratePct * 100); // 85.5% → 8550
  return roundTo(Math.round(list) * rate100, 10000, r.unit, r.dir);
}
// 原価 ×（1＋上乗せ%）→ 売価
export function markupPrice(cost, markPct, roundId) {
  const r = roundOf(roundId);
  const m100 = Math.round((100 + markPct) * 100);
  return roundTo(Math.round(cost) * m100, 10000, r.unit, r.dir);
}
// 税込 → 税抜（切り上げ）
export const exTax = (incl) => Math.ceil((Math.round(incl) * 10) / 11);

// パッドローター
// ローター1枚 =（セット − パッド）÷ 2。端数はパッド側に寄せる（パッドで利益を取る）
// ローターが原価以下なら：ローター = 原価 + 上乗せ、パッド = セット − ローター×2
export function padRotor({ set, pad, padCost, rotorCost, margin = 10 }) {
  if ([set, pad, padCost, rotorCost].some((v) => v == null)) return null;
  const half = Math.floor((set - pad) / 2);
  if (half > rotorCost) {
    const padPrice = set - half * 2;
    return {
      ok: true,
      rotor: half,
      pad: padPrice,
      padGain: padPrice - padCost,
      rotorGain: half - rotorCost,
      total: padPrice - padCost + (half - rotorCost) * 2,
    };
  }
  const rotor = rotorCost + margin;
  const padPrice = set - rotor * 2;
  const needDiscount = padPrice < padCost;
  return {
    ok: false,
    under: rotorCost - half,
    rotor,
    pad: padPrice,
    needDiscount,
    padGain: padPrice - padCost,
    rotorGain: margin,
    total: padPrice - padCost + margin * 2,
  };
}

// ---- 状態（行は端末に下書き保存。価格は外に出さない） ----
const LS = "secsearch.tools.v1";
const blank = (n, f) => Array.from({ length: n }, f);
function load() {
  try {
    return JSON.parse(localStorage.getItem(LS) || "null");
  } catch {
    return null;
  }
}
const T = Object.assign(
  {
    tab: "sale",
    sale: {
      mode: "rate",
      rate: 85,
      mark: 15,
      round: "up5",
      rows: blank(20, () => ({ n: "", l: "", c: "" })),
    },
    cost: { mode: "ex", rows: blank(20, () => ({ n: "", p: "", q: "" })) },
    pad: { pad: "", padCost: "", rotorCost: "", rotorList: "" },
    bulb: { q: "", shape: "", volt: "" },
  },
  load() || {},
);
let saveT;
function save() {
  clearTimeout(saveT);
  saveT = setTimeout(() => {
    try {
      localStorage.setItem(LS, JSON.stringify(T));
    } catch {}
  }, 300);
}

const setting = (k, d) => C.setting(k, d);
const padSets = () => ({
  normal: +setting("padSetNormal", 12000) || 12000,
  kei: +setting("padSetKei", 10000) || 10000,
});
const padMargin = () => +setting("padMargin", 10) || 10;

// ---- 画面 ----
const TABS = [
  ["sale", "売価計算"],
  ["cost", "原価計算"],
  ["pad", "パッド・ローター"],
  ["bulb", "バルブ検索"],
];
export function render(main) {
  main.innerHTML = `<div class="view-head"><div><h1 class="h1">ツール</h1>
    <div class="tip">入力したそばから計算します ／ Enterで下の行へ ／ Excelからの貼り付けもOK</div></div></div>
    <div class="subtabs tool-tabs">${TABS.map(([k, l]) => `<button class="subtab ${T.tab === k ? "on" : ""}" data-ttab="${k}">${l}</button>`).join("")}</div>
    <div id="tbody"></div>`;
  $$("[data-ttab]", main).forEach((b) =>
    b.addEventListener("click", () => {
      T.tab = b.dataset.ttab;
      save();
      render(main);
    }),
  );
  ({ sale: renderSale, cost: renderCost, pad: renderPad, bulb: renderBulb })[T.tab]();
}
// 描き直すたびにリスナーが重ならないよう、入れ物ごと作り直す
function fresh() {
  const o = $("#tbody");
  const n = o.cloneNode(false);
  o.replaceWith(n);
  return n;
}
export function onData(key) {
  if (T.tab === "bulb" && key === "bulbs" && $("#blist")) renderBulbList();
  if (T.tab === "pad" && key === "settings" && $("#padres") && !document.activeElement?.dataset?.pk)
    renderPad();
}

// 表の共通：Enterで下へ、矢印で上下、Excelの貼り付けを下に流す
function gridKeys(box, onPaste) {
  box.addEventListener("keydown", (e) => {
    const el = e.target;
    if (!el.matches("input[data-r]")) return;
    const r = +el.dataset.r,
      col = el.dataset.k;
    let to = null;
    if (e.key === "Enter" || e.key === "ArrowDown") to = r + 1;
    else if (e.key === "ArrowUp") to = r - 1;
    if (to == null) return;
    const next = $(`input[data-r="${to}"][data-k="${col}"]`, box);
    if (next) {
      e.preventDefault();
      next.focus();
      next.select();
    }
  });
  box.addEventListener("paste", (e) => {
    const el = e.target;
    if (!el.matches("input[data-r]")) return;
    const text = e.clipboardData?.getData("text") || "";
    if (!/[\n\t]/.test(text.trim())) return;
    e.preventDefault();
    const lines = text
      .replace(/\r/g, "")
      .replace(/\n+$/, "")
      .split("\n")
      .map((l) => l.split("\t"));
    onPaste(+el.dataset.r, el.dataset.k, lines);
  });
}
function fmtOnBlur(box) {
  box.addEventListener(
    "blur",
    (e) => {
      const el = e.target;
      if (!el.matches("input.num")) return;
      const n = toNum(el.value);
      el.value = n == null ? "" : n.toLocaleString();
    },
    true,
  );
  box.addEventListener(
    "focus",
    (e) => {
      const el = e.target;
      if (!el.matches("input.num")) return;
      el.value = el.value.replace(/,/g, "");
      setTimeout(() => el.select(), 0);
    },
    true,
  );
}
const showNum = (v) => {
  const n = toNum(v);
  return n == null ? "" : n.toLocaleString();
};

// ================= 売価計算 =================
function renderSale() {
  const box = fresh();
  const s = T.sale;
  const rates = [70, 75, 80, 85, 90, 95, 100];
  box.innerHTML = `<div class="tbar">
      <div class="seg2" role="tablist"><button class="${s.mode === "rate" ? "on" : ""}" data-smode="rate">掛け率で</button><button class="${s.mode === "mark" ? "on" : ""}" data-smode="mark">利益上乗せで</button></div>
      ${
        s.mode === "rate"
          ? `<div class="tfield"><label>売り掛け率（定価に対して）</label><div class="rate-row">${rates
              .map((r) => `<button class="chip ${+s.rate === r ? "on" : ""}" data-rate="${r}">${r}%</button>`)
              .join(
                "",
              )}<span class="pin"><input id="rate" class="inp num sm" inputmode="decimal" value="${esc(s.rate)}"><i>%</i></span></div></div>`
          : `<div class="tfield"><label>上乗せする利益（原価に対して）</label><span class="pin"><input id="mark" class="inp num sm" inputmode="decimal" value="${esc(s.mark)}"><i>%</i></span></div>`
      }
      <div class="tfield"><label>端数</label><div class="segs">${ROUNDS.map((r) => `<button class="${s.round === r.id ? "on" : ""}" data-round="${r.id}">${r.label}</button>`).join("")}</div></div>
    </div>
    <div class="tgrid sale ${s.mode}" id="sgrid">
      <div class="th"><span>#</span><span>品名・品番（任意）</span>${s.mode === "rate" ? "<span>定価</span>" : "<span class='opt'>定価（任意）</span>"}<span>原価</span><span>仕入れ掛</span><span>売価</span><span>粗利</span><span></span></div>
      ${s.rows.map((r, i) => saleRow(r, i)).join("")}
    </div>
    <div class="tfoot" id="sfoot"></div>
    <div class="tacts"><button class="btn" data-add="sale">＋ 20行ふやす</button><button class="btn" data-copy="sale">売価をコピー</button><span class="sp"></span><button class="btn danger" data-clear="sale">全部消す</button></div>`;
  wireSale(box);
  calcSale();
}
function saleRow(r, i) {
  return `<div class="tr" data-row="${i}"><span class="no">${i + 1}</span>
    <input class="inp nm" data-r="${i}" data-k="n" value="${esc(r.n)}" placeholder="—" autocomplete="off">
    <input class="inp num" data-r="${i}" data-k="l" value="${esc(showNum(r.l))}" inputmode="numeric" placeholder="定価" autocomplete="off">
    <input class="inp num" data-r="${i}" data-k="c" value="${esc(showNum(r.c))}" inputmode="numeric" placeholder="原価" autocomplete="off">
    <div class="outs"><output class="o-buy"></output><output class="o-sale"></output><output class="o-gain"></output></div>
    <button class="x" data-delrow="${i}" title="この行を消す" aria-label="この行を消す">×</button></div>`;
}
function wireSale(box) {
  const s = T.sale;
  box.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const d = b.dataset;
    if (d.smode) {
      s.mode = d.smode;
      if (s.mode === "mark" && s.round === "up5") s.round = "up10";
      if (s.mode === "rate" && s.round === "up10") s.round = "up5";
    } else if (d.rate) s.rate = +d.rate;
    else if (d.round) s.round = d.round;
    else if (d.add) s.rows.push(...blank(20, () => ({ n: "", l: "", c: "" })));
    else if (d.delrow) {
      s.rows.splice(+d.delrow, 1);
      if (s.rows.length < 20) s.rows.push({ n: "", l: "", c: "" });
    } else if (d.clear) return clearRows("sale");
    else if (d.copy) return copyCol("sale");
    else return;
    save();
    renderSale();
  });
  $("#rate", box)?.addEventListener("input", (e) => {
    s.rate = toNum(e.target.value) ?? 0;
    $$("[data-rate]", box).forEach((c) => c.classList.toggle("on", +c.dataset.rate === s.rate));
    save();
    calcSale();
  });
  $("#mark", box)?.addEventListener("input", (e) => {
    s.mark = toNum(e.target.value) ?? 0;
    save();
    calcSale();
  });
  const grid = $("#sgrid", box);
  grid.addEventListener("input", (e) => {
    const el = e.target;
    if (!el.matches("input[data-r]")) return;
    s.rows[+el.dataset.r][el.dataset.k] = el.value;
    save();
    calcSale();
  });
  gridKeys(grid, (r0, k0, lines) => {
    const cols = ["n", "l", "c"];
    const c0 = cols.indexOf(k0);
    while (s.rows.length < r0 + lines.length) s.rows.push(...blank(20, () => ({ n: "", l: "", c: "" })));
    lines.forEach((cells, i) =>
      cells.forEach((v, j) => cols[c0 + j] && (s.rows[r0 + i][cols[c0 + j]] = v.trim())),
    );
    save();
    renderSale();
    C.toast(`${lines.length}行 貼り付けました`);
  });
  fmtOnBlur(grid);
}
function saleCalcRow(r) {
  const s = T.sale;
  const l = toNum(r.l),
    c = toNum(r.c);
  let sale = null;
  if (s.mode === "rate" && l != null) sale = salePrice(l, s.rate, s.round);
  if (s.mode === "mark" && c != null) sale = markupPrice(c, s.mark, s.round);
  const buy = l && c != null ? (c / l) * 100 : null;
  const gain = sale != null && c != null ? sale - c : null;
  return { l, c, sale, buy, gain, gp: gain != null && sale ? (gain / sale) * 100 : null };
}
function calcSale() {
  const grid = $("#sgrid");
  if (!grid) return;
  let n = 0,
    tl = 0,
    tc = 0,
    ts = 0,
    tg = 0,
    warn = 0;
  T.sale.rows.forEach((r, i) => {
    const row = $(`.tr[data-row="${i}"]`, grid);
    if (!row) return;
    const x = saleCalcRow(r);
    $(".o-buy", row).innerHTML =
      x.buy != null ? `${pct(x.buy)}<small>${(x.buy / 10).toFixed(1)}掛</small>` : "";
    $(".o-sale", row).textContent = x.sale != null ? yen(x.sale) : "";
    $(".o-gain", row).innerHTML = x.gain != null ? `${yen(x.gain)}<small>${pct(x.gp)}</small>` : "";
    const bad = x.gain != null && x.gain < 0;
    row.classList.toggle("bad", bad);
    row.classList.toggle("filled", x.sale != null);
    if (x.sale != null) {
      n++;
      ts += x.sale;
      if (x.l != null) tl += x.l;
      if (x.c != null) tc += x.c;
      if (x.gain != null) tg += x.gain;
      if (bad) warn++;
    }
  });
  const s = T.sale;
  const how =
    s.mode === "rate"
      ? `定価 × ${s.rate}% → ${roundOf(s.round).label}`
      : `原価 ×（100＋${s.mark}）% → ${roundOf(s.round).label}`;
  $("#sfoot").innerHTML = `<div class="sum"><span>${n}件</span><small class="how">${esc(how)}</small></div>
    <div class="sum minor"><small>定価計</small><b>${yen(tl)}</b></div><div class="sum minor"><small>原価計</small><b>${yen(tc)}</b></div>
    <div class="sum hl"><small>売価計</small><b>${yen(ts)}</b></div><div class="sum"><small>粗利計</small><b class="${tg < 0 ? "ng" : "ok"}">${yen(tg)}</b><small>${ts ? pct((tg / ts) * 100) : ""}</small></div>
    ${warn ? `<div class="sum warn">原価割れ ${warn}件</div>` : ""}`;
}

// ================= 原価計算 =================
function renderCost() {
  const box = fresh();
  const c = T.cost;
  box.innerHTML = `<div class="tbar">
      <div class="seg2"><button class="${c.mode === "ex" ? "on" : ""}" data-cmode="ex">税抜きで足す（通常）</button><button class="${c.mode === "in" ? "on" : ""}" data-cmode="in">税込みから税抜きに（スズキ）</button></div>
      <div class="tnote">${c.mode === "in" ? "1行ずつ 税込み ÷ 1.1 を<b>1円単位で切り上げ</b>てから足します" : "入れた金額をそのまま足します。数量を入れると × します"}</div>
    </div>
    <div class="tgrid cost ${c.mode}" id="cgrid">
      <div class="th"><span>#</span><span>品名・品番（任意）</span><span>${c.mode === "in" ? "税込み" : "税抜き"}</span><span>数量</span>${c.mode === "in" ? "<span>税抜き</span>" : ""}<span>小計</span><span></span></div>
      ${c.rows.map((r, i) => costRow(r, i)).join("")}
    </div>
    <div class="tfoot" id="cfoot"></div>
    <div class="tacts"><button class="btn" data-add="cost">＋ 20行ふやす</button><button class="btn" data-copy="cost">合計をコピー</button><span class="sp"></span><button class="btn danger" data-clear="cost">全部消す</button></div>`;
  wireCost(box);
  calcCost();
}
function costRow(r, i) {
  return `<div class="tr" data-row="${i}"><span class="no">${i + 1}</span>
    <input class="inp nm" data-r="${i}" data-k="n" value="${esc(r.n)}" placeholder="—" autocomplete="off">
    <input class="inp num" data-r="${i}" data-k="p" value="${esc(showNum(r.p))}" inputmode="numeric" placeholder="金額" autocomplete="off">
    <input class="inp num qty" data-r="${i}" data-k="q" value="${esc(r.q)}" inputmode="numeric" placeholder="1" autocomplete="off">
    <div class="outs">${T.cost.mode === "in" ? '<output class="o-ex"></output>' : ""}<output class="o-sub"></output></div>
    <button class="x" data-delrow="${i}" title="この行を消す" aria-label="この行を消す">×</button></div>`;
}
function wireCost(box) {
  const c = T.cost;
  box.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const d = b.dataset;
    if (d.cmode) c.mode = d.cmode;
    else if (d.add) c.rows.push(...blank(20, () => ({ n: "", p: "", q: "" })));
    else if (d.delrow) {
      c.rows.splice(+d.delrow, 1);
      if (c.rows.length < 20) c.rows.push({ n: "", p: "", q: "" });
    } else if (d.clear) return clearRows("cost");
    else if (d.copy) return copyCol("cost");
    else return;
    save();
    renderCost();
  });
  const grid = $("#cgrid", box);
  grid.addEventListener("input", (e) => {
    const el = e.target;
    if (!el.matches("input[data-r]")) return;
    c.rows[+el.dataset.r][el.dataset.k] = el.value;
    save();
    calcCost();
  });
  gridKeys(grid, (r0, k0, lines) => {
    const cols = ["n", "p", "q"];
    const c0 = cols.indexOf(k0);
    while (c.rows.length < r0 + lines.length) c.rows.push(...blank(20, () => ({ n: "", p: "", q: "" })));
    lines.forEach((cells, i) =>
      cells.forEach((v, j) => cols[c0 + j] && (c.rows[r0 + i][cols[c0 + j]] = v.trim())),
    );
    save();
    renderCost();
    C.toast(`${lines.length}行 貼り付けました`);
  });
  fmtOnBlur(grid);
}
function costCalcRow(r) {
  const p = toNum(r.p);
  if (p == null) return null;
  const q = toNum(r.q) ?? 1;
  const unit = T.cost.mode === "in" ? exTax(p) : Math.round(p);
  return { unit, q, sub: unit * q };
}
function costTotal() {
  return T.cost.rows.reduce(
    (a, r) => {
      const x = costCalcRow(r);
      if (x) {
        a.sum += x.sub;
        a.n++;
        a.q += x.q;
      }
      return a;
    },
    { sum: 0, n: 0, q: 0 },
  );
}
function calcCost() {
  const grid = $("#cgrid");
  if (!grid) return;
  T.cost.rows.forEach((r, i) => {
    const row = $(`.tr[data-row="${i}"]`, grid);
    if (!row) return;
    const x = costCalcRow(r);
    const ex = $(".o-ex", row);
    if (ex) ex.textContent = x ? yen(x.unit) : "";
    $(".o-sub", row).textContent = x ? yen(x.sub) : "";
    row.classList.toggle("filled", !!x);
  });
  const t = costTotal();
  $("#cfoot").innerHTML = `<div class="sum"><span>${t.n}行</span><small>${t.q}点</small></div>
    <div class="sum hl big"><small>合計（税抜き）</small><b>${yen(t.sum)}</b></div>
    <div class="sum minor"><small>参考：税込み（×1.1）</small><b>${yen(Math.floor((t.sum * 11) / 10))}</b></div>`;
}

async function clearRows(kind) {
  if (!(await C.confirmBox("入力した行を全部消しますか？", "全部消す", true))) return;
  if (kind === "sale") T.sale.rows = blank(20, () => ({ n: "", l: "", c: "" }));
  else T.cost.rows = blank(20, () => ({ n: "", p: "", q: "" }));
  save();
  render($("#main"));
}
async function copyCol(kind) {
  let text = "";
  if (kind === "sale") {
    text = T.sale.rows
      .map(saleCalcRow)
      .filter((x) => x.sale != null)
      .map((x) => x.sale)
      .join("\n");
  } else text = String(costTotal().sum);
  if (!text) return C.toast("コピーするものがありません", true);
  await C.copyText(text);
  C.toast(
    kind === "sale" ? "売価をコピーしました（Excelに貼れます）" : `合計 ${yen(+text)} をコピーしました`,
  );
}

// ================= パッド・ローター =================
function renderPad() {
  const box = fresh();
  const p = T.pad;
  const sets = padSets();
  const f = (k, label, ph) =>
    `<div class="field"><label>${label}</label><input class="inp num big" data-pk="${k}" value="${esc(showNum(p[k]))}" inputmode="numeric" placeholder="${ph}" autocomplete="off"></div>`;
  box.innerHTML = `<div class="pad-grid">
    <div class="card"><h3>入力</h3>
      ${f("pad", "パッド通常売価", "例：4,910")}
      ${f("padCost", "パッド原価", "例：3,510")}
      ${f("rotorCost", "ローター原価（1枚）", "例：3,290")}
      ${f("rotorList", "ローター通常売価（1枚・任意）", "お得額の表示用")}
      <div class="pad-set">セット価格：普通車 <b>${yen(sets.normal)}</b> ／ 軽 <b>${yen(sets.kei)}</b> ／ 原価割れ時のローター利益 <b>+${padMargin()}円</b>
        <button class="linkbtn" data-padset="1">変更</button></div>
      <button class="btn sm" data-padclear="1">入力を消す</button>
    </div>
    <div id="padres" class="pad-res"></div></div>
    <details class="pad-rule"><summary>計算のルール</summary><ol>
      <li>ローターを安くして、パッドで利益を取る</li><li>原価は割らない（値引きを使わない）</li>
      <li>ローター1枚 ＝（セット価格 − パッド通常売価）÷ 2。割り切れないときは、ローターを1円下げてパッドに回す</li>
      <li>ローターが原価割れするとき：ローター ＝ 原価＋${padMargin()}円、パッド ＝ セット価格 − ローター×2</li>
      <li>そのパッドがパッド原価を下回るときは「値引き必要」</li></ol></details>`;
  box.addEventListener("input", (e) => {
    const k = e.target.dataset.pk;
    if (!k) return;
    p[k] = e.target.value;
    save();
    calcPad();
  });
  fmtOnBlur(box);
  box.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !e.target.dataset.pk) return;
    const all = $$("[data-pk]", box);
    const i = all.indexOf(e.target);
    if (all[i + 1]) {
      e.preventDefault();
      all[i + 1].focus();
    } else e.target.blur();
  });
  box.addEventListener("click", async (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    if (b.dataset.padclear) {
      T.pad = { pad: "", padCost: "", rotorCost: "", rotorList: "" };
      save();
      renderPad();
    }
    if (b.dataset.padset) editPadSet();
  });
  calcPad();
}
function padCard(label, set, x, rotorList, pad) {
  if (!x)
    return `<div class="pres empty"><h4>${label} <span class="mono">${yen(set)}</span></h4><p>パッド通常売価・パッド原価・ローター原価を入れてください</p></div>`;
  const normalTotal = rotorList != null && pad != null ? pad + rotorList * 2 : null;
  const head = x.ok
    ? `<div class="verdict ok"><span class="lamp g"></span><b>そのままでOK</b></div>`
    : x.needDiscount
      ? `<div class="verdict ng"><span class="lamp r"></span><b>値引きが必要</b><small>ローターが原価割れ（${yen(x.under)}）。原価＋${padMargin()}円にしても、パッドが原価を ${yen(-x.padGain)} 下回ります</small></div>`
      : `<div class="verdict warn"><span class="lamp a"></span><b>パッドで調整</b><small>ローターが原価割れ（${yen(x.under)}）するので、ローター＝原価＋${padMargin()}円、残りをパッドに</small></div>`;
  return `<div class="pres ${x.ok ? "ok" : x.needDiscount ? "ng" : "warn"}"><h4>${label} <span class="mono">${yen(set)}</span></h4>${head}
    <table class="kv"><tr><td>パッド</td><td class="mono">${yen(x.pad)}</td><td class="mono g ${x.padGain < 0 ? "ng" : ""}">${x.padGain >= 0 ? "+" : ""}${yen(x.padGain)}</td></tr>
    <tr class="hl"><td>ローター 1枚</td><td class="mono">${yen(x.rotor)} <small>×2</small></td><td class="mono g">+${yen(x.rotorGain)}<small>/枚</small></td></tr>
    <tr class="tot"><td>セット合計</td><td class="mono">${yen(x.pad + x.rotor * 2)}</td><td class="mono g ${x.total < 0 ? "ng" : ""}">利益 ${yen(x.total)}</td></tr>
    ${normalTotal != null ? `<tr><td>通常で買うと</td><td class="mono">${yen(normalTotal)}</td><td class="mono save">${normalTotal > set ? `${yen(normalTotal - set)} お得` : ""}</td></tr>` : ""}</table></div>`;
}
function calcPad() {
  const box = $("#padres");
  if (!box) return;
  const p = T.pad;
  const v = { pad: toNum(p.pad), padCost: toNum(p.padCost), rotorCost: toNum(p.rotorCost) };
  const rl = toNum(p.rotorList);
  const sets = padSets();
  const m = padMargin();
  box.innerHTML =
    padCard("普通車", sets.normal, padRotor({ set: sets.normal, ...v, margin: m }), rl, v.pad) +
    padCard("軽自動車", sets.kei, padRotor({ set: sets.kei, ...v, margin: m }), rl, v.pad);
}
async function editPadSet() {
  const s = padSets();
  const r = await C.modal(`<h3>パッド・ローターの設定</h3><form>
    <div class="grid2"><div class="field"><label>普通車のセット価格</label><input class="inp mono" name="normal" inputmode="numeric" value="${s.normal}" required></div>
    <div class="field"><label>軽自動車のセット価格</label><input class="inp mono" name="kei" inputmode="numeric" value="${s.kei}" required></div></div>
    <div class="field"><label>原価割れのときのローター利益（円）</label><input class="inp mono" name="margin" inputmode="numeric" value="${padMargin()}" required></div></form>
    <div class="modal-foot"><button class="btn" data-m="cancel">キャンセル</button><button class="btn pri" data-m="ok">保存</button></div>`);
  if (!r) return;
  const n = toNum(r.data.normal),
    k = toNum(r.data.kei),
    mg = toNum(r.data.margin);
  if (!n || !k || mg == null) return C.toast("数字を入れてください", true);
  await C.store.update("settings", { padSetNormal: n, padSetKei: k, padMargin: mg });
  C.toast("保存しました");
}

// ================= バルブ検索 =================
const prettySpec = (s) =>
  String(s || "")
    .replace(/^(バルブ|BULB|ハロゲン|バルフ)\s*/i, "")
    .replace(/(\d+)\s*v\s*-?\s*/i, "$1V ")
    .replace(/([\d./]+)\s*w\b/i, "$1W")
    .trim();
const voltOf = (b) => (/24\s*v/i.test(b.spec + b.cat) ? "24V" : /12\s*v/i.test(b.spec) ? "12V" : "");
const shapeKey = (s) =>
  String(s || "")
    .toUpperCase()
    .replace(/\s/g, "");
// お客さんの呼び方 ⇔ カタログの書き方
const BULB_SYN = [
  "ウインカー=方向指示=ターン",
  "ブレーキ=制動=ストップ",
  "テール=尾灯",
  "ナンバー=ライセンス",
  "バック=後退",
  "ルーム=室内=車内",
  "ポジション=車幅",
  "ヘッドライト=前照灯=ヘッド",
  "フォグ=霧灯",
].map((l) => l.split("=").map(S.norm));
let bulbIdx = null,
  bulbSrc = null;
function bulbList() {
  const raw = C.st.bulbs || {};
  if (raw === bulbSrc && bulbIdx) return bulbIdx;
  bulbSrc = raw;
  bulbIdx = Object.entries(raw).map(([id, b]) => ({
    id,
    ...b,
    _s: S.norm(`${b.spec} ${b.code} ${b.cat} ${b.shape} ${b.use} ${b.loc} ${b.note}`),
    _v: voltOf(b),
    _sh: shapeKey(b.shape),
  }));
  return bulbIdx;
}
function renderBulb() {
  const box = fresh();
  box.innerHTML = `<label class="search">${C.IC.search}<input id="bq" type="search" autocomplete="off" placeholder="形状・規格・用途・品番・棚番（例：t10、h4、ウインカー、24v）" value="${esc(T.bulb.q)}"><span class="cnt" id="bcnt"></span></label>
    <div class="chips" id="bchips"></div>
    <div id="blist"></div>
    <div class="tacts"><input type="file" id="fbulb" accept=".xlsx,.xlsm,.xls,.json" hidden><button class="btn" data-bpick="1">Excelから取り込み</button><button class="btn pri" data-bedit="">＋ バルブを追加</button></div>`;
  const q = $("#bq", box);
  q.addEventListener("input", () => {
    T.bulb.q = q.value;
    save();
    renderBulbList();
  });
  q.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      q.value = T.bulb.q = "";
      renderBulbList();
    }
  });
  $("#fbulb", box).addEventListener("change", (e) =>
    importBulbs(e.target.files[0]).finally(() => (e.target.value = "")),
  );
  box.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const d = b.dataset;
    if (d.bpick) $("#fbulb").click();
    else if (d.bedit !== undefined) bulbForm(d.bedit || null);
    else if (d.bshape !== undefined) {
      T.bulb.shape = T.bulb.shape === d.bshape ? "" : d.bshape;
      save();
      renderBulbList();
    } else if (d.bvolt !== undefined) {
      T.bulb.volt = T.bulb.volt === d.bvolt ? "" : d.bvolt;
      save();
      renderBulbList();
    } else if (d.bcopy) C.copyText(d.bcopy).then(() => C.toast(`${d.bcopy} をコピーしました`));
    else if (d.bgoogle) C.googleOpen(d.bgoogle);
  });
  renderBulbList();
  if (matchMedia("(hover: hover)").matches) setTimeout(() => q.focus(), 0);
}
function renderBulbList() {
  const list = bulbList();
  const box = $("#blist");
  if (!box) return;
  // 形状チップは多い順
  const cnt = {};
  for (const b of list) if (b._sh) cnt[b._sh] = (cnt[b._sh] || 0) + 1;
  const shapes = Object.keys(cnt).sort(
    (a, b) => cnt[b] - cnt[a] || a.localeCompare(b, "ja", { numeric: true }),
  );
  $("#bchips").innerHTML =
    shapes
      .map(
        (s) =>
          `<button class="chip ${T.bulb.shape === s ? "on" : ""}" data-bshape="${esc(s)}">${esc(s)}</button>`,
      )
      .join("") +
    `<span class="chip-sep"></span>` +
    ["12V", "24V"]
      .map((v) => `<button class="chip ${T.bulb.volt === v ? "on" : ""}" data-bvolt="${v}">${v}</button>`)
      .join("");
  if (!list.length) {
    $("#bcnt").textContent = "";
    box.innerHTML = `<div class="panel"><h3>バルブのデータがまだありません</h3><p style="margin:0;font-size:12.5px;color:var(--mu);line-height:1.7">下の「Excelから取り込み」で、バルブ適合のExcel（DBシート＋ロケーションのシート）か bulbs.json を選んでください。</p></div>`;
    return;
  }
  const toks = String(T.bulb.q || "")
    .normalize("NFKC")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((t) => S.expand(t, [...BULB_SYN, ...C.st.syn]));
  const hits = list.filter(
    (b) =>
      (!T.bulb.shape || b._sh === T.bulb.shape) &&
      (!T.bulb.volt || b._v === T.bulb.volt) &&
      toks.every((alts) => alts.some((a) => b._s.includes(a))),
  );
  hits.sort(
    (a, b) =>
      (b.note?.includes("一番") ? 1 : 0) - (a.note?.includes("一番") ? 1 : 0) ||
      a._sh.localeCompare(b._sh, "ja", { numeric: true }),
  );
  $("#bcnt").textContent = `${hits.length} / ${list.length}件`;
  box.innerHTML =
    hits
      .map(
        (b) => `<div class="brow">
      <div class="bloc"><small>棚</small><b>${esc(b.loc || "—")}</b></div>
      <div class="bmain"><div class="bt1"><span class="shape">${esc(b.shape || "?")}</span><b>${esc(prettySpec(b.spec))}</b><span class="cat">${esc(b.cat)}</span>${b.note ? `<span class="hot">${b.note.includes("一番") ? "★" : ""}${esc(b.note)}</span>` : ""}</div>
        <div class="buse">${esc(b.use)}</div></div>
      <div class="bside"><button class="bcode" data-bcopy="${esc(b.code)}" title="品番をコピー">${esc(b.code)}</button><small>${b.unit ? `${esc(b.unit)}個入` : ""}</small>
        <div class="bbtn"><button class="btn sm" data-bgoogle="${esc(b.code)}">Google</button><button class="btn sm" data-bedit="${esc(b.id)}">編集</button></div></div></div>`,
      )
      .join("") || '<p class="empty">該当なし</p>';
}
async function bulbForm(id) {
  const b = id
    ? C.st.bulbs[id]
    : { spec: "", code: "", cat: "", shape: "", use: "", loc: "", unit: "", note: "" };
  if (!b) return;
  const f = (k, l, extra = "", cls = "") =>
    `<div class="field"><label>${l}</label><input class="inp ${cls}" name="${k}" value="${esc(b[k])}" ${extra}></div>`;
  const r = await C.modal(`<h3>${id ? "バルブを編集" : "バルブを追加"}</h3><form>
    <div class="grid2">${f("code", "品番", "required", "mono")}${f("loc", "棚番（ロケーション）")}</div>
    <div class="grid2">${f("shape", "形状（T10・H4 など）", "required")}${f("spec", "規格（12V 5W など）")}</div>
    <div class="grid2">${f("cat", "分類（ウェッジ・シングル球 など）")}${f("unit", "販売単位（個入）", 'inputmode="numeric"')}</div>
    ${f("use", "用途・呼び方")}${f("note", "メモ（よく出る、など）")}</form>
    <div class="modal-foot">${id ? '<button class="btn danger left" data-m="delete">削除</button>' : ""}<button class="btn" data-m="cancel">キャンセル</button><button class="btn pri" data-m="ok">保存</button></div>`);
  if (!r) return;
  if (r.act === "delete") {
    if (!(await C.confirmBox(`${b.code}（${b.shape}）を削除しますか？`, "削除", true))) return;
    await C.store.remove(`bulbs/${id}`);
    return C.toast("削除しました");
  }
  const N = (s) =>
    String(s ?? "")
      .normalize("NFKC")
      .trim();
  const rec = {};
  for (const k of ["code", "loc", "shape", "spec", "cat", "use", "note"]) rec[k] = N(r.data[k]);
  rec.code = rec.code.toUpperCase();
  rec.unit = toNum(r.data.unit) ?? "";
  if (id) await C.store.set(`bulbs/${id}`, rec);
  else await C.store.push("bulbs", rec);
  C.toast("保存しました");
}

// Excel（DBシート＋ロケーションのシート）または bulbs.json
function bulbsFromRows(sheets) {
  const N = (s) =>
    String(s ?? "")
      .normalize("NFKC")
      .replace(/\s+/g, " ")
      .trim();
  let base = null;
  const loc = {};
  for (const rows of sheets) {
    for (let h = 0; h < Math.min(rows.length, 10); h++) {
      const hd = (rows[h] || []).map((c) => N(c).replace(/\s/g, ""));
      const iSpec = hd.indexOf("規格"),
        iCode = hd.indexOf("品番"),
        iShape = hd.indexOf("形状");
      if (iSpec >= 0 && iCode >= 0 && iShape >= 0 && !base) {
        const iCat = hd.indexOf("分類"),
          iUse = hd.findIndex((c) => c.startsWith("用途"));
        base = rows
          .slice(h + 1)
          .filter((r) => N(r[iCode]))
          .map((r) => {
            const raw = N(r[iCode]).toUpperCase();
            const m = raw.match(/^([A-Z0-9]{3,6}-[A-Z0-9]{3,6})\s*(.*)$/);
            return {
              code: m ? m[1] : raw,
              note: m ? m[2].replace(/^[(（]|[)）]$/g, "").trim() : "",
              spec: N(r[iSpec]),
              cat: iCat >= 0 ? N(r[iCat]) : "",
              shape: N(r[iShape]),
              use: iUse >= 0 ? N(r[iUse]) : "",
            };
          });
        break;
      }
      const iLoc = hd.findIndex((c) => c.startsWith("ロケーション") || c === "棚番");
      const iNo = hd.findIndex((c) => c.startsWith("部品番号") || c === "品番");
      if (iLoc >= 0 && iNo >= 0) {
        const iUnit = hd.indexOf("販売単位");
        for (const r of rows.slice(h + 1)) {
          const c = N(r[iNo]).toUpperCase();
          if (c) loc[c] = { loc: N(r[iLoc]), unit: iUnit >= 0 ? (toNum(r[iUnit]) ?? "") : "" };
        }
        break;
      }
    }
  }
  if (!base) throw new Error("「規格」「品番」「形状」の列が見つかりませんでした");
  return base.map((b) => ({ ...b, loc: loc[b.code]?.loc || "", unit: loc[b.code]?.unit ?? "" }));
}
async function importBulbs(file) {
  if (!file) return;
  let list;
  try {
    if (file.name.toLowerCase().endsWith(".json")) {
      const j = JSON.parse(await file.text());
      list = Array.isArray(j) ? j : j.bulbs;
    } else {
      const X = await C.loadXlsx();
      const wb = X.read(await file.arrayBuffer(), { type: "array" });
      list = bulbsFromRows(
        wb.SheetNames.map((n) => X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: "" })),
      );
    }
    if (!Array.isArray(list) || !list.length) throw new Error("バルブのデータが見つかりませんでした");
  } catch (e) {
    return C.toast("取り込めませんでした：" + e.message, true);
  }
  const have = Object.keys(C.st.bulbs || {}).length;
  const withLoc = list.filter((b) => b.loc).length;
  const ok = await C.confirmBox(
    `${list.length}件 のバルブを取り込みます（棚番つき ${withLoc}件）。${have ? `\n今の ${have}件 は入れ替わります。` : ""}`,
    "取り込む",
  );
  if (!ok) return;
  const obj = {};
  list.forEach((b) => {
    obj[C.localKey()] = {
      code: String(b.code || "").toUpperCase(),
      spec: b.spec || "",
      cat: b.cat || "",
      shape: b.shape || "",
      use: b.use || "",
      loc: b.loc || "",
      unit: b.unit ?? "",
      note: b.note || "",
    };
  });
  await C.store.set("bulbs", obj);
  C.toast(`${list.length}件 取り込みました`);
}
