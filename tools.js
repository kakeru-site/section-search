// 店頭ツール：売価計算・原価計算・パッドローター・バルブ検索・オイル／ケミカル検索
// 計算は全部整数でやる（Excelの小数のずれを出さないため）
import * as S from "./search.js?v=20261014s";
import { buildOilWorkbook } from "./oilxlsx.js?v=20261014s";
import { readBase, injectSheets, uniqueName } from "./xlsxinject.js?v=20261014s";

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
// スズキのシートと同じ：消費税 = ROUND(税込み/11)、税抜き = 税込み - 消費税
// ExcelのROUNDは0から遠い方へ丸めるので、マイナス（返品）も符号ごと扱う
const xlRound = (n) => Math.sign(n) * Math.round(Math.abs(n));
export const taxOf = (incl) => xlRound(Math.round(incl) / 11);
export const exTax = (incl) => Math.round(incl) - taxOf(incl);

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
      rows: blank(20, () => ({ l: "", c: "" })),
    },
    cost: { mode: "ex", rows: blank(20, () => ({ p: "", q: "" })) },
    pad: { pad: "", padCost: "", rotorCost: "", rotorList: "" },
    bulb: { q: "", shape: "", volt: "" },
    oil: { q: "", cat: "", protect: true },
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
// 値引きが必要なときの提案：パッド原価にいくら乗せるか
const padPlus = () => [+setting("padPlus1", 500) || 500, +setting("padPlus2", 1000) || 1000];

// ---- 画面 ----
const TABS = [
  ["sale", "売価計算"],
  ["cost", "原価計算"],
  ["pad", "パッド・ローター"],
  ["bulb", "バルブ検索"],
  ["oil", "オイル・ケミカル"],
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
  ({ sale: renderSale, cost: renderCost, pad: renderPad, bulb: renderBulb, oil: renderOil })[T.tab]();
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
  if (T.tab === "oil" && (key === "oils" || key === "memos") && $("#olist")) renderOilList();
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
  const rates = [50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100];
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
      <button class="btn clear-all" data-clear="sale" title="金額・数量を全部消す（掛け率などの設定はそのまま）">入力を全部消す</button>
    </div>
    <div class="tgrid sale ${s.mode}" id="sgrid">
      <div class="th"><span>#</span>${s.mode === "rate" ? "<span>定価</span>" : ""}<span>原価</span>${s.mode === "rate" ? "<span>仕入れ掛</span>" : ""}<span>売価</span><span>粗利</span><span></span></div>
      ${s.rows.map((r, i) => saleRow(r, i)).join("")}
    </div>
    <div class="tfoot" id="sfoot"></div>
    <div class="tacts"><button class="btn" data-add="sale">＋ 20行ふやす</button><button class="btn" data-copy="sale">売価をコピー</button><span class="sp"></span><button class="btn danger" data-clear="sale">入力を全部消す</button></div>`;
  wireSale(box);
  calcSale();
}
function saleRow(r, i) {
  const rate = T.sale.mode === "rate";
  return `<div class="tr" data-row="${i}"><span class="no">${i + 1}</span>
    ${rate ? `<input class="inp num" data-r="${i}" data-k="l" value="${esc(showNum(r.l))}" inputmode="numeric" placeholder="定価" autocomplete="off">` : ""}
    <input class="inp num" data-r="${i}" data-k="c" value="${esc(showNum(r.c))}" inputmode="numeric" placeholder="原価" autocomplete="off">
    <div class="outs">${rate ? '<output class="o-buy"></output>' : ""}<output class="o-sale"></output><output class="o-gain"></output></div>
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
    else if (d.add) s.rows.push(...blank(20, () => ({ l: "", c: "" })));
    else if (d.delrow) {
      s.rows.splice(+d.delrow, 1);
      if (s.rows.length < 20) s.rows.push({ l: "", c: "" });
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
    const cols = s.mode === "rate" ? ["l", "c"] : ["c"];
    const c0 = cols.indexOf(k0);
    while (s.rows.length < r0 + lines.length) s.rows.push(...blank(20, () => ({ l: "", c: "" })));
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
  const buy = s.mode === "rate" && l && c != null ? (c / l) * 100 : null;
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
  const s = T.sale;
  s.rows.forEach((r, i) => {
    const row = $(`.tr[data-row="${i}"]`, grid);
    if (!row) return;
    const x = saleCalcRow(r);
    const ob = $(".o-buy", row);
    if (ob) ob.innerHTML = x.buy != null ? `${pct(x.buy)}<small>${(x.buy / 10).toFixed(1)}掛</small>` : "";
    $(".o-sale", row).textContent = x.sale != null ? yen(x.sale) : "";
    $(".o-gain", row).innerHTML = x.gain != null ? `${yen(x.gain)}<small>${pct(x.gp)}</small>` : "";
    const bad = x.gain != null && x.gain < 0;
    row.classList.toggle("bad", bad);
    row.classList.toggle("filled", x.sale != null);
    if (x.sale != null) {
      n++;
      ts += x.sale;
      if (x.l != null && s.mode === "rate") tl += x.l;
      if (x.c != null) tc += x.c;
      if (x.gain != null) tg += x.gain;
      if (bad) warn++;
    }
  });
  const how =
    s.mode === "rate"
      ? `定価 × ${s.rate}% → ${roundOf(s.round).label}`
      : `原価 ×（100＋${s.mark}）% → ${roundOf(s.round).label}`;
  $("#sfoot").innerHTML = `<div class="sum"><span>${n}件</span><small class="how">${esc(how)}</small></div>
    ${s.mode === "rate" ? `<div class="sum minor"><small>定価計</small><b>${yen(tl)}</b></div>` : ""}<div class="sum minor"><small>原価計</small><b>${yen(tc)}</b></div>
    <div class="sum hl"><small>売価計</small><b>${yen(ts)}</b></div><div class="sum"><small>粗利計</small><b class="${tg < 0 ? "ng" : "ok"}">${yen(tg)}</b><small>${ts ? pct((tg / ts) * 100) : ""}</small></div>
    ${warn ? `<div class="sum warn">原価割れ ${warn}件</div>` : ""}`;
}

// ================= 原価計算 =================
function renderCost() {
  const box = fresh();
  const c = T.cost;
  box.innerHTML = `<div class="tbar">
      <div class="seg2"><button class="${c.mode === "ex" ? "on" : ""}" data-cmode="ex">税抜きで足す（通常）</button><button class="${c.mode === "in" ? "on" : ""}" data-cmode="in">税込みから税抜きに（スズキ）</button></div>
      <div class="tnote">${c.mode === "in" ? "1行ずつ <b>消費税 ＝ 税込み ÷ 11（四捨五入）</b>、税抜き ＝ 税込み − 消費税 にしてから足します。マイナス（返品）もOK" : "入れた金額をそのまま足します。数量を入れると × します"}</div>
      <button class="btn clear-all" data-clear="cost" title="金額・数量を全部消す（掛け率などの設定はそのまま）">入力を全部消す</button>
    </div>
    <div class="tgrid cost ${c.mode}" id="cgrid">
      <div class="th"><span>#</span><span>${c.mode === "in" ? "税込み" : "税抜き"}</span><span>数量</span>${c.mode === "in" ? "<span>税抜き</span>" : ""}<span>小計</span><span></span></div>
      ${c.rows.map((r, i) => costRow(r, i)).join("")}
    </div>
    <div class="tfoot" id="cfoot"></div>
    <div class="tacts"><button class="btn" data-add="cost">＋ 20行ふやす</button><button class="btn" data-copy="cost">合計をコピー</button><span class="sp"></span><button class="btn danger" data-clear="cost">入力を全部消す</button></div>`;
  wireCost(box);
  calcCost();
}
function costRow(r, i) {
  return `<div class="tr" data-row="${i}"><span class="no">${i + 1}</span>
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
    else if (d.add) c.rows.push(...blank(20, () => ({ p: "", q: "" })));
    else if (d.delrow) {
      c.rows.splice(+d.delrow, 1);
      if (c.rows.length < 20) c.rows.push({ p: "", q: "" });
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
    const cols = ["p", "q"];
    const c0 = cols.indexOf(k0);
    while (c.rows.length < r0 + lines.length) c.rows.push(...blank(20, () => ({ p: "", q: "" })));
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
  return { unit, q, sub: unit * q, incl: Math.round(p) * q };
}
function costTotal() {
  return T.cost.rows.reduce(
    (a, r) => {
      const x = costCalcRow(r);
      if (x) {
        a.sum += x.sub;
        a.incl += x.incl;
        a.n++;
        a.q += x.q;
      }
      return a;
    },
    { sum: 0, n: 0, q: 0, incl: 0 },
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
    ${
      T.cost.mode === "in"
        ? `<div class="sum minor"><small>税込み計</small><b>${yen(t.incl)}</b></div><div class="sum minor"><small>消費税計</small><b>${yen(t.incl - t.sum)}</b></div>`
        : `<div class="sum minor"><small>参考：税込み（×1.1）</small><b>${yen(Math.floor((t.sum * 11) / 10))}</b></div>`
    }`;
}

// 確認は出さずに消して、少しの間だけ「元に戻す」を出す
function clearRows(kind) {
  const key = kind === "sale" ? "sale" : kind === "cost" ? "cost" : "pad";
  const before = JSON.stringify(T[key]);
  if (key === "sale") T.sale.rows = blank(20, () => ({ l: "", c: "" }));
  else if (key === "cost") T.cost.rows = blank(20, () => ({ p: "", q: "" }));
  else T.pad = { pad: "", padCost: "", rotorCost: "", rotorList: "" };
  save();
  render($("#main"));
  C.toastAction("入力を全部消しました", "元に戻す", () => {
    T[key] = JSON.parse(before);
    save();
    if ($("#tbody")) render($("#main"));
  });
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
    <div class="card"><div class="card-h"><h3>入力</h3><button class="btn clear-all" data-padclear="1">入力を全部消す</button></div>
      ${f("pad", "パッド通常売価", "例：4,910")}
      ${f("padCost", "パッド原価", "例：3,510")}
      ${f("rotorCost", "ローター原価（1枚）", "例：3,290")}
      ${f("rotorList", "ローター通常売価（1枚・任意）", "お得額の表示用")}
      <div class="pad-set">セット価格：普通車 <b>${yen(sets.normal)}</b> ／ 軽 <b>${yen(sets.kei)}</b> ／ 原価割れ時のローター利益 <b>+${padMargin()}円</b> ／ 値引き時のパッド <b>+${padPlus()[0]}円・+${padPlus()[1]}円</b>
        <button class="linkbtn" data-padset="1">変更</button></div>
    </div>
    <div id="padres" class="pad-res"></div></div>
    <details class="pad-rule"><summary>計算のルール</summary><ol>
      <li>ローターを安くして、パッドで利益を取る</li><li>原価は割らない（値引きを使わない）</li>
      <li>ローター1枚 ＝（セット価格 − パッド通常売価）÷ 2。割り切れないときは、ローターを1円下げてパッドに回す</li>
      <li>ローターが原価割れするとき：ローター ＝ 原価＋${padMargin()}円、パッド ＝ セット価格 − ローター×2</li>
      <li>そのパッドがパッド原価を下回るときは「値引き必要」→ ローター＝原価＋${padMargin()}円、パッド＝原価＋${padPlus()[0]}円 / ＋${padPlus()[1]}円 の2案を出す</li></ol></details>`;
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
    if (b.dataset.padclear) clearRows("pad");
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
      ? `<div class="verdict ng"><span class="lamp r"></span><b>値引きが必要</b><small>ローターが原価割れ（${yen(x.under)}）。原価＋${padMargin()}円にしても、パッドが原価を ${yen(-x.padGain)} 下回ります。下の提案価格を使ってください</small></div>`
      : `<div class="verdict warn"><span class="lamp a"></span><b>パッドで調整</b><small>ローターが原価割れ（${yen(x.under)}）するので、ローター＝原価＋${padMargin()}円、残りをパッドに</small></div>`;
  return `<div class="pres ${x.ok ? "ok" : x.needDiscount ? "ng" : "warn"}"><h4>${label} <span class="mono">${yen(set)}</span></h4>${head}
    <table class="kv"><tr><td>パッド</td><td class="mono">${yen(x.pad)}</td><td class="mono g ${x.padGain < 0 ? "ng" : ""}">${x.padGain >= 0 ? "+" : ""}${yen(x.padGain)}</td></tr>
    <tr class="hl"><td>ローター 1枚</td><td class="mono">${yen(x.rotor)} <small>×2</small></td><td class="mono g">+${yen(x.rotorGain)}<small>/枚</small></td></tr>
    <tr class="tot"><td>セット合計</td><td class="mono">${yen(x.pad + x.rotor * 2)}</td><td class="mono g ${x.total < 0 ? "ng" : ""}">利益 ${yen(x.total)}</td></tr>
    ${normalTotal != null ? `<tr><td>通常で買うと</td><td class="mono">${yen(normalTotal)}</td><td class="mono save">${normalTotal > set ? `${yen(normalTotal - set)} お得` : ""}</td></tr>` : ""}</table>
    ${x.needDiscount ? padOffers(x, normalTotal) : ""}</div>`;
}
// セット価格では原価割れ → ローターは原価＋10円、パッドは原価＋500円 / ＋1000円で出す
function padOffers(x, normalTotal) {
  const padCost = x.pad - x.padGain;
  return `<div class="offers"><div class="ot">提案価格（ローター 原価＋${padMargin()}円）</div>${padPlus()
    .map((plus) => {
      const pad = padCost + plus;
      const total = pad + x.rotor * 2;
      return `<div class="offer"><div class="oh">パッド 原価＋${plus.toLocaleString()}円</div>
        <table class="kv"><tr><td>パッド</td><td class="mono">${yen(pad)}</td><td class="mono g">+${yen(plus)}</td></tr>
        <tr class="hl"><td>ローター 1枚</td><td class="mono">${yen(x.rotor)} <small>×2</small></td><td class="mono g">+${yen(padMargin())}<small>/枚</small></td></tr>
        <tr class="tot"><td>セット合計</td><td class="mono">${yen(total)}</td><td class="mono g">利益 ${yen(plus + padMargin() * 2)}</td></tr>
        ${normalTotal != null && normalTotal > total ? `<tr><td>通常で買うと</td><td class="mono">${yen(normalTotal)}</td><td class="mono save">${yen(normalTotal - total)} お得</td></tr>` : ""}</table></div>`;
    })
    .join("")}</div>`;
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
    <div class="field"><label>原価割れのときのローター利益（円）</label><input class="inp mono" name="margin" inputmode="numeric" value="${padMargin()}" required></div>
    <div class="grid2"><div class="field"><label>値引きが必要なとき パッド原価＋（案1）</label><input class="inp mono" name="plus1" inputmode="numeric" value="${padPlus()[0]}" required></div>
    <div class="field"><label>値引きが必要なとき パッド原価＋（案2）</label><input class="inp mono" name="plus2" inputmode="numeric" value="${padPlus()[1]}" required></div></div></form>
    <div class="modal-foot"><button class="btn" data-m="cancel">キャンセル</button><button class="btn pri" data-m="ok">保存</button></div>`);
  if (!r) return;
  const n = toNum(r.data.normal),
    k = toNum(r.data.kei),
    mg = toNum(r.data.margin),
    p1 = toNum(r.data.plus1),
    p2 = toNum(r.data.plus2);
  if (!n || !k || mg == null || !p1 || !p2) return C.toast("数字を入れてください", true);
  await C.store.update("settings", {
    padSetNormal: n,
    padSetKei: k,
    padMargin: mg,
    padPlus1: p1,
    padPlus2: p2,
  });
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
    _s: S.norm(`${b.spec} ${b.code} ${b.cat} ${b.shape} ${b.use} ${b.note}`),
    _v: voltOf(b),
    _sh: shapeKey(b.shape),
  }));
  return bulbIdx;
}
function renderBulb() {
  const box = fresh();
  box.innerHTML = `<label class="search">${C.IC.search}<input id="bq" type="search" autocomplete="off" placeholder="形状・規格・用途・品番（例：t10、h4、ウインカー、24v）" value="${esc(T.bulb.q)}"><span class="cnt" id="bcnt"></span></label>
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
      <div class="bmain"><div class="bt1"><span class="shape">${esc(b.shape || "?")}</span><b>${esc(prettySpec(b.spec))}</b><span class="cat">${esc(b.cat)}</span>${b.note ? `<span class="hot">${b.note.includes("一番") ? "★" : ""}${esc(b.note)}</span>` : ""}</div>
        <div class="buse">${esc(b.use)}</div></div>
      <div class="bside"><button class="bcode" data-bcopy="${esc(b.code)}" title="品番をコピー">${esc(b.code)}</button><small>${b.unit ? `${esc(b.unit)}個入` : ""}</small>
        <div class="bbtn"><button class="btn sm" data-bgoogle="${esc(b.code)}">Google</button><button class="btn sm" data-bedit="${esc(b.id)}">編集</button></div></div></div>`,
      )
      .join("") || '<p class="empty">該当なし</p>';
}
async function bulbForm(id) {
  const b = id ? C.st.bulbs[id] : { spec: "", code: "", cat: "", shape: "", use: "", unit: "", note: "" };
  if (!b) return;
  const f = (k, l, extra = "", cls = "") =>
    `<div class="field"><label>${l}</label><input class="inp ${cls}" name="${k}" value="${esc(b[k])}" ${extra}></div>`;
  const r = await C.modal(`<h3>${id ? "バルブを編集" : "バルブを追加"}</h3><form>
    <div class="grid2">${f("code", "品番", "required", "mono")}${f("unit", "販売単位（個入）", 'inputmode="numeric"')}</div>
    <div class="grid2">${f("shape", "形状（T10・H4 など）", "required")}${f("spec", "規格（12V 5W など）")}</div>
    <div class="grid2">${f("cat", "分類（ウェッジ・シングル球 など）")}${f("note", "メモ（よく出る、など）")}</div>
    ${f("use", "用途・呼び方")}</form>
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
  for (const k of ["code", "shape", "spec", "cat", "use", "note"]) rec[k] = N(r.data[k]);
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
  const units = {}; // 2枚目のシートからは販売単位だけ使う
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
          if (c && iUnit >= 0) units[c] = toNum(r[iUnit]) ?? "";
        }
        break;
      }
    }
  }
  if (!base) throw new Error("「規格」「品番」「形状」の列が見つかりませんでした");
  return base.map((b) => ({ ...b, unit: units[b.code] ?? "" }));
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
  const ok = await C.confirmBox(
    `${list.length}件 のバルブを取り込みます。${have ? `\n今の ${have}件 は入れ替わります。` : ""}`,
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
      unit: b.unit ?? "",
      note: b.note || "",
    };
  });
  await C.store.set("bulbs", obj);
  C.toast(`${list.length}件 取り込みました`);
}


// ================= オイル・ケミカル検索 =================
const oilN = (s) =>
  String(s ?? "")
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
// ひらがなで打っても漢字の名称に当たるように（設定の「同義語」にも足せます）
const OIL_SYN_RAW = [
  "液状=えきじょう",
  "パッキン=ぱっきん=ガスケット",
  "耐熱=たいねつ",
  "潤滑=じゅんかつ",
  "洗浄=せんじょう",
  "防錆=ぼうせい=錆止め=さびどめ",
  "固着=こちゃく",
  "接着=せっちゃく",
  "冷却水=れいきゃくすい=LLC=クーラント",
  "添加剤=てんかざい",
  "洗剤=せんざい",
  "剥離=はくり",
  "鉱物油=こうぶつゆ",
  "化学合成油=かがくごうせいゆ",
  "部分合成油=ぶぶんごうせいゆ",
  "軽油=けいゆ",
  "耐水=たいすい",
  "共用=きょうよう",
  "塗料=とりょう",
  "消臭=しょうしゅう",
  "艶出し=つやだし",
].map((l) => l.split("="));
const OIL_SYN = OIL_SYN_RAW.map((g) => g.map(S.norm));
// 入力途中（「えきじ」など）でも、読みの前方一致で漢字の名称を拾う
const KANA_ONLY = /^[ァ-ヶー]+$/;
function oilExpand(tok) {
  const alts = new Set(S.expand(tok, [...OIL_SYN, ...C.st.syn]));
  const t = S.norm(tok);
  if (KANA_ONLY.test(t) && t.length >= 2)
    for (const g of [...OIL_SYN, ...C.st.syn])
      if (g.some((w) => KANA_ONLY.test(w) && w.length > t.length && w.startsWith(t))) for (const o of g) alts.add(o);
  return [...alts];
}
let oilIdx = null,
  oilSrc = null;
function oilList() {
  const raw = C.st.oils || {};
  if (raw === oilSrc && oilIdx) return oilIdx;
  oilSrc = raw;
  oilIdx = Object.entries(raw).map(([id, o]) => ({
    id,
    ...o,
    _s: S.norm(
      `${o.code} ${String(o.code || "").replace(/[^0-9A-Za-z]/g, "")} ${o.name} ${o.use} ${o.cat} ${o.cap}`,
    ),
    _c: S.norm(String(o.code || "").replace(/[^0-9A-Za-z]/g, "")),
  }));
  return oilIdx;
}
// 分類は登録済みの商品から集める（多い順）
function oilCats(list = oilList()) {
  const cnt = {};
  for (const o of list) if (o.cat) cnt[o.cat] = (cnt[o.cat] || 0) + 1;
  return Object.keys(cnt)
    .sort((a, b) => cnt[b] - cnt[a] || a.localeCompare(b, "ja"))
    .map((name) => ({ name, n: cnt[name] }));
}
// 写真本体は oilPhoto/{id}、oils/{id}/photo には更新時刻だけ
const oilPhotoCache = new Map();
function oilPhoto(id) {
  const o = C.st.oils[id];
  if (!o?.photo) return Promise.resolve(null);
  const hit = oilPhotoCache.get(id);
  if (hit && hit.v === o.photo) return hit.p;
  const p = C.store.get(`oilPhoto/${id}`).catch(() => null);
  oilPhotoCache.set(id, { v: o.photo, p });
  return p;
}
function fillOilPhotos(root) {
  $$("img[data-ophoto]", root).forEach(async (img) => {
    const u = await oilPhoto(img.dataset.ophoto);
    if (u) img.src = u;
    else img.closest(".oph")?.remove();
  });
}
let oilLimit = 30;
function renderOil() {
  const box = fresh();
  oilLimit = 30;
  box.innerHTML = `<label class="search">${C.IC.search}<input id="oq" type="search" autocomplete="off" placeholder="品番・名称・用途（例：5w30、ATF、ブレーキ）" value="${esc(T.oil.q)}"><span class="cnt" id="ocnt"></span></label>
    <div class="chips" id="ochips"></div>
    <div id="olist"></div>
    <div class="tacts"><input type="file" id="foil" accept=".xlsx,.xlsm,.xls,.csv" hidden><button class="btn" data-opick="1">Excelから取り込み</button><button class="btn" data-ocats="1">分類を整理</button><button class="btn" data-oxlsx="1" title="検索機能つきのExcelを作る">Excelに書き出し</button><button class="chip ${T.oil.protect !== false ? "on" : ""}" data-oprot="1" title="書き出すExcelの検索・使い方シートに保護をかける（パスワードなし）">シート保護 ${T.oil.protect !== false ? "ON" : "OFF"}</button><input type="file" id="foilbase" accept=".xlsx,.xlsm" hidden><button class="btn" data-oxbase="1" title="会社のExcelに検索シートを足す">会社のExcelに追加</button><button class="btn pri" data-oedit="">＋ 追加</button></div>`;
  const q = $("#oq", box);
  q.addEventListener("input", () => {
    T.oil.q = q.value;
    oilLimit = 30;
    save();
    renderOilList();
  });
  q.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      q.value = T.oil.q = "";
      renderOilList();
    }
  });
  $("#foil", box).addEventListener("change", (e) =>
    importOils(e.target.files[0]).finally(() => (e.target.value = "")),
  );
  $("#foilbase", box).addEventListener("change", (e) =>
    exportIntoBase(e.target.files[0]).finally(() => (e.target.value = "")),
  );
  box.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const d = b.dataset;
    const o = d.oid ? C.st.oils[d.oid] && { id: d.oid, ...C.st.oils[d.oid] } : null;
    if (d.opick) $("#foil").click();
    else if (d.oedit !== undefined) oilForm(d.oedit || null);
    else if (d.ocats) oilCatsForm();
    else if (d.oxlsx) exportOils();
    else if (d.oxbase) $("#foilbase").click();
    else if (d.oprot) {
      T.oil.protect = T.oil.protect === false;
      save();
      b.classList.toggle("on", T.oil.protect);
      b.textContent = `シート保護 ${T.oil.protect ? "ON" : "OFF"}`;
    }
    else if (d.ocat !== undefined) {
      T.oil.cat = T.oil.cat === d.ocat ? "" : d.ocat;
      oilLimit = 30;
      save();
      renderOilList();
    } else if (d.omore) {
      oilLimit += 60;
      renderOilList();
    } else if (d.ocopy) C.copyText(d.ocopy).then(() => C.toast(`${d.ocopy} をコピーしました`));
    else if (o && d.oact === "google") C.googleOpen(`${o.code} ${o.name}`);
    else if (o && d.oact === "memo") C.openMemoForOil(o);
    else if (o && d.oact === "edit") oilForm(o.id);
    else if (o && d.oact === "view") oilView(o.id);
  });
  renderOilList();
  if (matchMedia("(hover: hover)").matches) setTimeout(() => q.focus(), 0);
}
function renderOilList() {
  const list = oilList();
  const box = $("#olist");
  if (!box) return;
  const cats = oilCats(list);
  if (T.oil.cat && T.oil.cat !== "__none" && !cats.some((c) => c.name === T.oil.cat)) T.oil.cat = "";
  const noCat = list.filter((o) => !o.cat).length;
  $("#ochips").innerHTML =
    cats
      .map(
        (c) =>
          `<button class="chip ${T.oil.cat === c.name ? "on" : ""}" data-ocat="${esc(c.name)}">${esc(c.name)} ${c.n}</button>`,
      )
      .join("") +
    (noCat && cats.length
      ? `<button class="chip ${T.oil.cat === "__none" ? "on" : ""}" data-ocat="__none">未分類 ${noCat}</button>`
      : "");
  if (!list.length) {
    $("#ocnt").textContent = "";
    box.innerHTML = `<div class="panel"><h3>オイル・ケミカルがまだ登録されていません</h3><p style="margin:0;font-size:12.5px;color:var(--mu);line-height:1.7">下の「＋ 追加」で品番・名称・用途を入れて登録するか、「Excelから取り込み」で品番・名称（・容量）の入ったExcel／CSVを選んでください。列の並びはどんな形でも自動で探します。</p></div>`;
    return;
  }
  const toks = String(T.oil.q || "")
    .normalize("NFKC")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map(oilExpand);
  const qn = S.norm(String(T.oil.q || "").replace(/[^0-9A-Za-z]/g, ""));
  const hits = list.filter(
    (o) =>
      (!T.oil.cat || (T.oil.cat === "__none" ? !o.cat : o.cat === T.oil.cat)) &&
      toks.every((alts) => alts.some((a) => o._s.includes(a))),
  );
  // 品番が前方一致するものを先頭に
  hits.sort(
    (a, b) =>
      (qn && b._c.startsWith(qn) ? 1 : 0) - (qn && a._c.startsWith(qn) ? 1 : 0) ||
      String(a.name).localeCompare(String(b.name), "ja", { numeric: true }),
  );
  $("#ocnt").textContent = `${hits.length} / ${list.length}件`;
  const shown = hits.slice(0, oilLimit);
  const memoCnt = {};
  for (const m of Object.values(C.st.memos || {})) if (m.oilId) memoCnt[m.oilId] = (memoCnt[m.oilId] || 0) + 1;
  box.innerHTML =
    shown
      .map(
        (o) => `<div class="orow">
      ${o.photo ? `<button class="oph" data-oact="view" data-oid="${esc(o.id)}" title="写真を大きく見る"><img data-ophoto="${esc(o.id)}" alt=""></button>` : ""}
      <div class="omain">
        <div class="ot1"><b class="oname">${esc(o.name)}</b>${o.cap ? `<span class="ocap">${esc(o.cap)}</span>` : ""}${o.cat ? `<span class="cat">${esc(o.cat)}</span>` : ""}</div>
        <div class="ot2"><button class="bcode" data-ocopy="${esc(o.code)}" title="品番をコピー">${esc(o.code)}</button></div>
        ${o.use ? `<div class="buse">${esc(o.use)}</div>` : ""}
        <div class="oacts"><button class="btn" data-oact="google" data-oid="${esc(o.id)}" title="品番と名称でGoogle検索">${C.IC.search}Google</button><button class="btn ${memoCnt[o.id] ? "has" : ""}" data-oact="memo" data-oid="${esc(o.id)}">${C.IC.note}メモ${memoCnt[o.id] ? `<span class="cnt">${memoCnt[o.id]}</span>` : ""}</button><button class="btn sm" data-oact="edit" data-oid="${esc(o.id)}">編集</button></div>
      </div></div>`,
      )
      .join("") +
    (hits.length > shown.length
      ? `<button class="btn more-btn" data-omore="1">さらに表示（残り ${hits.length - shown.length}件）</button>`
      : "") || '<p class="empty">該当なし</p>';
  fillOilPhotos(box);
}
async function oilView(id) {
  const o = C.st.oils[id];
  const u = await oilPhoto(id);
  if (!o || !u) return;
  await C.modal(`<h3>${esc(o.name)} <span class="mono" style="color:var(--mu);font-size:14px">${esc(o.code)}</span></h3>
    <img src="${u}" alt="" style="width:100%;border-radius:8px;display:block">
    <div class="modal-foot"><button class="btn pri" data-m="cancel">閉じる</button></div>`);
}
async function oilForm(id) {
  const o = id ? C.st.oils[id] : null;
  if (id && !o) return;
  const b = o || { code: "", name: "", cap: "", use: "", cat: "" };
  let photo = null,
    changed = false;
  const f = (k, l, extra = "", cls = "") =>
    `<div class="field"><label>${l}</label><input class="inp ${cls}" name="${k}" value="${esc(b[k])}" ${extra}></div>`;
  const r = await C.modal(
    `<h3>${id ? "オイル・ケミカルを編集" : "オイル・ケミカルを追加"}</h3><form>
    ${f("code", "品番", "required autocomplete=off", "mono")}${f("name", "名称", "required autocomplete=off")}
    <div class="grid2">${f("cap", "容量（任意）", 'placeholder="4L" autocomplete=off')}${f("use", "使用用途（任意）", "autocomplete=off")}</div>
    <div class="field"><label>分類（任意）</label><input class="inp" name="cat" id="of-cat" value="${esc(b.cat)}" placeholder="下から選ぶか、新しい名前を入力" autocomplete="off">
      <div class="chips" id="of-cats" style="margin:6px 0 0">${oilCats()
        .map(
          (c) =>
            `<button type="button" class="chip ${c.name === b.cat ? "on" : ""}" data-pcat="${esc(c.name)}">${esc(c.name)}</button>`,
        )
        .join("")}</div></div>
    <div class="field"><label>商品の写真（任意）</label><div id="of-photo"></div></div></form>
    <div class="modal-foot">${id ? '<button class="btn danger left" data-m="delete">削除</button>' : ""}<button class="btn" data-m="cancel">キャンセル</button>${id ? "" : '<button class="btn" data-m="next">保存して次へ</button>'}<button class="btn pri" data-m="ok">保存</button></div>`,
    {
      onOpen(bg) {
        const wrap = $("#of-photo", bg);
        const draw = () => {
          wrap.innerHTML = C.photoBox(photo, "商品の写真（任意）");
          C.wirePhotoBox($(".kdrop", wrap), (url) => {
            photo = url;
            changed = true;
            draw();
          });
        };
        draw();
        if (id) oilPhoto(id).then((u) => !changed && u && ((photo = u), draw()));
        bg.addEventListener("paste", async (e) => {
          const file = C.pasteFile(e);
          if (!file) return;
          e.preventDefault();
          try {
            photo = await C.blobToPhoto(file);
            changed = true;
            draw();
          } catch (err) {
            C.toast(err.message, true);
          }
        });
        bg.addEventListener("click", (e) => {
          const c = e.target.closest("[data-pcat]");
          if (!c) return;
          const inp = $("#of-cat", bg);
          inp.value = inp.value === c.dataset.pcat ? "" : c.dataset.pcat;
          $$("#of-cats .chip", bg).forEach((x) => x.classList.toggle("on", x.dataset.pcat === inp.value));
        });
        $("#of-cat", bg).addEventListener("input", (e) =>
          $$("#of-cats .chip", bg).forEach((x) => x.classList.toggle("on", x.dataset.pcat === oilN(e.target.value))),
        );
      },
    },
  );
  if (!r) return;
  if (r.act === "delete") {
    if (!(await C.confirmBox(`${b.code}（${b.name}）を削除しますか？`, "削除", true))) return;
    await C.store.remove(`oils/${id}`);
    await C.store.remove(`oilPhoto/${id}`).catch(() => {});
    oilPhotoCache.delete(id);
    return C.toast("削除しました");
  }
  const rec = {
    code: oilN(r.data.code).toUpperCase(),
    name: oilN(r.data.name),
    cap: oilN(r.data.cap),
    use: oilN(r.data.use),
    cat: oilN(r.data.cat),
    photo: o?.photo ?? null,
    updatedAt: Date.now(),
  };
  if (!rec.code || !rec.name) {
    C.toast("品番と名称を入れてください", true);
    return oilForm(id);
  }
  if (
    !id &&
    oilList().some((x) => String(x.code).toUpperCase() === rec.code) &&
    !(await C.confirmBox(`品番 ${rec.code} はすでに登録されています。\nもうひとつ追加しますか？`, "追加する"))
  )
    return;
  const key = id || C.localKey();
  if (changed) {
    rec.photo = photo ? Date.now() : null;
    if (photo) await C.store.set(`oilPhoto/${key}`, photo);
    else await C.store.remove(`oilPhoto/${key}`).catch(() => {});
    oilPhotoCache.delete(key);
  }
  await C.store.set(`oils/${key}`, rec);
  C.toast("保存しました");
  if (r.act === "next") oilForm(null);
}
// 分類の名前を直す・消す（商品は消えません）
async function oilCatsForm() {
  const cats = oilCats();
  if (!cats.length) return C.toast("分類はまだありません。商品の登録時に入力すると増えます");
  const r = await C.modal(`<h3>分類を整理</h3><p style="margin:0 0 10px;font-size:12px;color:var(--mu);line-height:1.6">名前を直すと、その分類の商品すべてに反映されます。同じ名前にすると1つにまとまります。削除しても商品は残り「未分類」になります。</p><form>
    ${cats
      .map(
        (c, i) => `<div class="catrow"><input class="inp" name="r${i}" value="${esc(c.name)}"><small>${c.n}件</small>
      <label class="switch"><input type="checkbox" name="d${i}"> 削除</label></div>`,
      )
      .join("")}</form>
    <div class="modal-foot"><button class="btn" data-m="cancel">キャンセル</button><button class="btn pri" data-m="ok">保存</button></div>`);
  if (!r) return;
  const upd = {};
  let n = 0;
  cats.forEach((c, i) => {
    const del = r.data[`d${i}`] === "on";
    const to = del ? "" : oilN(r.data[`r${i}`]);
    if (to === c.name) return;
    for (const o of oilList()) if (o.cat === c.name) (upd[`${o.id}/cat`] = to), n++;
  });
  if (!n) return;
  await C.store.update("oils", upd);
  C.toast("分類を更新しました");
}

// ---- Excel / CSV の取り込み：列の並びがどうでも、品番・名称・容量を自動で探す ----
const hdKey = (s) =>
  oilN(s)
    .toLowerCase()
    .replace(/[\s・･_\-－ー/／.()（）［］\[\]]/g, "");
// [完全一致, 部分一致, 除外] ／ 完全一致が最優先
const OIL_COL = {
  code: [
    ["品番", "部品番号", "商品コード", "商品番号", "品目コード", "品目番号", "製品番号", "製品コード", "型番", "コード", "品番コード", "pn", "partno", "partnumber", "itemcode", "sku", "アイテムコード", "メーカー品番", "商品cd", "品目cd"],
    ["品番", "コード", "番号", "partno", "品目"],
    /jan|ean|バーコード|棚|ロケ|価格|単価|在庫|数量|区分|分類|仕入|取引/,
  ],
  name: [
    ["名称", "品名", "商品名", "製品名", "品目名", "部品名", "名前", "商品名称", "品名規格", "名称規格", "品名称", "description", "name", "itemname", "商品"],
    ["名称", "品名", "商品名", "製品名", "名"],
    /メーカー|分類|コード|番号|価格|単価|在庫|数量|区分|仕入|取引|カナ|ｶﾅ|略/,
  ],
  cap: [
    ["容量", "内容量", "規格", "サイズ", "荷姿", "容器", "容量規格", "volume", "size", "リットル"],
    ["容量", "内容量", "規格", "サイズ", "荷姿"],
    /価格|単価|在庫|数量|分類/,
  ],
  use: [
    ["用途", "使用用途", "適合", "適用", "特徴", "使用箇所", "使い方", "用途特徴", "対象", "対応"],
    ["用途", "適合", "適用", "使用", "特徴"],
    /価格|単価|在庫|数量/,
  ],
  cat: [
    ["分類", "カテゴリ", "カテゴリー", "区分", "種別", "種類", "品種", "大分類", "中分類", "小分類", "ジャンル", "category", "グループ"],
    ["分類", "カテゴリ", "区分", "種別", "品種", "ジャンル"],
    /^$/,
  ],
};
function scoreHeader(h, field) {
  const k = hdKey(h);
  if (!k) return 0;
  const [exact, part, ng] = OIL_COL[field];
  if (exact.some((e) => hdKey(e) === k)) return 3;
  if (ng.test(k)) return 0;
  return part.some((p) => k.includes(hdKey(p))) ? 1 : 0;
}
const colName = (i) => {
  let s = "",
    n = i + 1;
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};
const CAP_RE = /(\d+(?:\.\d+)?)\s*(?:ℓ|l|ml|ｍｌ|リットル|kg|g|㎖)(?![a-z])/i;
const CAP_ONLY = /^\d+(?:\.\d+)?\s*(?:ℓ|l|ml|リットル|kg|g|㎖|缶|本|個|セット|入)?(?:\s*[×x*]\s*\d+\s*(?:缶|本|個|入)?)?$/i;
const JP = /[ぁ-んァ-ヶ一-龠]/;
const looksCode = (v) => /^[A-Za-z0-9][A-Za-z0-9\-_./ ]{2,29}$/.test(v) && /\d/.test(v) && !CAP_ONLY.test(v);

// 1枚のシートから商品を探す。見つからなければ null
function oilsFromSheet(rows) {
  const W = Math.max(0, ...rows.slice(0, 60).map((r) => (r || []).length));
  if (!W) return null;
  // 1) 見出し行を探す（上から40行）：品番・名称・容量・用途・分類の見出しが多い行
  let best = null;
  for (let h = 0; h < Math.min(rows.length, 40); h++) {
    const row = rows[h] || [];
    const pick = {};
    const used = new Set();
    // 強い一致の項目から順に、列を取り合わないよう割り当てる
    const cand = [];
    for (const f of Object.keys(OIL_COL))
      row.forEach((c, i) => {
        const sc = scoreHeader(c, f);
        if (sc) cand.push({ f, i, sc });
      });
    cand.sort((a, b) => b.sc - a.sc || a.i - b.i);
    for (const c of cand) if (!(c.f in pick) && !used.has(c.i)) ((pick[c.f] = c.i), used.add(c.i));
    const n = Object.keys(pick).length;
    const key = (pick.code != null ? 3 : 0) + (pick.name != null ? 3 : 0) + n;
    if ((pick.code != null || pick.name != null) && (!best || key > best.key)) best = { h, pick, key, hdr: row };
  }
  let h = best ? best.h : -1;
  const pick = best ? { ...best.pick } : {};
  const how = {};
  for (const f in pick) how[f] = `${colName(pick[f])}列「${oilN(best.hdr[pick[f]])}」`;
  const data = rows.slice(h + 1).filter((r) => (r || []).some((c) => oilN(c)));
  const sample = data.slice(0, 300);
  const col = (i) => sample.map((r) => oilN(r[i])).filter(Boolean);
  const taken = () => new Set(Object.values(pick));
  // 2) 見出しで見つからなかった列は、中身から推測する
  if (pick.code == null) {
    let bi = -1,
      bs = 0.55;
    for (let i = 0; i < W; i++) {
      if (taken().has(i)) continue;
      const v = col(i);
      if (v.length < 2) continue;
      const sc = (v.filter(looksCode).length / v.length) * (new Set(v).size / v.length);
      if (sc > bs) ((bs = sc), (bi = i));
    }
    if (bi >= 0) ((pick.code = bi), (how.code = `${colName(bi)}列（中身から判定）`));
  }
  if (pick.name == null) {
    let bi = -1,
      bs = 0;
    for (let i = 0; i < W; i++) {
      if (taken().has(i)) continue;
      const v = col(i);
      if (v.length < 2) continue;
      const jp = v.filter((x) => JP.test(x) && x.length >= 2).length / v.length;
      const len = v.reduce((a, x) => a + x.length, 0) / v.length;
      const sc = jp * Math.min(len, 30);
      if (jp >= 0.5 && len >= 3 && sc > bs) ((bs = sc), (bi = i));
    }
    if (bi >= 0) ((pick.name = bi), (how.name = `${colName(bi)}列（中身から判定）`));
  }
  if (pick.cap == null) {
    let bi = -1,
      bs = 0.5;
    for (let i = 0; i < W; i++) {
      if (taken().has(i)) continue;
      const v = col(i);
      if (v.length < 2) continue;
      const sc = v.filter((x) => x.length <= 14 && (CAP_ONLY.test(x) || CAP_RE.test(x))).length / v.length;
      if (sc > bs) ((bs = sc), (bi = i));
    }
    if (bi >= 0) ((pick.cap = bi), (how.cap = `${colName(bi)}列（中身から判定）`));
  }
  if (pick.code == null || pick.name == null) return null;
  const items = [];
  for (const r of data) {
    const code = oilN(r[pick.code]).toUpperCase();
    const name = oilN(r[pick.name]);
    if (!code || !name || code === name) continue;
    if (/^(合計|小計|計|total|品番|部品番号|商品コード)$/i.test(code)) continue;
    if (!/[0-9A-Za-z]/.test(code)) continue;
    const it = { code, name };
    if (pick.cap != null) it.cap = oilN(r[pick.cap]);
    if (!it.cap) {
      const m = name.match(CAP_RE);
      if (m) it.cap = m[0].replace(/\s+/g, "");
    }
    if (pick.use != null) it.use = oilN(r[pick.use]);
    if (pick.cat != null) it.cat = oilN(r[pick.cat]);
    items.push(it);
  }
  return items.length ? { items, how, hasCat: pick.cat != null } : null;
}
export function oilsFromWorkbook(named) {
  const found = [];
  for (const [name, rows] of named) {
    const r = oilsFromSheet(rows);
    if (r) found.push({ name, ...r });
  }
  if (!found.length) throw new Error("品番と名称の列が見つかりませんでした");
  // シートが複数あって分類の列が無いときは、シート名を分類にする（Sheet1 などは除く）
  const multi = found.length > 1;
  const byCode = new Map();
  for (const f of found)
    for (const it of f.items) {
      const x = { ...it };
      if (multi && !f.hasCat && !/^(sheet|シート)\s*\d*$/i.test(f.name.trim())) x.cat = f.name.trim();
      byCode.set(x.code, x); // 同じ品番は後のほうを使う
    }
  return { items: [...byCode.values()], sheets: found.map((f) => ({ name: f.name, n: f.items.length, how: f.how })) };
}
async function importOils(file) {
  if (!file) return;
  let res;
  try {
    const X = await C.loadXlsx();
    let wb;
    if (/\.csv$/i.test(file.name)) {
      const buf = await file.arrayBuffer();
      let txt;
      try {
        txt = new TextDecoder("utf-8", { fatal: true }).decode(buf);
      } catch {
        txt = new TextDecoder("shift_jis").decode(buf); // Excelで保存したCSVはShift-JIS
      }
      wb = X.read(txt, { type: "string" });
    } else wb = X.read(await file.arrayBuffer(), { type: "array" });
    res = oilsFromWorkbook(
      wb.SheetNames.map((n) => [n, X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: "" })]),
    );
  } catch (e) {
    return C.toast("取り込めませんでした：" + e.message, true);
  }
  const have = new Map(oilList().map((o) => [String(o.code).toUpperCase(), o]));
  const add = res.items.filter((x) => !have.has(x.code)).length;
  const upd = res.items.length - add;
  const L = { code: "品番", name: "名称", cap: "容量", use: "用途", cat: "分類" };
  const detail = res.sheets
    .map(
      (s) =>
        `■ ${s.name}（${s.n}件）\n` +
        Object.keys(L)
          .filter((k) => s.how[k])
          .map((k) => `　${L[k]} ← ${s.how[k]}`)
          .join("\n"),
    )
    .join("\n");
  const ex = res.items
    .slice(0, 3)
    .map((x) => `　${x.code}　${x.name}${x.cap ? `　${x.cap}` : ""}`)
    .join("\n");
  const ok = await C.confirmBox(
    `${res.items.length}件 を見つけました（新規 ${add}件／更新 ${upd}件）。\n今ある商品は消えません。同じ品番は上書きします。\n\n${detail}\n\n例：\n${ex}`,
    "取り込む",
  );
  if (!ok) return;
  const w = {};
  for (const it of res.items) {
    const old = have.get(it.code);
    if (old) {
      w[`${old.id}/name`] = it.name;
      if (it.cap) w[`${old.id}/cap`] = it.cap;
      if (it.use) w[`${old.id}/use`] = it.use;
      if (it.cat) w[`${old.id}/cat`] = it.cat;
      w[`${old.id}/updatedAt`] = Date.now();
    } else
      w[C.localKey()] = {
        code: it.code,
        name: it.name,
        cap: it.cap || "",
        use: it.use || "",
        cat: it.cat || "",
        updatedAt: Date.now(),
      };
  }
  const keys = Object.keys(w);
  try {
    for (let i = 0; i < keys.length; i += 500) {
      const part = {};
      for (const k of keys.slice(i, i + 500)) part[k] = w[k];
      await C.store.update("oils", part);
    }
    C.toast(`${res.items.length}件 取り込みました`);
  } catch (e) {
    C.toast("途中で失敗しました：" + e.message + "（もう一度取り込むと続きから入ります）", true);
  }
}

// ---- 検索機能つきExcelの書き出し（Excel 2016 の関数だけで動く） ----
const libPromises = {};
function loadLib(key, urls) {
  if (window[key]) return Promise.resolve(window[key]);
  if (!libPromises[key])
    libPromises[key] = new Promise((res, rej) => {
      const list = [...urls];
      const next = () => {
        const u = list.shift();
        if (!u) {
          libPromises[key] = null;
          return rej(new Error("Excel用の部品を読み込めませんでした。通信を確認してください"));
        }
        const el = document.createElement("script");
        el.src = u;
        el.onload = () => (window[key] ? res(window[key]) : next());
        el.onerror = next;
        document.head.appendChild(el);
      };
      next();
    });
  return libPromises[key];
}
const loadExcelJs = () =>
  loadLib("ExcelJS", [
    "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
  ]);
const loadJsZip = () =>
  loadLib("JSZip", [
    "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
  ]);
function oilExportItems() {
  return [...oilList()]
    .sort(
      (a, b) =>
        String(a.cat || "").localeCompare(String(b.cat || ""), "ja") ||
        String(a.name).localeCompare(String(b.name), "ja", { numeric: true }),
    )
    .map((o) => ({ code: o.code, name: o.name, cap: o.cap || "", use: o.use || "", cat: o.cat || "" }));
}
function downloadXlsx(buf, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(
    new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
  );
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
const oilGroups = () => [...OIL_SYN_RAW, ...(C.st.syn || [])];
async function exportOils() {
  if (!oilList().length) return C.toast("書き出す商品がまだありません", true);
  try {
    C.toast("Excelを作っています…");
    const X = await loadExcelJs();
    const items = oilExportItems();
    const { wb } = buildOilWorkbook(X, items, oilGroups(), { protect: T.oil.protect !== false });
    downloadXlsx(await wb.xlsx.writeBuffer(), "オイル・ケミカル検索.xlsx");
    C.toast(`${items.length}件 をExcelにしました`);
  } catch (e) {
    C.toast("Excelを作れませんでした：" + e.message, true);
  }
}
// 会社指定のExcelを選ぶ → 内容を確認 → 実行で、そのExcelに検索シートを足して保存
// 元ファイルの中身には手を触れず、シートだけを差し込む（グラフ・表・図・書式はそのまま）
async function exportIntoBase(file) {
  if (!file) return;
  if (!oilList().length) return C.toast("書き出す商品がまだありません", true);
  const ext = (file.name.match(/\.(xlsx|xlsm)$/i) || [])[1];
  if (!ext)
    return C.toast("対応しているのは .xlsx（.xlsm）です。古い .xls は、Excelで「.xlsx」として保存し直してください", true);
  let JSZip, base;
  try {
    C.toast("Excelを読んでいます…");
    JSZip = await loadJsZip();
    base = await readBase(JSZip, await file.arrayBuffer());
  } catch (e) {
    return C.toast("このExcelを開けませんでした：" + e.message, true);
  }
  const items = oilExportItems();
  const nameS = uniqueName(base.names, "オイル・ケミカル検索");
  const nameG = uniqueName(base.names, "使い方");
  const nameD = uniqueName(base.names, "オイル・ケミカルデータ");
  const r = await C.modal(`<h3>Excelに検索機能を追加</h3>
    <div class="xlinfo">
      <div><span>ファイル</span><b>${esc(file.name)}</b></div>
      <div><span>作成者</span><b>${esc(base.creator || "（記録なし）")}</b></div>
      <div><span>最終更新者</span><b>${esc(base.lastModifiedBy || "（記録なし）")}</b></div>
      ${base.company ? `<div><span>会社名</span><b>${esc(base.company)}</b></div>` : ""}
      <div><span>今あるシート</span><b>${esc(base.names.join("、") || "（なし）")}</b></div>
      <div><span>追加するシート</span><b>${esc(nameS)}（検索画面）・${esc(nameG)}・${esc(nameD)}（${items.length}件）</b></div>
      <div><span>シート保護</span><b>${T.oil.protect !== false ? "あり（検索・使い方のシート。パスワードなし）" : "なし"}</b></div>
    </div>
    <p style="margin:10px 0 0;font-size:12px;color:var(--mu);line-height:1.7">・今ある内容（シート・グラフ・表・書式など）には手を触れず、検索のシートだけを後ろに足します。<br>・作成者などの情報は、このファイルのままです。<br>・元のファイルは変えず、別のファイルとして保存します。</p>
    <div class="modal-foot"><button class="btn" data-m="cancel">やめる</button><button class="btn pri" data-m="ok" autofocus>実行</button></div>`);
  if (!r) return;
  try {
    C.toast("Excelを作っています…");
    const X = await loadExcelJs();
    const { wb } = buildOilWorkbook(X, items, oilGroups(), {
      names: { search: nameS, guide: nameG, data: nameD },
      protect: T.oil.protect !== false,
    });
    const out = await injectSheets(JSZip, base, await wb.xlsx.writeBuffer(), { nameS, nameG, nameD });
    downloadXlsx(out, `オイル・ケミカル検索.${ext.toLowerCase()}`);
    C.toast("保存しました。開いたら「編集を有効にする」を押してください");
  } catch (e) {
    C.toast("Excelを作れませんでした：" + e.message, true);
  }
}
