// 品番辞書（3万点くらい）
// DBには catalog/{品番キー} で1件ずつ置く。端末側はIndexedDBに丸ごと控えておいて、
// catalogMeta/version が変わったときだけ取り直す
import * as S from "./search.js?v=20261003f";

// ---- IndexedDB（だめなら控えなしで動く） ----
let dbp = null;
function idb() {
  if (dbp) return dbp;
  dbp = new Promise((resolve) => {
    try {
      const req = indexedDB.open("secsearch", 1);
      req.onupgradeneeded = () => req.result.createObjectStore("kv");
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return dbp;
}
export async function cacheGet(k) {
  const db = await idb();
  if (!db) return null;
  return new Promise((resolve) => {
    const r = db.transaction("kv").objectStore("kv").get(k);
    r.onsuccess = () => resolve(r.result ?? null);
    r.onerror = () => resolve(null);
  });
}
export async function cacheSet(k, v) {
  const db = await idb();
  if (!db) return;
  return new Promise((resolve) => {
    const tx = db.transaction("kv", "readwrite");
    tx.objectStore("kv").put(v, k);
    tx.oncomplete = () => resolve();
    tx.onerror = () => resolve();
  });
}

// ---- 正規化 ----
// キーは英数字だけ（KLAPC-05320(PS) → KLAPC05320PS）
export const catKey = (code) =>
  String(code ?? "")
    .normalize("NFKC")
    .toUpperCase()
    .replace(/[^0-9A-Z]/g, "");

// 半角カナ→全角、全角英数→半角、空白をそろえる
export const cleanText = (s) =>
  String(s ?? "")
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();

// 部品名用。古い半角データの「ｱﾌﾞｿ-ﾊﾞ-」の - を長音にする
export const cleanName = (s) => cleanText(s).replace(/(?<=[ァ-ヶー])-/g, "ー");

// 「ボルト,フランジ」→「フランジボルト」
function flipped(name) {
  const parts = name
    .split(/[,，、]/)
    .map((x) => x.trim())
    .filter(Boolean);
  if (parts.length < 2) return "";
  return parts.reverse().join("");
}

export function indexItem(key, v) {
  const n = v.n || "";
  return {
    key,
    c: v.c || key,
    n,
    a: v.a || "",
    m: v.m || "",
    u: v.u || 0,
    _n: S.normIndexed(n),
    _alt: S.norm(`${flipped(n)} ${v.a || ""}`),
    _c: S.norm(v.c || key),
    _m: S.norm(v.m || ""),
  };
}

export function searchCatalog(list, query, groups, { memo = false } = {}) {
  const toks = String(query || "")
    .normalize("NFKC")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  const alts = toks.map((t) => S.expand(t, groups)).filter((a) => a.length);
  if (!alts.length) return [];
  const out = [];
  for (const it of list) {
    let score = 0;
    const ranges = [];
    let ok = true;
    for (const as of alts) {
      let best = 99;
      for (let ai = 0; ai < as.length; ai++) {
        const a = as[ai];
        // 言い換えでヒットしたものは、そのままの言葉でヒットしたものより少し下げる
        const pen = ai === 0 ? 0 : 0.3;
        if (it._c === a) best = Math.min(best, -1 + pen);
        else if (it._c.startsWith(a)) best = Math.min(best, 0 + pen);
        else if (a.length >= 3 && it._c.includes(a)) best = Math.min(best, 1.5 + pen);
        const i = it._n.n.indexOf(a);
        if (i >= 0) {
          best = Math.min(best, (i === 0 ? 0.5 : 1) + pen);
          ranges.push([it._n.map[i], it._n.map[i + a.length - 1] + 1]);
        } else if (it._alt.includes(a)) best = Math.min(best, 1.2 + pen);
        if (memo && it._m.includes(a)) best = Math.min(best, 3 + pen);
      }
      if (best === 99) {
        ok = false;
        break;
      }
      score += best;
    }
    if (ok) out.push({ it, score, ranges: merge(ranges) });
  }
  out.sort((x, y) => x.score - y.score || x.it.n.length - y.it.n.length || x.it.c.localeCompare(y.it.c));
  return out;
}

function merge(r) {
  r.sort((a, b) => a[0] - b[0]);
  const o = [];
  for (const x of r) {
    if (o.length && x[0] <= o[o.length - 1][1]) o[o.length - 1][1] = Math.max(o[o.length - 1][1], x[1]);
    else o.push([...x]);
  }
  return o;
}

// ---- 取り込み用：表の行 → {key: {c, n, a}} ----
export function rowsToCatalog(rows) {
  for (let h = 0; h < Math.min(rows.length, 15); h++) {
    const hd = (rows[h] || []).map((c) =>
      String(c ?? "")
        .normalize("NFKC")
        .replace(/\s/g, ""),
    );
    const iCode = hd.findIndex((c) => /^(部品番号|品番|部品コード|部番|品目コード)$/.test(c));
    const iName = hd.findIndex((c) => /^(部品名|部品名称|品名|名称|商品名)$/.test(c));
    if (iCode < 0 || iName < 0) continue;
    const map = new Map();
    let read = 0;
    for (const r of rows.slice(h + 1)) {
      const c = cleanText(r[iCode]);
      const n = cleanName(r[iName]);
      const key = catKey(c);
      if (!key) continue;
      read++;
      const cur = map.get(key);
      if (!cur) map.set(key, { c, n, a: [] });
      // 同じ品番がハイフンあり・なしで2回あるときは、読みやすいハイフンありを表示に使う
      if (cur && !cur.c.includes("-") && c.includes("-")) cur.c = c;
      if (!cur) continue;
      else if (n && S.norm(n) !== S.norm(cur.n) && !cur.a.some((x) => S.norm(x) === S.norm(n))) {
        // 同じ品番で名前が違うものは別名として残す
        if (!cur.n) cur.n = n;
        else cur.a.push(n);
      }
    }
    const items = {};
    for (const [k, v] of map) items[k] = { c: v.c, n: v.n || "(名称なし)", a: v.a.join(" / ") };
    return { items, read, unique: map.size };
  }
  throw new Error("「部品番号」「部品名」の列が見つかりませんでした");
}
