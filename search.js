// 検索・型式判定

const SMALL = {
  ァ: "ア",
  ィ: "イ",
  ゥ: "ウ",
  ェ: "エ",
  ォ: "オ",
  ャ: "ヤ",
  ュ: "ユ",
  ョ: "ヨ",
  ッ: "ツ",
  ヮ: "ワ",
  ヵ: "カ",
  ヶ: "ケ",
};
const DROP = /[\s・･\-‐‑‒–—―−ーｰ〜~_.,、。]/;

/** 1文字を検索用に正規化（ひらがな→カタカナ、全角→半角、小文字化、小さいカナ→大きいカナ、長音・記号は消す） */
function normChar(c) {
  let out = "";
  for (let d of c.normalize("NFKC").toLowerCase()) {
    const code = d.charCodeAt(0);
    if (code >= 0x3041 && code <= 0x3096) d = String.fromCharCode(code + 0x60);
    d = SMALL[d] || d;
    if (DROP.test(d)) continue;
    out += d;
  }
  return out;
}

/** 文字列を正規化（元の文字位置の対応表つき） */
export function normIndexed(str) {
  str = String(str ?? "");
  let n = "",
    map = [];
  for (let i = 0; i < str.length; i++) {
    let ch = str[i];
    // サロゲートペア
    if (/[\uD800-\uDBFF]/.test(ch) && i + 1 < str.length) {
      ch += str[i + 1];
    }
    const t = normChar(ch);
    for (const x of t) {
      n += x;
      map.push(i);
    }
    if (ch.length === 2) i++;
  }
  return { n, map };
}
export const norm = (s) => normIndexed(String(s ?? "").normalize("NFKC")).n;

/** 言い換えリスト（"フィルター=エレメント" の行）を解析 */
export function parseSynonyms(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((l) =>
      l
        .split(/[=＝,，、]/)
        .map((w) => w.trim())
        .filter(Boolean),
    )
    .filter((g) => g.length >= 2)
    .map((g) => g.map(norm).filter(Boolean));
}
export const DEFAULT_SYNONYMS = [
  "フィルター=エレメント",
  "リア=リヤ",
  "ランプ=ライト",
  "パッキン=ガスケット",
  "ブッシュ=ブッシング",
  "ショック=ストラット=アブソーバー",
  "パワステ=パワーステアリング",
  "ワイパー=ブレード",
  "エアコン=クーラー",
].join("\n");

/** 1語 → 言い換えを含めた候補（正規化済み） */
function expand(tok, groups) {
  const t = norm(tok);
  if (!t) return [];
  const alts = new Set([t]);
  for (const g of groups) {
    for (const w of g) {
      if (w && t.includes(w)) for (const o of g) if (o !== w) alts.add(t.split(w).join(o));
    }
  }
  return [...alts];
}

/** セクションの並び順（数値 → 英字） */
export function secKey(sec) {
  const s = String(sec ?? "");
  const f = parseFloat(s);
  return /^\d/.test(s) && !isNaN(f) ? [0, f, s] : [1, 0, s];
}
export function cmpSec(a, b) {
  const x = secKey(a),
    y = secKey(b);
  return x[0] - y[0] || x[1] - y[1] || x[2].localeCompare(y[2], "ja");
}

/** セクションの大分類（チップ用） */
export const GROUPS = [
  { id: "all", label: "すべて" },
  { id: "hot", label: "★ 頻出" },
  { id: "1", label: "1xx", test: (s) => /^1\d/.test(s) },
  { id: "2", label: "2xx", test: (s) => /^2\d/.test(s) },
  { id: "3", label: "3xx", test: (s) => /^3\d/.test(s) },
  { id: "4", label: "4xx", test: (s) => /^4\d/.test(s) },
  { id: "5", label: "5xx", test: (s) => /^5\d/.test(s) },
  { id: "6", label: "6xx〜9xx", test: (s) => /^[6-9]\d/.test(s) },
  { id: "abc", label: "英字", test: (s) => /^[A-Za-z]/.test(s) },
];

/** 部品ごとに検索用インデックスを作る */
export function indexPart(p) {
  return {
    ...p,
    _name: normIndexed(p.name),
    _note: norm(p.note),
    _code: norm(p.code),
    _sec: norm(p.sec),
  };
}

/**
 * 検索本体
 * @returns [{p, score, ranges:[[s,e],...]}]
 */
export function searchParts(list, query, groups, opt = {}) {
  const toks = String(query || "")
    .normalize("NFKC")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  const out = [];
  const tokAlts = toks.map((t) => expand(t, groups)).filter((a) => a.length);
  for (const p of list) {
    if (opt.sec && String(p.sec) !== String(opt.sec)) continue;
    if (opt.group && opt.group !== "all") {
      if (opt.group === "hot") {
        if (!p.hot) continue;
      } else {
        const g = GROUPS.find((x) => x.id === opt.group);
        if (g && g.test && !g.test(String(p.sec))) continue;
      }
    }
    if (!tokAlts.length) {
      out.push({ p, score: p.hot ? 0 : 1, ranges: [] });
      continue;
    }
    let score = 0,
      ok = true;
    const ranges = [];
    for (const alts of tokAlts) {
      let best = 99;
      for (const a of alts) {
        const i = p._name.n.indexOf(a);
        if (i >= 0) {
          best = Math.min(best, i === 0 ? 0 : 1);
          const s = p._name.map[i];
          let e = p._name.map[i + a.length - 1] + 1;
          while (e < p.name.length && /[ーｰ\-‐]/.test(p.name[e])) e++; // 末尾の長音もハイライト
          ranges.push([s, e]);
        }
        if (p._code.startsWith(a)) best = Math.min(best, 0.5);
        else if (p._code.includes(a)) best = Math.min(best, 1.5);
        if (p._sec === a) best = Math.min(best, 0.8);
        else if (a.length >= 2 && p._sec.startsWith(a)) best = Math.min(best, 2);
        if (p._note.includes(a)) best = Math.min(best, 3);
      }
      if (best === 99) {
        ok = false;
        break;
      }
      score += best;
    }
    if (ok) out.push({ p, score: score - (p.hot ? 0.2 : 0), ranges: mergeRanges(ranges) });
  }
  out.sort(
    (a, b) =>
      a.score - b.score || cmpSec(a.p.sec, b.p.sec) || String(a.p.code).localeCompare(String(b.p.code)),
  );
  return out;
}
function mergeRanges(r) {
  r.sort((a, b) => a[0] - b[0]);
  const o = [];
  for (const x of r) {
    if (o.length && x[0] <= o[o.length - 1][1]) o[o.length - 1][1] = Math.max(o[o.length - 1][1], x[1]);
    else o.push([...x]);
  }
  return o;
}

// 型式
export const normKata = (s) =>
  String(s || "")
    .normalize("NFKC")
    .toUpperCase()
    .replace(/[\s　]/g, "")
    .replace(/[‐‑‒–—―−ー]/g, "-");

/**
 * 型式を解析
 * table: { C27: {name, verified, note}, ... }（キーは大文字）
 */
export function parseKata(input, table) {
  const s = normKata(input);
  if (!s) return null;
  const i = s.lastIndexOf("-");
  const prefix = i >= 0 ? s.slice(0, i + 1) : "";
  const body = (i >= 0 ? s.slice(i + 1) : s).replace(/[^0-9A-Z]/g, "");
  let best = null;
  for (const code of Object.keys(table || {})) {
    if (code && body.endsWith(code) && (!best || code.length > best.length)) best = code;
  }
  const guess = (body.match(/[A-Z]\d{2,3}[A-Z]?$/) || body.match(/[A-Z]{1,2}\d{1,3}[A-Z]?$/) || [""])[0];
  if (best)
    return {
      input: s,
      prefix,
      mid: body.slice(0, body.length - best.length),
      key: best,
      found: table[best],
      guess: best,
    };
  return {
    input: s,
    prefix,
    mid: guess ? body.slice(0, body.length - guess.length) : body,
    key: "",
    found: null,
    guess,
  };
}

/** 車種名から型式を逆引き */
export function reverseKata(q, table) {
  const n = norm(q);
  if (!n) return [];
  return Object.entries(table || {})
    .filter(([, v]) => norm(v.name).includes(n))
    .map(([code, v]) => ({ code, ...v }))
    .sort((a, b) => a.name.localeCompare(b.name, "ja") || a.code.localeCompare(b.code));
}
