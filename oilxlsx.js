// オイル・ケミカルを、関数だけで検索できるExcelにして書き出す（Excel 2016 で動く関数だけを使う）
import { yomiText } from "./yomi.js?v=20261019e";
import { fillDocScreen, fillDocPrint } from "./sphxlsx.js?v=20261019e";
import { fillQuick } from "./sphquick.js?v=20261019e";
const kataOf = (s) => s.replace(/[ぁ-ゖ]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0x60));
const hiraOf = (s) => s.replace(/[ァ-ヶ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x60));
const NF = (s) => String(s ?? "").normalize("NFKC").toLowerCase();

// 読み・別名（固定の補助文字列）：カタカナ版・ひらがな版・同義語（読み・別名）
// 「検索用」の列は、品番・名称などの今の内容＋この文字列から数式で作る（書き換えても検索に反映される）
// groups は「液状=えきじょう」のような同義語を [["液状","えきじょう"],…] にしたもの
export function readingsText(o, groups) {
  const raw = `${o.code} ${o.name} ${o.cap} ${o.use} ${o.cat}`;
  const base = NF(raw);
  const kata = kataOf(base);
  const out = [];
  if (base !== raw.toLowerCase()) out.push(base); // 全角の英数字などは半角にした形も
  // 漢字の読みも入れる（「もれ」で「漏れ」が出るように）
  const y0 = hiraOf(yomiText(base, 0));
  const y1 = hiraOf(yomiText(base, 1));
  out.push(kata, y0, kataOf(y0));
  if (y1 !== y0) out.push(y1, kataOf(y1));
  for (const g of groups) {
    const ws = g.map(NF).filter(Boolean);
    if (ws.some((w) => kata.includes(kataOf(w)))) for (const w of ws) out.push(w, kataOf(w), hiraOf(w));
  }
  return [...new Set(out)].join(" ").slice(0, 6000);
}
// 検索用の列と同じ中身（数式の計算済みの値として入れる）
const searchCell = (o, readings) =>
  [o.code, o.name, o.cap || "", o.use || "", o.cat || "", readings, `${o.code}${o.name}`.replace(/-/g, "").replace(/ /g, "")]
    .join(" ")
    .toLowerCase();

// 種類ごとの言葉づかい（オイル・ケミカル／バルブ）
export const KINDS = {
  oil: {
    title: "オイル・ケミカル",
    item: "商品",
    nameH: "名称",
    capH: "容量",
    catH: "分類",
    capEx: "容量（例：4L）",
    eg: [
      ["atf", "「ATF」が入っている商品が出ます。大文字・小文字はどっちでもOK！"],
      ["オイル 4l", "空白で区切ると、両方入っているものだけ出ます（3つまで）"],
      ["もれ", "ひらがな・カタカナでもOK。「漏れ」のような漢字の言葉も出ます（読みが入っている言葉だけ。出ないときは漢字で入れてください）"],
      ["0w20", "ハイフンなしでも「0W-20」が出ます"],
    ],
  },
  sph: {
    title: "スフィアライト",
    item: "商品",
    nameH: "商品名",
    capH: "明るさ・色温度",
    catH: "種類",
    capEx: "明るさ・色温度（例：6000lm／6000K）",
    eg: [
      ["6000k", "「6000K」の商品が出ます。大文字・小文字はどっちでもOK！"],
      ["h4 6000k", "空白で区切ると、両方入っているものだけ出ます（3つまで）"],
      ["bs09ls300b", "品番（BS09L-S300B）は、ハイフンなしでも探せます"],
      ["ふぉぐ", "ひらがなでもOK。「フォグ」の商品が出ます"],
    ],
  },
  bulb: {
    title: "バルブ",
    item: "バルブ",
    nameH: "規格",
    capH: "入数",
    catH: "形状",
    capEx: "入数（例：10個入）",
    eg: [
      ["t10", "「T10」のバルブが出ます。大文字・小文字はどっちでもOK！"],
      ["h4 12v", "空白で区切ると、両方入っているものだけ出ます（3つまで）"],
      ["ういんかー", "ひらがなでもOK。「ウインカー」が出ます。漢字の言葉も、読みで探せます（載っていないときは漢字で）"],
      ["24v", "24Vのバルブだけが出ます"],
    ],
  },
};

// HYPERLINK は、リンク先が255文字までしか使えない（日本語は1文字が9文字ぶんに増える）。
// 長いときは、名称の先頭だけにして必ず収まるようにする
const gLink = (code, name) => {
  const pre = '"https://www.google.com/search?q="';
  const full = `_xlfn.ENCODEURL(${code}&" "&${name})`;
  const short = `_xlfn.ENCODEURL(${code}&" "&LEFT(${name},8))`;
  const plain = `_xlfn.ENCODEURL(${code})`;
  // 長ければ名称の先頭だけ → それでもだめなら品番だけ
  return `IFERROR(HYPERLINK(${pre}&IF(LEN(${full})>170,${short},${full}),"Google"),HYPERLINK(${pre}&${plain},"Google"))`;
};

const MAXROWS = 200; // 検索シートに出す最大件数
const SPARE = 300; // データシートの空き行（あとから足せる）

// 文字の幅（全角は2、半角は1）。列の幅と行の高さを、中身に合わせて決めるのに使う
const dispW = (s) => {
  let w = 0;
  for (const ch of String(s ?? "")) {
    const c = ch.codePointAt(0);
    w += c > 0x7f && !(c >= 0xff61 && c <= 0xff9f) ? 2 : 1;
  }
  return w;
};
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
// どの商品も途切れず見えるように、列の幅と1行の高さを決める
// （検索の結果は数式なので、行の高さが自動では変わらない。どの商品が出ても足りる高さに揃える）
export const USE_LINES = 6;
export const USE_SUFFIX = "…（ここをクリックで全文）";
export const cutUse = (t, limit) => {
  const x = String(t ?? "");
  return x.length > limit ? x.slice(0, limit) + USE_SUFFIX : x;
};
export function layoutFor(items, heads = {}) {
  const val = (o, f) => o[f];
  const maxW = (f) => Math.max(0, ...items.map((o) => dispW(val(o, f))));
  const w = {
    code: clamp(Math.ceil(maxW("code") * 1.15) + 2, 14, 28),
    name: clamp(Math.ceil(maxW("name") * 1.1) + 2, 24, 56),
    cap: clamp(Math.ceil(maxW("cap") * 1.15) + 2, 8, 16),
    use: clamp(Math.ceil(maxW("use") * 1.1) + 2, 24, 110),
    cat: clamp(Math.ceil(maxW("cat") * 1.15) + 2, 12, 26),
  };
  // 見出しの文字（例：「明るさ・色温度」）も切れないように、列の幅の下限にする
  for (const f of Object.keys(w)) if (heads[f]) w[f] = Math.max(w[f], Math.ceil(dispW(heads[f]) * 1.1) + 3);
  let lines = 1;
  for (const o of items)
    for (const f of Object.keys(w)) lines = Math.max(lines, Math.ceil((dispW(val(o, f)) * 1.1) / (w[f] - 1)));
  // 行の高さは、いちばん長い用途に合わせると高くなりすぎて、画面にほとんど入らず、スクロールで飛ぶ。
  // そこで高さは最大 USE_LINES 行ぶんにして、それより長い用途は「…」で省略する（全文は印刷用・データのシートで見られる）
  const useLines = Math.ceil((maxW("use") * 1.1) / (w.use - 1));
  lines = Math.min(lines, USE_LINES);
  const useLimit =
    useLines > USE_LINES
      ? // HYPERLINK の表示する文字は、255文字までしか使えない（超えると #VALUE! になる）ので、余裕をみて250文字以内にする
        Math.min(Math.floor(((USE_LINES * (w.use - 1)) / 1.1 / 1.9) * 0.92), 250) - USE_SUFFIX.length
      : Infinity;
  return { w, lines, height: Math.max(24, lines * 19 + 4), useLimit };
}


// ---- 使い方シート ----
const GW = { b: 7, c: 28, d: 62 }; // 列の幅（B・C・D）
const guideBlocks = (protect, K, nameP) => [
  ["title", `${K.title}検索の使い方`],
  ["sub", "お客さんに聞かれたとき、ぱっと探してぱっと答えるためのシートです。むずかしくないので、さくっと読んでみてください！"],
  ["h", "まずはこれだけ（3ステップ）"],
  ["step", "①", "「検索」シートの黄色いところに、探したい言葉を入れて Enter！（品番でも名前でも用途でもOK）"],
  ["step", "②", `下に結果がズラッと出ます。多すぎるときは「${K.catH}」のプルダウンで絞ってください`],
  ["step", "③", "もっと調べたいときは、右端の「Google」を押せば検索ページが開きます"],
  ["h", "検索のコツ"],
  ["th", "入れる言葉", "こうなる"],
  ...K.eg.map(([a, b]) => ["tr", a, b]),
  ["tr", "（何も入れない）", "全部出ます（200件まで）"],
  ["b", "英数字は半角で入れてください。全角だと当たらないことがあります"],
  ["b", "「?」と「*」は検索で特別な意味になる記号なので、入れないでください"],
  ["h", `${K.item}を追加・修正・削除する（すべて「${K.title}データ」シートで行います）`],
  ["step", "追加", `「${K.title}データ」シートの、データが入っている最後の行のすぐ下（空き行）に入れてください。空き行は300行ぶん用意してあります`],
  ["b", `A列：品番（必ず入れる）　B列：${K.nameH}（必ず入れる）　C列：Google（自動。触らない）`],
  ["b", `D列：${K.capEx}　E列：用途　F列：${K.catH}　← この3つは空でもOK`],
  ["step", "修正", `「${K.title}データ」シートの直したい行を、そのまま書き換えればOK。検索にもすぐ反映されます`],
  ["step", "削除", `「${K.title}データ」シートで、その行のA・B・D・E・F列の中身だけを消してください（選んで Delete）`],
  ["step", K.catH, `F列に文字を入れるだけで、新しい${K.catH}も作れます。検索シートのプルダウンにも出したいときは、L列の「${K.catH}一覧」の一番下に足してください`],
  ["b", `並べ替えや絞り込みは、「${K.title}データ」シートの見出しのボタンから自由にどうぞ`],
  ["ng", "行ごと削除したり、データの途中や空き行の間に行を「挿入」したりしないでください（検索に出なくなります）"],
  ["ng", "G〜K列は計算用の隠し列です。表示したり、中身を消したり、列を削除したりしないでください"],
  ["h", "紙に印刷して使う（パソコンが苦手な方もOK）"],
  ["b", "パソコンが苦手な方は、紙に印刷して使うこともできます！ 印刷しておけば、パソコンを開かなくても、手元でぱっと調べられます"],
  ["step", "全部", `「${nameP}」シートを開いて、そのまま印刷（Ctrl+P）。文字は大きめ（14ポイント）、A4たてで、1つずつ見やすく並びます`],
  ["step", K.catH + "ごと", `「${nameP}」の1ページ目に、${K.catH}ごとのページ番号の表があります。印刷の「設定」のページ指定に、そのページ番号を入れてください`],
  ["step", "変更", `今ある${K.item}の内容は、「${K.title}データ」シートを書き換えると、「${nameP}」にもそのまま変わります（印刷用シートは直接さわらないでください）`],
  ["b", `新しく足した${K.item}は、「${nameP}」には出ません（検索シートで見てください）。文字を長く書き換えて切れたときは、その行の高さを広げてください`],
  ["h", "困ったときは"],
  ["b", "何も出ない・計算されない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", `あるはずなのに出ない → 全角になっていないか、${K.catH}が「すべて」になっているか見てください`],
  ["b", "追加したのに出ない → 品番（A列）が空になっていないか、行を挿入していないか見てください"],
  ["b", "200件で止まる → 言葉を足して絞り込んでください"],
  ["b", "おかしくなった → 元のファイルをもう一回コピーすれば、元どおりです"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "検索欄に入れた言葉は、保存しなくて大丈夫です。閉じるときは「保存しない」でOK！"],
  ["b", "検索結果の「用途」は、長いと途中で「…」と省略されます（見やすいように、行の高さをそろえているためです）。その「用途」をクリックすると、「データ」シートのその行に飛んで、全文が読めます（戻るときは、下の「検索」シートの名前を押してください）。「印刷用」シートでも全文が読めます"],
  ["b", "検索結果の表（数式が入ってます）は、消したり上書きしたりしないでください"],
  ...(protect
    ? [
        ["b", "「検索」と「使い方」のシートは、間違って消さないように保護してあります。入力できるのは黄色い欄とプルダウンだけです"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

export function fillGuide(ws, { font, protect, K, nameP, blocks }) {
  const AC = "FFF2C230";
  const DK = "FF17191C";
  ws.columns = [{ width: 2 }, { width: GW.b }, { width: GW.c }, { width: GW.d }];
  const spanW = GW.c + GW.d; // 文章がまたぐ幅（C〜D）
  const lines = (text, width) => Math.max(1, Math.ceil((dispW(text) * 1.12) / (width - 2)));
  const hairs = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
  let r = 1;
  for (const blk of blocks || guideBlocks(protect, K, nameP)) {
    const [kind, a, b] = blk;
    const row = ws.getRow(r);
    const span = (text, o = {}) => {
      ws.mergeCells(r, 3, r, 4);
      const c = row.getCell(3);
      c.value = text;
      c.font = o.font || font;
      c.alignment = { vertical: "middle", wrapText: true };
    };
    if (kind === "title") {
      ws.mergeCells(r, 2, r, 4);
      const c = row.getCell(2);
      c.value = a;
      c.font = { ...font, size: 20, bold: true, color: { argb: DK } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AC } };
      c.alignment = { vertical: "middle", indent: 1 };
      for (let i = 2; i <= 4; i++) row.getCell(i).fill = c.fill;
      row.height = 40;
    } else if (kind === "sub") {
      ws.mergeCells(r, 2, r, 4);
      const c = row.getCell(2);
      c.value = a;
      c.font = { ...font, color: { argb: "FF4A4F57" } };
      c.alignment = { vertical: "middle", wrapText: true, indent: 1 };
      row.height = lines(a, GW.b + spanW) * 19 + 10;
    } else if (kind === "h") {
      r++; // 見出しの前に1行あける
      ws.getRow(r - 1).height = 10;
      const hr = ws.getRow(r);
      ws.mergeCells(r, 2, r, 4);
      const c = hr.getCell(2);
      c.value = a;
      c.font = { ...font, size: 12, bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DK } };
      c.alignment = { vertical: "middle", indent: 1 };
      for (let i = 2; i <= 4; i++) hr.getCell(i).fill = c.fill;
      hr.height = 24;
    } else if (kind === "step") {
      const n = row.getCell(2);
      n.value = a;
      n.font = { ...font, bold: true, size: a.length > 1 ? 11 : 14, color: { argb: "FFB06A00" } };
      n.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      span(b);
      row.height = Math.max(24, lines(b, spanW) * 19 + 6);
      for (let i = 2; i <= 4; i++) row.getCell(i).border = hairs;
    } else if (kind === "b" || kind === "ng") {
      const m = row.getCell(2);
      m.value = kind === "ng" ? "×" : "・";
      m.font = { ...font, bold: true, color: { argb: kind === "ng" ? "FFD93025" : "FF7A7F87" } };
      m.alignment = { vertical: "middle", horizontal: "center" };
      span(a, kind === "ng" ? { font: { ...font, bold: true, color: { argb: "FFD93025" } } } : {});
      row.height = Math.max(22, lines(a, spanW) * 19 + 4);
    } else if (kind === "th" || kind === "tr") {
      const c1 = row.getCell(3);
      const c2 = row.getCell(4);
      c1.value = a;
      c2.value = b;
      const head = kind === "th";
      c1.font = head ? { ...font, bold: true } : { name: "Consolas", size: 11, bold: true };
      c2.font = head ? { ...font, bold: true } : font;
      for (const c of [c1, c2]) {
        c.alignment = { vertical: "middle", wrapText: true };
        c.border = { top: { style: "thin", color: { argb: "FF8A8F98" } }, bottom: { style: "thin", color: { argb: "FF8A8F98" } }, left: { style: "thin", color: { argb: "FF8A8F98" } }, right: { style: "thin", color: { argb: "FF8A8F98" } } };
        if (head) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3C4" } };
      }
      row.height = Math.max(24, Math.max(lines(a, GW.c), lines(b, GW.d)) * 19 + 6);
    }
    r++;
  }
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true };
  if (protect) ws.sheetProtection = { sheet: true };
}


// ---- 印刷用シート：紙で見やすい大きな文字（14pt）・A4たて・1商品ずつ区切って並べる ----
// 分類ごとに新しいページから始める。先頭ページ（目次）に、分類ごとのページ番号を載せる。
// 関数ではなく書き出した時点の内容をそのまま入れる（ページの区切りを正確に決めるため）
const PR = { a: 24, b: 36, c: 20 }; // 列の幅（品番・名称・容量）。A4たてに余裕を持って収まる幅
const PT = (u) => u * 5.25; // 列の幅（文字数）→ 点（pt）のおおよその換算
const PAGE_CAP = 700; // 1ページに積める高さ（pt）。A4たての印刷できる高さより少し小さめ
export function planPrint(items) {
  const cw = { code: PT(PR.a), name: PT(PR.b), cap: PT(PR.c) };
  const total = PT(PR.a + PR.b + PR.c);
  const perLine = (width) => Math.max(4, Math.floor((width - 10) / 7.1));
  const linesOf = (text, width) => Math.max(1, Math.ceil((dispW(text) * 1.08) / perLine(width)));
  const useLines = (t) => linesOf(t, total);
  // 用途が長いとき、1行の高さの上限（409pt）を超えないよう、いくつかの行に分ける
  const chunks = (text) => {
    const out = [];
    const limit = (17 * perLine(total)) / 1.08;
    let cur = "";
    let w = 0;
    for (const ch of String(text ?? "")) {
      const cwid = dispW(ch);
      if (w + cwid > limit && /[、。，．\s]/.test(cur.slice(-1))) (out.push(cur), (cur = ""), (w = 0));
      else if (w + cwid > limit * 1.08) (out.push(cur), (cur = ""), (w = 0));
      cur += ch;
      w += cwid;
    }
    if (cur) out.push(cur);
    return out;
  };
  const block = (o, idx) => {
    const lines = Math.max(
      1,
      Math.ceil((dispW(o.code) * 1.2 * 7.7) / (cw.code - 8)),
      linesOf(o.name, cw.name),
      linesOf(o.cap || "", cw.cap),
    );
    const head = lines * 21 + 12;
    const uses = chunks(o.use).map((t) => ({ t, h: useLines(t) * 21 + 12 }));
    return { o, idx, head, uses };
  };
  const bycat = new Map();
  for (const o of items) {
    const k = o.cat || "未分類";
    if (!bycat.has(k)) bycat.set(k, []);
    bycat.get(k).push(block(o, items.indexOf(o)));
  }
  const TITLE_H = 40;
  // ページに詰める（商品の見出し＋用途の最初の行は、必ず同じページに）
  const sections = [];
  for (const [cat, blocks] of bycat) {
    const pages = [[]];
    let used = TITLE_H;
    for (const b of blocks) {
      const need = b.head + (b.uses[0]?.h || 0) + 8;
      if (used + need > PAGE_CAP && pages[pages.length - 1].length) (pages.push([]), (used = TITLE_H));
      pages[pages.length - 1].push(b);
      used += b.head;
      b.uses.forEach((u, i) => {
        if (i > 0 && used + u.h > PAGE_CAP) (pages.push([]), (used = TITLE_H));
        used += u.h;
      });
      used += 8;
    }
    sections.push({ cat, n: blocks.length, pages });
  }
  return sections;
}

function fillPrint(ws, items, { font, protect, K, date, QD }) {
  const DK = "FF17191C";
  ws.columns = [{ width: PR.a }, { width: PR.b }, { width: PR.c }];
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = {
    paperSize: 9,
    orientation: "portrait",
    scale: 100,
    horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.55, bottom: 0.7, header: 0.25, footer: 0.3 },
  };
  ws.headerFooter = { oddFooter: `&L&11${K.title}　一覧（${date}）&R&12ページ &P / &N` };
  const total = PT(PR.a + PR.b + PR.c);
  const sections = planPrint(items);
  const grey = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEDEDED" } };
  const line = (style) => ({ style, color: { argb: "FF303030" } });
  // 目次のページ数（1行28pt）
  const TOC_FIRST = 13; // 1ページ目に入る分類の数（見出しなどの下）
  const TOC_NEXT = 21; // 2ページ目以降
  const tocPages = 1 + Math.ceil(Math.max(0, sections.length - TOC_FIRST) / TOC_NEXT);
  let page = tocPages + 1;
  const start = new Map();
  for (const sec of sections) {
    start.set(sec.cat, [page, page + sec.pages.length - 1]);
    page += sec.pages.length;
  }
  let r = 1;
  const merged = (row, v, o = {}) => {
    ws.mergeCells(row, 1, row, 3);
    const c = ws.getRow(row).getCell(1);
    c.value = v;
    c.font = o.font || { name: font.name, size: 14 };
    c.alignment = { vertical: "middle", wrapText: true, horizontal: o.h || "left", indent: o.indent || 0 };
    if (o.fill) for (let i = 1; i <= 3; i++) ws.getRow(row).getCell(i).fill = o.fill;
    ws.getRow(row).height = o.height || 24;
    return c;
  };
  // --- 目次 ---
  merged(r++, `${K.title}　一覧`, { font: { name: font.name, size: 26, bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: DK } }, height: 52, indent: 1 });
  merged(r++, `作成日　${date}　／　全 ${items.length} 件`, { font: { name: font.name, size: 15 }, height: 30, indent: 1 });
  r++;
  ws.getRow(r - 1).height = 10;
  merged(r++, "分類ごとに印刷するとき", { font: { name: font.name, size: 16, bold: true }, height: 30 });
  const howTo = [
    "① 下の表で、印刷したい分類の「ページ」を見る",
    "② ［ファイル］→［印刷］の「設定」で、ページ指定の「から」「まで」に、そのページ番号を入れる",
    "③ 全部印刷するときは、ページ指定をしないで、そのまま印刷する",
  ];
  for (const t of howTo) {
    // 長い行は折り返して2行以上になるので、文字の量に合わせて高さを決める
    const n = Math.max(1, Math.ceil((dispW(t) * 1.1) / Math.floor((total - 24) / 7.1)));
    merged(r++, t, { font: { name: font.name, size: 14 }, height: Math.max(30, n * 21 + 10), indent: 1 });
  }
  r++;
  ws.getRow(r - 1).height = 12;
  const hdr = ws.getRow(r);
  hdr.height = 30;
  // 分類の名前が長くても切れないように、「分類（件数）」は A〜B をまとめて使う
  ws.mergeCells(r, 1, r, 2);
  [[`${K.catH}（件数）`, 1], ["ページ", 3]].forEach(([t, i]) => {
    const c = hdr.getCell(i);
    c.value = t;
    c.font = { name: font.name, size: 15, bold: true };
    c.fill = grey;
    c.border = { top: line("medium"), bottom: line("medium") };
    c.alignment = { vertical: "middle", horizontal: i === 1 ? "left" : "center", indent: i === 1 ? 1 : 0 };
  });
  hdr.getCell(2).fill = grey;
  hdr.getCell(2).alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  hdr.getCell(2).border = { top: line("medium"), bottom: line("medium") };
  r++;
  sections.forEach((sec, i) => {
    if (i === TOC_FIRST || (i > TOC_FIRST && (i - TOC_FIRST) % TOC_NEXT === 0)) ws.getRow(r - 1).addPageBreak();
    const row = ws.getRow(r);
    const [a, b] = start.get(sec.cat);
    // 分類の名前と件数は A〜B（幅60）にまとめて入れる。長い名前は折り返して、行の高さを増やす
    const label = `${sec.cat}　${sec.n}件`;
    const lines = Math.max(1, Math.ceil((dispW(label) * 16 * 0.0952 * 1.1) / (PR.a + PR.b - 3)));
    row.height = Math.max(30, lines * 24 + 8);
    ws.mergeCells(r, 1, r, 2);
    row.getCell(1).value = {
      richText: [
        { text: sec.cat, font: { name: font.name, size: 16, bold: true } },
        { text: `　${sec.n}件`, font: { name: font.name, size: 14, color: { argb: "FF4A4F57" } } },
      ],
    };
    row.getCell(3).value = a === b ? `${a}` : `${a} 〜 ${b}`;
    row.getCell(3).font = { name: font.name, size: 16, bold: true };
    for (let k = 1; k <= 3; k++) {
      row.getCell(k).border = { bottom: { style: "thin", color: { argb: "FF808080" } } };
      row.getCell(k).alignment = { vertical: "middle", horizontal: k <= 2 ? "left" : "center", indent: k === 1 ? 1 : 0, wrapText: k === 1 };
    }
    r++;
  });
  // 目次が複数ページになるとき（分類がとても多いとき）は、ここで区切る
  const brk = () => {
    ws.getRow(r - 1).addPageBreak();
  };
  // --- 各分類 ---
  for (const sec of sections) {
    sec.pages.forEach((blocks, pi) => {
      brk();
      const of = sec.pages.length;
      const barText = `${sec.cat}　${of > 1 ? `（${pi + 1} / ${of}）` : ""}`;
      // 名前が長いときは、1行に収まる大きさまで文字を小さくする（行の高さを変えると、ページの区切りがずれるため）
      const barSize = [20, 18, 16, 14, 12].find((sz) => dispW(barText) * sz * 0.0952 * 1.06 <= total / 5.25 - 4 - 3) || 12;
      merged(r++, barText, {
        font: { name: font.name, size: barSize, bold: true, color: { argb: "FFFFFFFF" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: DK } },
        height: 40,
        indent: 1,
      });
      for (const b of blocks) {
        const o = b.o;
        const hr = ws.getRow(r++);
        hr.height = b.head;
        const vals = [o.code, o.name, o.cap || ""];
        const SRC = ["A", "B", "D"]; // データシートの列（品番・名称・容量）
        vals.forEach((v, i) => {
          const c = hr.getCell(i + 1);
          // 「データ」シートの同じ商品の行を表示する（データを直すと、ここも変わる）
          c.value = { formula: `IF(${QD}!$A${b.idx + 2}="","",${QD}!$${SRC[i]}${b.idx + 2}&"")`, result: String(v) };
          c.font = i === 0 ? { name: "Consolas", size: 15, bold: true } : { name: font.name, size: 14, bold: i === 1 };
          c.fill = grey;
          c.border = { top: line("medium") };
          c.alignment = { vertical: "middle", wrapText: true, horizontal: i === 2 ? "right" : "left", indent: i === 0 ? 1 : 0 };
        });
        for (const u of b.uses) {
          ws.mergeCells(r, 1, r, 3);
          const c = ws.getRow(r).getCell(1);
          c.value =
            b.uses.length === 1
              ? { formula: `IF(${QD}!$A${b.idx + 2}="","",${QD}!$E${b.idx + 2}&"")`, result: u.t }
              : u.t; // 長すぎて何行かに分けた用途は、そのままの文字
          c.font = { name: font.name, size: 14 };
          c.alignment = { vertical: "top", wrapText: true, indent: 1 };
          ws.getRow(r).height = u.h;
          r++;
        }
        ws.getRow(r++).height = 8; // 商品と商品のあき
      }
    });
  }
  if (protect) ws.sheetProtection = { sheet: true };
}

// 既存のブックに同じ名前のシートがあれば「検索2」のように避ける
export function freeName(wb, base) {
  let n = base,
    i = 2;
  while (wb.getWorksheet(n)) n = `${base}${i++}`;
  return n;
}
const quote = (n) => `'${n.replace(/'/g, "''")}'`;

// opt.wb を渡すと、そのブック（会社のExcel）に検索・データのシートを足す。作成者などは元のまま
export function buildOilWorkbook(ExcelJS, items, groups, opt = {}) {
  const K = KINDS[opt.kind || "oil"];
  const wb = opt.wb || new ExcelJS.Workbook();
  if (!opt.wb) wb.creator = "セクション検索";
  // 単独のファイルでは開いたときに再計算させる。会社のExcelに差し込むときは、計算の設定には触れず、
  // 最初の状態（検索語なし）の答えを計算済みの値として入れておく
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nameS = opt.names?.search || freeName(wb, `${K.title}検索`);
  const QS = quote(nameS);
  const nameG = opt.names?.guide || freeName(wb, "使い方");
  const nameP = opt.names?.print || freeName(wb, `${K.title}印刷用`);
  const nameD = opt.names?.data || freeName(wb, `${K.title}データ`);
  const QD = quote(nameD);
  const nameQuick = opt.quick ? opt.names?.quick || freeName(wb, "品番早見表") : null;
  const di = opt.docInfo || {};
  const nameDoc = opt.docs ? opt.names?.doc || freeName(wb, di.screen || "補足資料") : null;
  const nameDocP = opt.docs ? opt.names?.docPrint || freeName(wb, di.print || "補足資料印刷用") : null;
  const N = items.length;
  const LAST = N + 1 + SPARE; // データシートの最終行
  const font = { name: "Yu Gothic UI", size: 11 };
  const L = layoutFor(items, { code: "品番", name: K.nameH, cap: K.capH, use: "用途", cat: K.catH });
  const AC = "FFF2C230";
  const DK = "FF17191C";
  const thin = { style: "thin", color: { argb: "FF8A8F98" } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };

  // 既存のシートが選ばれたままだと、まとめて編集されてしまうので外す
  for (const w of wb.worksheets) w.views = (w.views || []).map((v) => ({ ...v, tabSelected: false }));
  const ws = wb.addWorksheet(nameS, {
    views: [{ state: "frozen", ySplit: 6, showGridLines: false, tabSelected: true }],
  });
  const wg = wb.addWorksheet(nameG, { views: [{ showGridLines: false }] });
  const wp = wb.addWorksheet(nameP);
  const wquick = opt.quick ? wb.addWorksheet(nameQuick) : null;
  const wdoc = opt.docs ? wb.addWorksheet(nameDoc, { views: [{ showGridLines: false }] }) : null;
  const wdocp = opt.docs ? wb.addWorksheet(nameDocP) : null;
  const wd = wb.addWorksheet(nameD, { views: [{ state: "frozen", ySplit: 1 }] });

  // ---- データ ----
  wd.columns = [
    { header: "品番", width: L.w.code },
    { header: K.nameH, width: L.w.name },
    { header: "Google", width: 10 },
    { header: K.capH, width: L.w.cap },
    { header: "用途", width: L.w.use },
    { header: K.catH, width: L.w.cat },
    { header: "読み・別名（自動）", width: 16 },
    { header: "検索用（自動）", width: 18 },
    { header: "一致", width: 6 },
    { header: "順位", width: 6 },
    { header: "", width: 3 },
    { header: `${K.catH}一覧`, width: 14 },
  ];
  const hr = wd.getRow(1);
  hr.font = { ...font, bold: true, color: { argb: DK } };
  hr.eachCell((c, i) => {
    if (i <= 6) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AC } };
  });
  const cats = [...new Set(items.map((o) => o.cat).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ja"));
  // 分類一覧の列は、いちばん長い名前が入る幅にする（切れて見えないように）
  wd.getColumn(12).width = clamp(Math.max(14, ...cats.map((c) => Math.ceil(dispW(c) * 1.05) + 2)), 14, 46);
  wd.getCell("L2").value = "すべて";
  cats.forEach((c, i) => (wd.getCell(`L${i + 3}`).value = c));
  for (let r = 2; r <= LAST; r++) {
    const o = items[r - 2];
    const row = wd.getRow(r);
    if (o) {
      row.getCell(1).value = o.code;
      row.getCell(2).value = o.name;
      row.getCell(4).value = o.cap || "";
      row.getCell(5).value = o.use || "";
      row.getCell(6).value = o.cat || "";
      row.getCell(7).value = readingsText(o, groups);
    }
    // Google：品番＋名称から数式で作る（行を足しても、品番を直しても、そのまま使える）
    row.getCell(3).value = {
      formula: `IF($A${r}="","",${gLink(`$A${r}`, `$B${r}`)})`,
      result: o ? "Google" : "",
    };
    row.getCell(3).font = { ...font, color: { argb: "FF1A73E8" }, underline: true };
    // 検索用：今の品番・名称などから数式で作る（書き換え・追加しても検索に反映される）
    row.getCell(8).value = {
      formula: `IF($A${r}="","",LOWER($A${r}&" "&$B${r}&" "&$D${r}&" "&$E${r}&" "&$F${r}&" "&$G${r}&" "&SUBSTITUTE(SUBSTITUTE($A${r}&$B${r},"-","")," ","")))`,
      result: o ? searchCell(o, row.getCell(7).value) : "",
    };
    row.getCell(9).value = {
      result: o ? 1 : 0,
      formula: `IF($A${r}="",0,IF(AND(OR(${QS}!$K$2="",ISNUMBER(SEARCH(${QS}!$K$2,$H${r}))),OR(${QS}!$K$3="",ISNUMBER(SEARCH(${QS}!$K$3,$H${r}))),OR(${QS}!$K$4="",ISNUMBER(SEARCH(${QS}!$K$4,$H${r}))),OR(${QS}!$B$3="すべて",${QS}!$B$3="",$F${r}=${QS}!$B$3)),1,0))`,
    };
    row.getCell(10).value = { formula: `IF($I${r}=1,COUNTIF($I$2:$I${r},1),"")`, result: o ? r - 1 : "" };
  }
  wd.autoFilter = { from: "A1", to: "F1" };
  for (const c of [7, 8, 9, 10, 11]) wd.getColumn(c).hidden = true; // 計算用の列は見せない（触らせない）
  for (let r = 2; r <= N + 1; r++) for (const c of [1, 2, 3, 4, 5, 6]) wd.getRow(r).getCell(c).alignment = { vertical: "middle", wrapText: true };
  for (const c of ["G", "H", "I", "J"]) wd.getColumn(c).font = { ...font, color: { argb: "FF9AA0A6" } };

  // ---- 検索 ----
  ws.columns = [
    { width: 6 },
    { width: L.w.code },
    { width: L.w.name },
    { width: L.w.cap },
    { width: L.w.use },
    { width: L.w.cat },
    { width: 10 },
    { width: 3 },
    { width: 6 },
    { width: 3 },
    { width: 20 },
  ];
  ws.getCell("A1").value = `${K.title}検索`;
  ws.getCell("A1").font = { ...font, size: 18, bold: true };
  ws.getCell("A2").value = "検索";
  ws.getCell("A3").value = K.catH;
  for (const a of ["A2", "A3"]) ws.getCell(a).font = { ...font, bold: true };
  ws.mergeCells("B2:E2");
  ws.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3C4" } };
  ws.getCell("B2").font = { ...font, size: 14, bold: true };
  for (const a of ["B2", "C2", "D2", "E2"]) ws.getCell(a).border = box;
  ws.getRow(2).height = 26;
  ws.getCell("B3").value = "すべて";
  ws.getCell("B3").border = box;
  ws.getCell("B3").font = font;
  ws.getCell("B3").dataValidation = {
    type: "list",
    allowBlank: false,
    formulae: [`${QD}!$L$2:$L$${Math.max(cats.length + 2, 30)}`],
  };
  ws.getCell("C3").value = `← ▼から${K.catH}を選ぶと、その${K.catH}だけに絞れます`;
  ws.getCell("C3").font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
  ws.getCell("E3").value = "↑ 黄色い欄に入れて Enter（空白で区切ると絞り込み・ひらがなもOK）";
  ws.getCell("E3").font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
  ws.getCell("E3").alignment = { horizontal: "right", vertical: "middle" };
  ws.getRow(5).height = 8;
  ws.getCell("A4").value = {
    result: `${N} 件ヒット（全 ${N} 件）${N > MAXROWS ? `　※先頭の${MAXROWS}件まで表示。語を足して絞ってください` : ""}`,
    formula: `SUM(${QD}!$I$2:$I$${LAST})&" 件ヒット（全 "&COUNTA(${QD}!$A$2:$A$${LAST})&" 件）"&IF(SUM(${QD}!$I$2:$I$${LAST})>${MAXROWS},"　※先頭の${MAXROWS}件まで表示。語を足して絞ってください","")`,
  };
  ws.getCell("A4").font = { ...font, bold: true, color: { argb: "FFB06A00" } };
  // 検索語を3つまでに分ける（計算用・K列は非表示）
  ws.getCell("K1").value = "計算用";
  ws.getCell("K2").value = { formula: `IFERROR(LEFT($K$5,FIND(" ",$K$5)-1),$K$5)`, result: "" };
  ws.getCell("K3").value = { formula: `IFERROR(LEFT($L$5,FIND(" ",$L$5)-1),$L$5)`, result: "" };
  ws.getCell("K4").value = { formula: `IFERROR(MID($L$5,FIND(" ",$L$5)+1,200),"")`, result: "" };
  ws.getCell("K5").value = { formula: `LOWER(TRIM(SUBSTITUTE($B$2,"　"," ")))`, result: "" };
  ws.getCell("L5").value = { formula: `IFERROR(MID($K$5,FIND(" ",$K$5)+1,200),"")`, result: "" };
  const FIRST = 7; // 一覧の最初の行（見出しは6行目）
  const H = ["No", "品番", K.nameH, K.capH, "用途", K.catH, "Google"];
  H.forEach((h, i) => {
    const c = ws.getRow(FIRST - 1).getCell(i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AC } };
    c.border = box;
  });
  for (let k = 1; k <= MAXROWS; k++) {
    const r = FIRST - 1 + k;
    const row = ws.getRow(r);
    const o = items[k - 1];
    row.getCell(9).value = { formula: `IFERROR(MATCH(${k},${QD}!$J$1:$J$${LAST},0),"")`, result: o ? k + 1 : "" };
    const at = (col, v) => ({
      formula: `IF($I${r}="","",INDEX(${QD}!$${col}$1:$${col}$${LAST},$I${r})&"")`,
      result: o ? String(v ?? "") : "",
    });
    row.getCell(1).value = { formula: `IF($I${r}="","",${k})`, result: o ? k : "" };
    row.getCell(2).value = at("A", o?.code);
    row.getCell(3).value = at("B", o?.name);
    row.getCell(4).value = at("D", o?.cap);
    if (Number.isFinite(L.useLimit)) {
      // 長い用途は途中までにして「…」を付ける（行の高さをそろえたまま、どの商品が出ても収まるように）
      const x = `INDEX(${QD}!$E$1:$E$${LAST},$I${r})`;
      row.getCell(5).value = {
        // 省略したものは、押すと「データ」シートのその行に飛んで、全文が読める
        // 万一リンクが作れなくても、エラーにせず、省略した文字だけは出す
        formula: `IF($I${r}="","",IF(LEN(${x})>${L.useLimit},IFERROR(HYPERLINK("#'${nameD.replace(/'/g, "''")}'!E"&$I${r},LEFT(${x},${L.useLimit})&"${USE_SUFFIX}"),LEFT(${x},${L.useLimit})&"${USE_SUFFIX}"),${x}&""))`,
        result: o ? cutUse(o.use, L.useLimit) : "",
      };
    } else row.getCell(5).value = at("E", o?.use);
    row.getCell(6).value = at("F", o?.cat);
    row.getCell(7).value = {
      formula: `IF($I${r}="","",${gLink(`$B${r}`, `$C${r}`)})`,
      result: o ? "Google" : "",
    };
    for (let i = 1; i <= 7; i++) {
      const c = row.getCell(i);
      c.font = i === 7 ? { ...font, color: { argb: "FF1A73E8" }, underline: true } : font;
      c.border = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
      c.alignment = { vertical: "middle", wrapText: i >= 2 && i <= 6 };
    }
    if (k <= N + 30) row.height = L.height; // 商品の数＋あとから足す分だけ。それより下の空の行は普通の高さ
    row.getCell(2).font = { name: "Consolas", size: 11, bold: true };
  }
  ws.getColumn(9).hidden = true;
  ws.getColumn(11).hidden = true;
  ws.getColumn(12).hidden = true;
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  const gBlocks = guideBlocks(!!opt.protect, K, nameP);
  if (opt.quick) {
    const at = gBlocks.findIndex((b) => b[0] === "h" && b[1] === "困ったときは");
    gBlocks.splice(
      at,
      0,
      ["h", "品番早見表（紙に印刷して使う）"],
      ["step", "早見表", `「${nameQuick}」シートは、シリーズ・形状・色から品番をさがす表です。A4たて1枚に収まるようにしてあるので、そのまま印刷（Ctrl+P）できます`],
      ["b", "書き出した時点の商品が載っています。あとから足した商品は、もう一度書き出すと載ります"],
    );
  }
  if (opt.docs) {
    // 「困ったときは」の前に、補足資料の説明を入れる
    const at = gBlocks.findIndex((b) => b[0] === "h" && b[1] === "困ったときは");
    gBlocks.splice(
      at,
      0,
      ["h", di.heading || "補足資料（お客さんに説明するとき）"],
      ["step", "画面", di.screenNote ? di.screenNote(nameDoc) : `「${nameDoc}」シートを開くと、「30秒でこう説明する」「車検の話」「聞かれたときの答え方」などが、読みやすくまとまっています。一番上の「もくじ」を押すと、その場所に飛びます`],
      ["step", "印刷", `「${nameDocP}」シートは、紙に印刷する用です。文字を大きく（14ポイント）して、A4たてで、章ごとに（バルブは1つの形状ごとに）新しいページから始まります。そのまま印刷（Ctrl+P）してください`],
      ["b", di.note || "数字や言葉は、カタログに載っているものだけです。数値は規格値で、保証値ではありません"],
    );
  }
  fillGuide(wg, { font, protect: !!opt.protect, K, nameP, blocks: gBlocks });
  const dt = new Date();
  fillPrint(wp, items, { font, protect: !!opt.protect, K, date: `${dt.getFullYear()}年${dt.getMonth() + 1}月${dt.getDate()}日`, QD });
  if (opt.quick) {
    const d0 = new Date();
    fillQuick(wquick, opt.quick, { font, protect: !!opt.protect, date: `${d0.getFullYear()}年${d0.getMonth() + 1}月${d0.getDate()}日`, title: K.title });
  }
  if (opt.docs) {
    const date = `${dt.getFullYear()}年${dt.getMonth() + 1}月${dt.getDate()}日`;
    // 図（模式図）は、ブックに登録して、各シートに絵として置く
    const images = {};
    for (const [key, im] of Object.entries(opt.docImages || {})) images[key] = wb.addImage({ base64: im.base64, extension: im.ext || "png" });
    fillDocScreen(wdoc, opt.docs, { font, protect: !!opt.protect, name: nameDoc, images });
    fillDocPrint(wdocp, opt.docs, { font, protect: !!opt.protect, name: nameDocP, date, title: opt.docs.title, images });
  }
  // 検索シートから使い方へ飛べるリンク
  ws.getCell("F1").value = { formula: `HYPERLINK("#"&"'${nameG.replace(/'/g, "''")}'!A1","使い方はこちら →")`, result: "使い方はこちら →" };
  ws.getCell("F1").font = { ...font, color: { argb: "FF1A73E8" }, underline: true };
  if (opt.protect) {
    // 入力できるのは、黄色い検索欄とプルダウンだけ（パスワードなし。「校閲→シート保護の解除」で外せる）
    for (const a of ["B2", "C2", "D2", "E2", "B3"]) ws.getCell(a).protection = { locked: false };
    ws.sheetProtection = { sheet: true };
  }
  const idx = wb.worksheets.indexOf(ws);
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: idx, visibility: "visible" }];
  // シートの並び（会社のExcelに差し込むときの順番）
  const list = [nameS, nameG, nameP, ...(nameQuick ? [nameQuick] : []), ...(nameDoc ? [nameDoc, nameDocP] : []), nameD];
  return { wb, nameS, nameG, nameP, nameD, nameDoc, nameDocP, nameQuick, list };
}
