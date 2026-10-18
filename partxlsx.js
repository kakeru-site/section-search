// 外装検索・部品検索（内装・機関）を、1つのExcelに書き出す（オイル・ケミカルのExcelと同じデザイン）
// シート：①探せる一覧（キーワードで黄色く光る・フィルター） ②部位カタログ（絵と解説つき・A4印刷） ③使い方
import { fillGuide } from "./oilxlsx.js";

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const hair = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const font = { name: "Yu Gothic UI", size: 11 };
const quote = (n) => `'${n.replace(/'/g, "''")}'`;
const dispW = (s) => [...String(s ?? "")].reduce((a, ch) => a + (ch.charCodeAt(0) > 255 ? 2 : 1), 0);
const linesOf = (text, width) =>
  String(text ?? "")
    .split("\n")
    .reduce((a, t) => a + Math.max(1, Math.ceil((dispW(t) * 1.08) / Math.max(4, width - 1.5))), 0);
const norm = (s) => String(s ?? "").normalize("NFKC").toLowerCase();

const guideBlocks = (protect, nm) => [
  ["title", "外装・部品検索の使い方"],
  ["sub", "外装の部位・部品（図番つき）と、エンジンや足回りなどの部品（図番つき）を、1冊にまとめたブックです。絵つきのカタログと、キーワードで探せる一覧が入っています。"],
  ["h", "まずはこれだけ"],
  ["step", "①", `「${nm.data}」シートの上の黄色い欄に、探したい言葉（例：ヘッドライト、ドラシャ、ワイパーゴム）を入れます`],
  ["step", "②", "当てはまる部品だけが、下に並びます（ほかは消えます）。見つかった件数は、すぐ横に出ます。言葉を消すと、全部が出ます"],
  ["step", "③", `部位の絵や解説を見たいときは、「${nm.cat}」シートを開きます（分類ごとに並んでいます）`],
  ["h", `「${nm.data}」シート`],
  ["b", "言葉は3つまで入れられます。3つとも当てはまる部品だけが並びます（1つだけでもOK）"],
  ["b", "検索の対象は、部位名・部品名・英語名・図番・右左・備考（お客さんの呼び方・場所）と、部位の呼ばれ方です"],
  ["b", "半角・全角は区別されます。うまく出ないときは、短い言葉（例：「ミラー」）で試してください。英数字は半角、カタカナは全角で入れます"],
  ["b", `全部の一覧は「${nm.raw}」シートに入っています。見出しの▼で、区分・分類・部位・右左を絞り込めます（フィルター）`],
  ["b", "備考には、お客さんがよく言う呼び方と、部品がある場所を入れてあります。「要確認」と書いてあるものは、位置を図から推測したものです"],
  ["h", `「${nm.cat}」シート`],
  ["b", "部位ごとに、形のイメージの絵・やさしい説明・お客さんの呼ばれ方・部品の一覧（図番つき）をまとめています"],
  ["b", "絵は、場所と形をつかむための簡単なイメージです。実物とは細部が違います"],
  ["b", "A4よこで印刷できます（部位ごとに途中で切れないように区切っています）"],
  ["h", "困ったとき"],
  ["b", "何も出ない・並ばない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", "言葉を入れてから、少し待つと、並び直します（数が多いので、1〜2秒かかることがあります）"],
  ["b", "備考が長くて途中で切れるときは、そのセルをクリックすると、全文が上に出ます"],
  ["b", "図番や名称が足りない → 店で使っているサイトの「外装検索」「部品検索」で追加してから、もう一度書き出してください"],
  ["h", "これだけは守ってください"],
  ["b", "図番は、車種・年式・仕様で変わることがあります。注文の前に、車台番号などで確認してください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ...(protect
    ? [
        ["b", "間違って消さないように、シートは保護してあります。入力できるのは黄色い欄だけです（フィルターは使えます）"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

// ---- ① 探せる一覧：言葉を入れると、当てはまる部品だけが並ぶ（数式。マクロなし） ----
const RAW_COLS = [11, 14, 30, 38, 34, 12, 7, 62];
function fillRaw(ws, nm, rows, protect) {
  const sq = quote(nm.data);
  const kw = (c) => `IF(${sq}!$${c}$3="","?",${sq}!$${c}$3)`;
  ws.columns = [{ width: 2 }, ...RAW_COLS.map((width) => ({ width })), { width: 10 }, { width: 8 }];
  ["区分", "分類", "部位", "名称", "英語名", "図番", "右左", "備考（お客さんの呼び方・場所）"].forEach((h, i) => {
    const c = ws.getCell(1, i + 2);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getCell(1, 10).value = "検索用";
  ws.getRow(1).height = 28;
  rows.forEach((o, k) => {
    const r = 2 + k;
    const v = [o.kind, o.cat, o.group, o.nameJa, o.nameEn, o.code, o.side, o.note];
    v.forEach((x, i) => {
      const c = ws.getCell(r, i + 2);
      c.value = x || "";
      c.font = i === 5 ? { name: "Consolas", size: 11, bold: true } : font;
      c.border = hair;
      c.alignment = { vertical: "middle", wrapText: true, horizontal: i === 6 || i === 0 ? "center" : "left" };
    });
    ws.getCell(r, 10).value = norm(o.hay);
    // 隠しの列K：ここまでに、言葉に当てはまった部品が何件あったか（上から数えた通し番号）
    const hit = ["C", "F", "I"].map((c) => `ISNUMBER(SEARCH(${kw(c)},$J${r}))`).join(",");
    ws.getCell(r, 11).value = { formula: `IF(AND(${hit}),1,0)+N(K${r - 1})`, result: k + 1 };
    ws.getRow(r).height = Math.max(20, Math.max(...v.map((x, i) => linesOf(x, RAW_COLS[i]))) * 16 + 5);
  });
  ws.getColumn(10).hidden = true;
  ws.getColumn(11).hidden = true;
  ws.autoFilter = { from: { row: 1, column: 2 }, to: { row: rows.length + 1, column: 9 } };
  ws.views = [{ state: "frozen", ySplit: 1, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "1:1" };
  if (protect) ws.sheetProtection = { sheet: true, autoFilter: false, sort: false, formatColumns: true, formatRows: true };
}
function fillSearch(ws, nm, rows, protect) {
  const N = rows.length;
  const W = { b: 11, c: 14, d: 30, e: 38, f: 34, g: 12, h: 7, i: 62 };
  ws.columns = [
    { width: 2 }, { width: W.b }, { width: W.c }, { width: W.d }, { width: W.e }, { width: W.f }, { width: W.g }, { width: W.h }, { width: W.i }, { width: 10 }, { width: 8 },
  ];
  const raw = quote(nm.raw);
  const rg = (col) => `${raw}!$${col}$2:$${col}$${N + 1}`;
  ws.mergeCells("B1:H1");
  const t = ws.getCell("B1");
  t.value = "外装・部品検索";
  t.font = { ...font, size: 18, bold: true, color: { argb: DK } };
  t.alignment = { vertical: "middle", indent: 1 };
  for (let i = 2; i <= 9; i++) ws.getCell(1, i).fill = solid(AC);
  ws.getCell("I1").value = { formula: `HYPERLINK("#"&"${quote(nm.guide)}!A1","使い方はこちら →")`, result: "使い方はこちら →" };
  ws.getCell("I1").font = { ...font, color: { argb: "FF1A4FA0" }, underline: true };
  ws.getCell("I1").alignment = { vertical: "middle", horizontal: "right" };
  ws.getRow(1).height = 36;
  ws.getRow(2).height = 8;
  const label = (addr, text, o = {}) => {
    const c = ws.getCell(addr);
    c.value = text;
    c.font = { ...font, bold: true };
    c.alignment = { vertical: "middle", shrinkToFit: true, ...o };
  };
  const mk = (addr) => {
    const c = ws.getCell(addr);
    c.fill = solid(IN);
    c.border = box;
    c.font = { ...font, size: 13, bold: true };
    c.alignment = { vertical: "middle", horizontal: "left" };
    c.protection = { locked: false };
  };
  label("B3", "探す言葉");
  ws.mergeCells("C3:D3");
  mk("C3");
  ws.getCell("D3").border = box;
  ws.getCell("D3").fill = solid(IN);
  label("E3", "＋ もう1つ", { horizontal: "right" });
  mk("F3");
  ws.mergeCells("G3:H3");
  label("G3", "＋ もう1つ", { horizontal: "right" });
  mk("I3");
  ws.getRow(3).height = 28;
  // 空の欄は「何でも1文字」の記号（?）にして、全部に当てはまるようにする（ExcelでもLibreOfficeでも同じ結果）
  const idle = `AND($C$3="",$F$3="",$I$3="")`;
  label("B4", "ヒット");
  ws.mergeCells("C4:D4");
  ws.getCell("C4").value = {
    formula: `IF(${idle},"全${N}件を出しています",MAX(${rg("K")})&"件が見つかりました")`,
    result: `全${N}件を出しています`,
  };
  ws.getCell("C4").font = { ...font, size: 12, bold: true, color: { argb: "FFB06A00" } };
  ws.getCell("C4").alignment = { vertical: "middle" };
  ws.getCell("E4").value = "← 当てはまる部品だけが、下に並びます。3つ入れたときは、3つとも当てはまるものだけ";
  ws.getCell("E4").font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
  ws.getCell("E4").alignment = { vertical: "middle" };
  ws.getRow(4).height = 22;
  ws.getRow(5).height = 8;
  ["区分", "分類", "部位", "名称", "英語名", "図番", "右左", "備考（お客さんの呼び方・場所）"].forEach((h, i) => {
    const c = ws.getCell(6, i + 2);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(6).height = 30;
  const FIRST = 7;
  const src = ["B", "C", "D", "E", "F", "G", "H", "I"];
  const vals = (o) => [o.kind, o.cat, o.group, o.nameJa, o.nameEn, o.code, o.side, o.note];
  for (let k = 0; k < N; k++) {
    const r = FIRST + k;
    // 隠しの列K：k番目に見つかった行の番号（なければ空）
    ws.getCell(r, 11).value = {
      formula: `IFERROR(MATCH(${k + 1},${rg("K")},0),"")`,
      result: k + 1,
    };
    const v = vals(rows[k]);
    src.forEach((col, i) => {
      const c = ws.getCell(r, i + 2);
      c.value = { formula: `IF($K${r}="","",INDEX(${rg(col)},$K${r})&"")`, result: v[i] || "" };
      c.font = i === 5 ? { name: "Consolas", size: 11, bold: true } : { ...font, size: 10.5 };
      c.alignment = { vertical: "middle", wrapText: true, horizontal: i === 6 || i === 0 ? "center" : "left" };
    });
    ws.getRow(r).height = 33;
  }
  const LAST = FIRST + N - 1;
  ws.getColumn(11).hidden = true;
  ws.addConditionalFormatting({
    ref: `B${FIRST}:I${LAST}`,
    rules: [{ type: "expression", formulae: [`$K${FIRST}<>""`], style: { border: { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } } } }],
  });
  ws.views = [{ state: "frozen", ySplit: 6, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "6:6" };
  if (protect) ws.sheetProtection = { sheet: true };
}

// ---- ② 部位カタログ（絵つき） ----
function fillCatalog(wb, ws, nm, groups, images, protect) {
  const W = { img: 30, name: 38, code: 14, side: 7, note: 62 };
  ws.columns = [{ width: 2 }, { width: W.img }, { width: W.name }, { width: W.code }, { width: W.side }, { width: W.note }];
  ws.mergeCells("B1:E1");
  const t = ws.getCell("B1");
  t.value = "部位カタログ（絵と解説つき）";
  t.font = { ...font, size: 18, bold: true, color: { argb: DK } };
  t.alignment = { vertical: "middle", indent: 1 };
  for (let i = 2; i <= 6; i++) ws.getCell(1, i).fill = solid(AC);
  ws.getCell("F1").value = { formula: `HYPERLINK("#"&"${quote(nm.guide)}!A1","使い方はこちら →")`, result: "使い方はこちら →" };
  ws.getCell("F1").font = { ...font, color: { argb: "FF1A4FA0" }, underline: true };
  ws.getCell("F1").alignment = { vertical: "middle", horizontal: "right" };
  ws.getRow(1).height = 36;
  ws.getRow(2).height = 8;
  const imgPxW = Math.round(W.img * 7.2);
  const imgPxH = Math.round(imgPxW * 0.61);
  let r = 3;
  const breaks = [];
  let lastKind = "";
  let lastCat = "";
  for (const g of groups) {
    // 区分・分類の見出し
    if (g.kind !== lastKind || g.cat !== lastCat) {
      if (g.kind !== lastKind) {
        ws.mergeCells(r, 2, r, 6);
        const kc = ws.getCell(r, 2);
        kc.value = `【${g.kind}】`;
        kc.font = { ...font, size: 15, bold: true, color: { argb: "FFFFFFFF" } };
        for (let i = 2; i <= 6; i++) ws.getCell(r, i).fill = solid(DK);
        kc.alignment = { vertical: "middle", indent: 1 };
        ws.getRow(r).height = 30;
        if (r > 3) breaks.push(r - 1);
        r++;
      }
      ws.mergeCells(r, 2, r, 6);
      const cc = ws.getCell(r, 2);
      cc.value = `${g.cat}${g.catHint ? "　" + g.catHint : ""}`;
      cc.font = { ...font, size: 12, bold: true, color: { argb: DK } };
      for (let i = 2; i <= 6; i++) ws.getCell(r, i).fill = solid("FFFFF3C4");
      cc.alignment = { vertical: "middle", indent: 1 };
      ws.getRow(r).height = 24;
      r++;
      lastKind = g.kind;
      lastCat = g.cat;
    }
    const start = r;
    // 部位名
    ws.mergeCells(r, 2, r, 6);
    const hc = ws.getCell(r, 2);
    hc.value = g.name + (g.en ? `　${g.en}` : "");
    hc.font = { ...font, size: 14, bold: true, color: { argb: DK } };
    hc.alignment = { vertical: "middle", indent: 1 };
    for (let i = 2; i <= 6; i++) {
      ws.getCell(r, i).fill = solid(AC);
      ws.getCell(r, i).border = box;
    }
    ws.getRow(r).height = 28;
    r++;
    // 説明
    const desc = [g.one, g.alias?.length ? `お客さんの呼ばれ方：${g.alias.slice(0, 14).join("、")}` : ""].filter(Boolean).join("\n");
    const descRowH = Math.max(34, linesOf(desc, W.name + W.code + W.side + W.note) * 17 + 8);
    ws.mergeCells(r, 3, r, 6);
    const dc = ws.getCell(r, 3);
    dc.value = desc || "（説明は準備中です）";
    dc.font = { ...font, size: 10.5, color: { argb: "FF4A4F57" } };
    dc.alignment = { vertical: "middle", wrapText: true, indent: 1 };
    for (let i = 3; i <= 6; i++) ws.getCell(r, i).border = box;
    ws.getRow(r).height = descRowH;
    const imgTop = r;
    r++;
    // 部品の見出し
    ["名称", "図番", "右左", "備考（呼び方・場所）"].forEach((h, i) => {
      const c = ws.getCell(r, i + 3);
      c.value = h;
      c.font = { ...font, bold: true, size: 10, color: { argb: DK } };
      c.fill = solid("FFEDEFF2");
      c.border = box;
      c.alignment = { horizontal: "center", vertical: "middle" };
    });
    ws.getRow(r).height = 20;
    r++;
    const body0 = r;
    const rows = g.parts;
    const used = [];
    rows.forEach((o) => {
      const v = [o.nameJa + (o.nameEn ? `\n${o.nameEn}` : ""), o.code, o.side, o.note];
      v.forEach((x, i) => {
        const c = ws.getCell(r, i + 3);
        c.value = x || "";
        c.font = i === 1 ? { name: "Consolas", size: 11, bold: true } : i === 0 ? { ...font, size: 10.5 } : { ...font, size: 10 };
        c.border = hair;
        c.alignment = { vertical: "middle", wrapText: true, horizontal: i === 1 || i === 2 ? "center" : "left" };
      });
      // 名称の2行目（英語名）は小さく見せるため、1行ぶん多めに
      const h = Math.max(
        20,
        (linesOf(o.nameJa, W.name) + (o.nameEn ? linesOf(o.nameEn, W.name) * 0.85 : 0)) * 16 + 5,
        linesOf(o.note, W.note) * 15 + 5,
      );
      ws.getRow(r).height = h;
      used.push(h);
      r++;
    });
    // 絵の高さぶん、行が足りなければ空き行を足す
    const need = imgPxH * 0.75 + 4; // px → pt
    let have = descRowH + 20 + used.reduce((a, b) => a + b, 0);
    while (have < need) {
      for (let i = 3; i <= 6; i++) ws.getCell(r, i).border = hair;
      ws.getRow(r).height = 18;
      have += 18;
      r++;
    }
    // 左の絵の欄
    ws.mergeCells(imgTop, 2, r - 1, 2);
    const ic = ws.getCell(imgTop, 2);
    ic.border = box;
    ic.fill = solid("FFF7F8FA");
    for (let rr = imgTop; rr < r; rr++) ws.getCell(rr, 2).border = box;
    const id = images?.[g.key];
    if (id !== undefined) ws.addImage(id, { tl: { col: 1.04, row: imgTop - 1 + 0.06 }, ext: { width: imgPxW - 12, height: imgPxH - 8 } });
    // 部品が無いとき
    if (!rows.length) {
      ws.getCell(body0, 3).value = "（部品は、まだ入っていません）";
      ws.getCell(body0, 3).font = { ...font, color: { argb: "FF9AA0A8" } };
    }
    // 部位は途中で切れないよう、前に改ページを入れる（長い部位だけ）
    ws.getRow(r).height = 10;
    r++;
    if ((r - start) > 0 && (r % 1) === 0) {
      /* ページ区切りは、まとまりごと（分類）の前に入れる */
    }
  }
  ws.views = [{ state: "frozen", ySplit: 2, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "1:1" };
  for (const b of breaks) ws.getRow(b).addPageBreak?.();
  if (protect) ws.sheetProtection = { sheet: true };
}

// opt: { protect, names?:{data,cat,guide}, groups:[{kind,cat,catHint,name,en,one,alias,parts:[{nameJa,nameEn,code,side,note}],key}], images:{key:{base64,ext}} }
export function buildPartWorkbook(ExcelJS, opt = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = {
    data: opt.names?.data || "外装・部品検索",
    cat: opt.names?.cat || "部位カタログ",
    raw: opt.names?.raw || "外装・部品データ",
    guide: opt.names?.guide || "外装・部品の使い方",
  };
  const wsD = wb.addWorksheet(nm.data, { views: [{ showGridLines: false, tabSelected: true }] });
  const wsC = wb.addWorksheet(nm.cat, { views: [{ showGridLines: false }] });
  const wsR = wb.addWorksheet(nm.raw, { views: [{ showGridLines: false }] });
  const wsG = wb.addWorksheet(nm.guide, { views: [{ showGridLines: false }] });
  const protect = !!opt.protect;
  const groups = opt.groups || [];
  const images = {};
  for (const [k, im] of Object.entries(opt.images || {})) images[k] = wb.addImage({ base64: im.base64, extension: im.ext || "png" });
  const rows = [];
  for (const g of groups)
    for (const p of g.parts)
      rows.push({
        kind: g.kind,
        cat: g.cat,
        group: g.name,
        nameJa: p.nameJa,
        nameEn: p.nameEn,
        code: p.code,
        side: p.side,
        note: p.note,
        hay: `${g.kind} ${g.cat} ${g.name} ${g.en || ""} ${(g.alias || []).join(" ")} ${p.nameJa} ${p.nameEn || ""} ${p.code} ${p.side || ""} ${p.note || ""}`,
      });
  fillSearch(wsD, nm, rows, protect);
  fillRaw(wsR, nm, rows, protect);
  fillCatalog(wb, wsC, nm, groups, images, protect);
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list: [nm.data, nm.cat, nm.raw, nm.guide], names: nm, rows: rows.length };
}
