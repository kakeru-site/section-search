// ユーザー（得意先）検索と、部品名→部品番号の検索を、関数だけのExcelにして書き出す（オイル・ケミカルのExcelと同じデザイン）
// 入力は黄色いセルだけ。マクロなし・Excel 2016 までの関数だけ。
//  ・ユーザー検索：名前（漢字・ひらがな・カタカナ・会社名の一部・担当者名）・電話番号・コード・特記で、ユーザーを出す
//    電話を折り返す時・問い合わせ・売り上げる時に、すばやく引くためのもの
//  ・部品検索：部品名（の一部・ひらがなでもOK）か、部品番号の一部から、部品番号と部品名を出す
import { fillGuide } from "./oilxlsx.js?v=20261020h";

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const hair = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const font = { name: "Yu Gothic UI", size: 11 };
const quote = (n) => `'${n.replace(/'/g, "''")}'`;

let FN = "DBCS"; // 全角にそろえる関数（確認用に、別の名前に替えられる）
export const LK_MAX = 10000; // データを見る行数（2行目〜）
const SHOW = 100; // 検索結果に並べる件数
const STRONG = 30; // 最初から枠を整えてある行数

// ひらがな⇄カタカナ（SUBSTITUTEを約30個ずつ、3つのセルに分けて重ねる。関数のネストの上限（64）を超えないため）
const HIRA = [];
for (let c = 0x3041; c <= 0x3096; c++) HIRA.push(String.fromCharCode(c));
const toKata = (ch) => String.fromCharCode(ch.charCodeAt(0) + 0x60);
const chain = (pairs, inner) => pairs.reduce((acc, [a, b]) => `SUBSTITUTE(${acc},"${a}","${b}")`, inner);
const chunks = (arr, n) => {
  const out = [];
  const size = Math.ceil(arr.length / n);
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
};
const H2K = chunks(HIRA.map((h) => [h, toKata(h)]), 3);
const K2H = chunks(HIRA.map((h) => [toKata(h), h]), 3);

const guideBlocks = (protect, nm) => [
  ["title", "ユーザー検索・部品検索の使い方"],
  ["sub", "電話を折り返す時・問い合わせ・売り上げる時に、ユーザーの名前や電話番号から「ユーザーコード・担当・特記」を、部品名や部品番号の一部から「部品番号」を、すぐに探すブックです。黄色いところに言葉を入れるだけ。"],
  ["h", "まずはこれだけ"],
  ["step", "①", `「${nm.user}」か「${nm.part}」のシートを開きます`],
  ["step", "②", "上の黄色いところに、探す言葉を入れます。入れたそばから、下に候補が出ます"],
  ["step", "③", "いちばん上の候補が、すぐ下の「詳しく見る」に、大きく出ます。コード・担当・電話番号・特記が、ひと目で見られます"],
  ["b", "言葉を消すときは、黄色いところを選んで Delete キーを押します"],
  ["h", `「${nm.user}」シート`],
  ["b", "探せるもの：名前（漢字・ひらがな・カタカナ）、会社名の一部、担当者名、ユーザーコード、電話番号（ハイフンなしでもOK）、特記の言葉"],
  ["b", "半角・全角、大文字・小文字は、気にしなくてOKです。ひらがなで入れても、カタカナで入れても、同じように探せます"],
  ["b", "「もう1つの言葉」に入れると、両方に当てはまるものだけに絞れます（例：「運送」と「田中」）"],
  ["b", "「担当で絞る」に、担当の数字（1〜2桁）を入れると、その担当のユーザーだけが出ます。言葉と合わせて使えます"],
  ["b", "候補の並び順：コードや電話番号が、ぴったり合うもの → 名前の先頭が合うもの → 途中が合うもの"],
  ["b", "「詳しく見る」の番号を変えると、その番号の人が、大きく出ます（初期値は1番）"],
  ["b", "特記があるユーザーは、特記の欄が、黄色くなります。見落とさないようにしてください"],
  ["b", `漢字の名前を「ひらがなで」探したいときは、「${nm.userData}」シートの「よみ」の列に、よみ（ひらがな）を入れておきます`],
  ["h", `「${nm.part}」シート`],
  ["b", "部品名の一部、ひらがな（カタカナの名前もOK）、部品番号の一部で探せます。ハイフンなしでもOKです"],
  ["b", `漢字の部品名を「ひらがなで」探したいときは、「${nm.partData}」シートの「よみ」の列に入れておきます`],
  ["h", "データを入れる（更新する）"],
  ["step", "①", `「${nm.userData}」シートのA〜F列に、ユーザーコード・担当・正式名称・よみ（なくてもOK）・電話番号・特記を貼ります（見出しは1行目）`],
  ["step", "②", `「${nm.partData}」シートのA〜C列に、部品番号・部品名・よみ（なくてもOK）を貼ります`],
  ["step", "③", "右の灰色の列は、数式です。消さないでください。貼り付けるだけで、検索に反映されます"],
  ["b", `最初に入っている行は、見本です。本物を貼る前に、見本を消してください`],
  ["b", `どちらのデータも、${LK_MAX.toLocaleString()}行目まで見ます。候補は、上から${SHOW}件まで出ます。多すぎるときは、言葉を足して絞ってください`],
  ["b", "言葉を何も入れていないときは、候補は出ません（動きを軽くするため）"],
  ["h", "困ったとき"],
  ["b", "何も出ない・動かない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", "探しても出ない → 全角・半角、スペース、ハイフンは自動で整えています。言葉を短くして試してください（例：「いちい自動車」→「自動車」）"],
  ["b", "ひらがなで出ない → データの「よみ」の列が、入っているか見てください"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "ユーザーの名前・電話番号は、個人情報を含みます。社外に出さないでください"],
  ...(protect
    ? [
        ["b", "数式を間違って消さないように、検索のシートは保護してあります。入力できるのは黄色い欄だけです（データのシートは、保護していません）"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

const dbr = (sheet, col) => `${quote(sheet)}!$${col}$2:$${col}$${LK_MAX + 1}`;
const label = (ws, addr, text, o = {}) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, bold: true, ...o.font };
  c.alignment = { vertical: "middle", shrinkToFit: true, ...o.alignment };
};
const yellow = (ws, r, c1, c2, o = {}) => {
  if (c2 > c1) ws.mergeCells(r, c1, r, c2);
  for (let i = c1; i <= c2; i++) {
    const c = ws.getCell(r, i);
    c.fill = solid(IN);
    c.border = box;
    c.protection = { locked: false };
  }
  const c = ws.getCell(r, c1);
  c.font = { ...font, size: o.size || 16, bold: true };
  c.alignment = { vertical: "middle", horizontal: o.center ? "center" : "left", indent: o.center ? 0 : 1 };
  if (o.fmt) c.numFmt = o.fmt;
};

// ---- 検索シート ----
// cfg: { title, data, cols:[{head,col,width,mono,center,big,wrap,note}], rankCol, filter?:{label,dataCol}, hint1, hint2 }
function fillFinder(ws, nm, cfg, protect) {
  const cols = cfg.cols;
  const nC = cols.length;
  const lastCol = 2 + nC; // 最後の列（B=No, C.. が項目）
  ws.columns = [{ width: 2 }, { width: 6 }, ...cols.map((c) => ({ width: c.width })), { width: 2 }];
  for (let i = 16; i <= 27; i++) ws.getColumn(i).hidden = true; // 作業用（P〜AA）
  // タイトル
  ws.mergeCells(1, 2, 1, lastCol - 1);
  const t = ws.getCell(1, 2);
  t.value = cfg.title;
  t.font = { ...font, size: 20, bold: true, color: { argb: DK } };
  t.alignment = { vertical: "middle", indent: 1 };
  for (let i = 1; i <= lastCol + 1; i++) ws.getCell(1, i).fill = solid(AC);
  const g = ws.getCell(1, lastCol);
  g.value = { formula: `HYPERLINK("#"&"${quote(nm.guide)}!A1","使い方 →")` };
  g.font = { ...font, color: { argb: "FF1A4FA0" }, underline: true };
  g.alignment = { vertical: "middle", horizontal: "right" };
  ws.getRow(1).height = 40;
  ws.getRow(2).height = 8;
  // 入力：言葉1・言葉2・（担当）
  const inEnd = lastCol; // 入力の右端
  const rows = [
    [3, cfg.label1, cfg.hint1],
    [4, "もう1つの言葉（絞り込み）", cfg.hint2],
  ];
  for (const [r, text] of rows) {
    ws.mergeCells(r, 2, r, 3);
    label(ws, `B${r}`, text, { font: { size: 12 } });
    yellow(ws, r, 4, inEnd);
    ws.getRow(r).height = 36;
  }
  const run = dbr(cfg.data, cfg.rankCol);
  const msg = `IF($X$3=0,"← ${cfg.filter ? "言葉か担当" : "言葉"}を入れてください（${cfg.hint1}）",IF(COUNT(${run})=0,"見つかりませんでした",IF(COUNT(${run})>${SHOW},COUNT(${run})&"件（上の${SHOW}件を表示。言葉を足して絞ってください）",COUNT(${run})&"件 見つかりました")))`;
  if (cfg.filter) {
    ws.mergeCells(5, 2, 5, 3);
    label(ws, "B5", cfg.filter.label, { font: { size: 12 } });
    yellow(ws, 5, 4, 4, { center: true, size: 18 });
    ws.mergeCells(5, 5, 5, lastCol);
    const c = ws.getCell("E5");
    c.value = { formula: msg };
    c.font = { ...font, size: 13, bold: true, color: { argb: "FFB06A00" } };
    c.alignment = { vertical: "middle", indent: 1 };
    ws.getRow(5).height = 32;
  } else {
    ws.mergeCells(5, 2, 5, 3);
    label(ws, "B5", "見つかった数", { font: { size: 12 } });
    ws.mergeCells(5, 4, 5, lastCol);
    const c = ws.getCell("D5");
    c.value = { formula: msg };
    c.font = { ...font, size: 13, bold: true, color: { argb: "FFB06A00" } };
    c.alignment = { vertical: "middle", indent: 1 };
    ws.getRow(5).height = 28;
  }
  ws.getRow(6).height = 8;
  // 作業用：整えた言葉
  for (const r of [3, 4]) {
    const raw = `$D$${r}`;
    ws.getCell(`P${r}`).value = { formula: `LOWER(SUBSTITUTE(SUBSTITUTE(${FN}(${raw}&""),"　","")," ",""))` };
    H2K.forEach((pairs, i) => (ws.getCell(`${["Q", "R", "S"][i]}${r}`).value = { formula: chain(pairs, i === 0 ? `$P${r}` : `$${["Q", "R", "S"][i - 1]}${r}`) }));
    K2H.forEach((pairs, i) => (ws.getCell(`${["T", "U", "V"][i]}${r}`).value = { formula: chain(pairs, i === 0 ? `$P${r}` : `$${["T", "U", "V"][i - 1]}${r}`) }));
    ws.getCell(`W${r}`).value = { formula: `SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($P${r},"－",""),"‐",""),"−","")` };
  }
  // X3：検索を動かすか／AA3：担当（そろえたもの）／Z3：詳しく見る番号／Z4：その行
  ws.getCell("X3").value = { formula: cfg.filter ? `IF(AND($P$3="",$P$4="",$D$5=""),0,1)` : `IF(AND($P$3="",$P$4=""),0,1)` };
  if (cfg.filter) ws.getCell("AA3").value = { formula: `IF($D$5="","",IFERROR(VALUE($D$5),$D$5&""))` };
  // 詳しく見る
  const CR = 7;
  ws.getCell(`B${CR}`).value = 1;
  yellow(ws, CR, 2, 2, { center: true, size: 16 });
  ws.getCell(`B${CR}`).numFmt = "0";
  ws.mergeCells(CR, 3, CR, lastCol);
  const cb = ws.getCell(`C${CR}`);
  cb.value = "◀ この番号の人を、下に大きく表示します（番号を変えられます）";
  cb.font = { ...font, size: 11, bold: true, color: { argb: "FFFFFFFF" } };
  for (let i = 3; i <= lastCol; i++) ws.getCell(CR, i).fill = solid(DK);
  cb.alignment = { vertical: "middle", indent: 1 };
  ws.getRow(CR).height = 28;
  ws.getCell("Z3").value = { formula: `IF(ISNUMBER($B$${CR}),MAX(1,ROUND($B$${CR},0)),1)` };
  ws.getCell("Z4").value = { formula: `IFERROR(MATCH($Z$3,${run},0),"")` };
  // 詳しく見る：見出し小＋大きな値
  const L1 = CR + 1;
  const V1 = CR + 2;
  ["No", ...cols.map((c) => c.head)].forEach((h, i) => {
    const c = ws.getCell(L1, 2 + i);
    c.value = h;
    c.font = { ...font, size: 9, bold: true, color: { argb: "FF4A4F57" } };
    c.fill = solid("FFEDEFF2");
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle" };
  });
  ws.getRow(L1).height = 18;
  const nb = ws.getCell(`B${V1}`);
  nb.value = { formula: `IF($Z$4="","",$Z$3)` };
  nb.font = { ...font, size: 16, bold: true, color: { argb: "FF7A7F87" } };
  nb.alignment = { horizontal: "center", vertical: "middle" };
  nb.border = box;
  cols.forEach((col, i) => {
    const c = ws.getCell(V1, 3 + i);
    c.value = { formula: `IF($Z$4="","",INDEX(${dbr(cfg.data, col.col)},$Z$4)&"")` };
    c.font = col.mono ? { name: "Consolas", size: 20, bold: true } : { ...font, size: col.wrap ? 12 : col.big || 16, bold: !col.wrap };
    c.alignment = { vertical: "middle", horizontal: col.center ? "center" : "left", wrapText: true, indent: col.center ? 0 : 1 };
    c.border = box;
  });
  ws.getRow(V1).height = 64;
  ws.getRow(V1 + 1).height = 10;
  // 一覧
  const HR = V1 + 2;
  ["No", ...cols.map((c) => c.head)].forEach((h, i) => {
    const c = ws.getCell(HR, 2 + i);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(HR).height = 28;
  const F = HR + 1;
  const rowH = cols.some((c) => c.wrap) ? 36 : 26;
  for (let k = 1; k <= SHOW; k++) {
    const r = F + k - 1;
    const strong = k <= STRONG;
    ws.getCell(`Y${r}`).value = { formula: `IFERROR(MATCH(${k},${run},0),"")` };
    ws.getCell(`B${r}`).value = { formula: `IF($Y${r}="","",${k})` };
    cols.forEach((col, i) => {
      const c = ws.getCell(r, 3 + i);
      c.value = { formula: `IF($Y${r}="","",INDEX(${dbr(cfg.data, col.col)},$Y${r})&"")` };
      c.font = col.mono ? { name: "Consolas", size: 13, bold: true } : { ...font, size: 12, bold: !!col.center };
      c.alignment = { vertical: "middle", horizontal: col.center ? "center" : "left", wrapText: true, indent: col.center ? 0 : 1 };
      c.border = strong ? box : hair;
    });
    const b = ws.getCell(`B${r}`);
    b.font = { ...font, color: { argb: "FF7A7F87" } };
    b.alignment = { vertical: "middle", horizontal: "center" };
    b.border = strong ? box : hair;
    ws.getRow(r).height = rowH;
  }
  ws.getColumn(25).hidden = true;
  const lastRow = F + SHOW - 1;
  const colL = (n) => String.fromCharCode(64 + n);
  // しま模様・「詳しく見る」の行・特記
  ws.addConditionalFormatting({
    ref: `B${F}:${colL(lastCol)}${lastRow}`,
    rules: [
      { type: "expression", formulae: [`AND($B${F}<>"",$B${F}=$Z$3)`], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFE680" } }, font: { bold: true } } },
      { type: "expression", formulae: [`AND($B${F}<>"",MOD($B${F},2)=0)`], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF4F5F7" } } } },
    ],
  });
  const noteCol = cols.findIndex((c) => c.note);
  if (noteCol >= 0) {
    const L = colL(3 + noteCol);
    const amber = { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFF3C4" } }, font: { color: { argb: "FF8A5A00" } } };
    ws.addConditionalFormatting({ ref: `${L}${F}:${L}${lastRow}`, rules: [{ type: "expression", formulae: [`${L}${F}<>""`], style: amber }] });
    ws.addConditionalFormatting({ ref: `${L}${V1}`, rules: [{ type: "expression", formulae: [`${L}${V1}<>""`], style: amber }] });
  }
  ws.views = [{ state: "frozen", ySplit: HR, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: `${HR}:${HR}` };
  if (protect) ws.sheetProtection = { sheet: true, formatColumns: true, formatRows: true };
}

// ---- データシート ----
// cfg: { inHeads, inWidths, hay(r), score(r), hayCol, flagCol, scoreCol, rankCol, finder, filterCol? }
function fillData(ws, nm, cfg, rows) {
  const n = cfg.inHeads.length;
  ws.columns = [...cfg.inWidths.map((width) => ({ width })), { width: 3 }, { width: 24 }, { width: 6 }, { width: 6 }, { width: 6 }];
  cfg.inHeads.forEach((h, i) => {
    const c = ws.getCell(1, i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  [["検索用（消さない）", cfg.hayCol], ["一致", cfg.flagCol], ["近さ", cfg.scoreCol], ["順位", cfg.rankCol]].forEach(([h, col]) => {
    const c = ws.getCell(`${col}1`);
    c.value = h;
    c.font = { ...font, size: 9, bold: true, color: { argb: "FF7A7F87" } };
    c.fill = solid("FFEDEFF2");
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(1).height = 30;
  const S = quote(cfg.finder);
  const one = (row, r) =>
    `OR(${S}!$P$${row}="",ISNUMBER(SEARCH(${S}!$P$${row},$${cfg.hayCol}${r})),ISNUMBER(SEARCH(${S}!$S$${row},$${cfg.hayCol}${r})),ISNUMBER(SEARCH(${S}!$V$${row},$${cfg.hayCol}${r})),ISNUMBER(SEARCH(${S}!$W$${row},$${cfg.hayCol}${r})))`;
  for (let r = 2; r <= LK_MAX + 1; r++) {
    const filt = cfg.filterCol ? `,OR(${S}!$AA$3="",IFERROR(VALUE($${cfg.filterCol}${r}),$${cfg.filterCol}${r}&"")=${S}!$AA$3)` : "";
    ws.getCell(`${cfg.hayCol}${r}`).value = { formula: cfg.hay(r) };
    ws.getCell(`${cfg.flagCol}${r}`).value = { formula: `IF($${cfg.hayCol}${r}="","",IF(${S}!$X$3=0,0,IF(AND(${one(3, r)},${one(4, r)}${filt}),1,0)))` };
    ws.getCell(`${cfg.scoreCol}${r}`).value = { formula: `IF($${cfg.flagCol}${r}=1,${cfg.score(r)},"")` };
    ws.getCell(`${cfg.rankCol}${r}`).value = { formula: `IF($${cfg.flagCol}${r}=1,COUNTIF($${cfg.scoreCol}$2:$${cfg.scoreCol}$${LK_MAX + 1},">"&$${cfg.scoreCol}${r})+COUNTIF($${cfg.scoreCol}$2:$${cfg.scoreCol}${r},$${cfg.scoreCol}${r}),"")` };
  }
  rows.forEach((row, i) => row.forEach((v, j) => v !== null && v !== undefined && (ws.getCell(i + 2, j + 1).value = v)));
  ws.getCell("A1").note = cfg.note;
  ws.views = [{ state: "frozen", ySplit: 1, showGridLines: true }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: n } };
  const c = ws.getCell(1, n + 7);
  c.value = "← 見本の行は、本物を貼る前に消してください。右の灰色の列は数式です（消さないでください）";
  c.font = { ...font, size: 9, bold: true, color: { argb: "FFD93025" } };
}

// opt: { protect, sample:boolean, names?, dbcsName? }
export function buildLookupWorkbook(ExcelJS, opt = {}) {
  FN = opt.dbcsName || "DBCS";
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = {
    user: opt.names?.user || "ユーザー検索",
    part: opt.names?.part || "部品検索",
    userData: opt.names?.userData || "ユーザーデータ",
    partData: opt.names?.partData || "部品データ",
    guide: opt.names?.guide || "検索の使い方",
  };
  const v = (name, first) => wb.addWorksheet(name, { views: [{ showGridLines: false, ...(first ? { tabSelected: true } : {}) }] });
  const wsU = v(nm.user, true);
  const wsP = v(nm.part);
  const wsUD = wb.addWorksheet(nm.userData);
  const wsPD = wb.addWorksheet(nm.partData);
  const wsG = v(nm.guide);
  const protect = !!opt.protect;
  const S = (sh, c, row) => `${quote(sh)}!$${c}$${row}`;

  // ユーザーデータ：A コード B 担当 C 正式名称 D よみ E 電話番号 F 特記 ／ H 検索用 I 一致 J 近さ K 順位
  const norm = (x) => `${FN}(LOWER(SUBSTITUTE(SUBSTITUTE(${x},"　","")," ","")))`;
  const userHay = (r) => `IF($A${r}&$C${r}="","",${norm(`$A${r}&$C${r}&$D${r}&$F${r}`)}&"｜"&${FN}(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($E${r},"-","")," ",""),"　","")))`;
  const userScore = (r) => {
    const f = nm.user;
    const phone = `${FN}(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($E${r},"-","")," ",""),"　",""))`;
    const code = norm(`$A${r}`);
    const name = norm(`$C${r}&$D${r}`);
    return `IF(OR(AND(${S(f, "W", 3)}<>"",OR(${code}=${S(f, "W", 3)},${phone}=${S(f, "W", 3)})),AND(${S(f, "P", 3)}<>"",${code}=${S(f, "P", 3)})),4,IF(OR(IFERROR(SEARCH(${S(f, "P", 3)},${name})=1,FALSE),IFERROR(SEARCH(${S(f, "S", 3)},${name})=1,FALSE),IFERROR(SEARCH(${S(f, "V", 3)},${name})=1,FALSE),IFERROR(SEARCH(${S(f, "W", 3)},${phone})=1,FALSE),IFERROR(SEARCH(${S(f, "P", 3)},${code})=1,FALSE)),3,2))`;
  };
  // 部品データ：A 部品番号 B 部品名 C よみ ／ E 検索用 F 一致 G 近さ H 順位
  const partHay = (r) => `IF($A${r}&$B${r}="","",${norm(`$A${r}&$B${r}&$C${r}`)}&"｜"&${FN}(SUBSTITUTE(SUBSTITUTE($A${r},"-","")," ","")))`;
  const partScore = (r) => {
    const f = nm.part;
    const codeK = `${FN}(SUBSTITUTE(SUBSTITUTE(LOWER($A${r}),"-","")," ",""))`;
    const name = norm(`$B${r}&$C${r}`);
    return `IF(AND(${S(f, "W", 3)}<>"",${codeK}=${S(f, "W", 3)}),4,IF(OR(IFERROR(SEARCH(${S(f, "S", 3)},${name})=1,FALSE),IFERROR(SEARCH(${S(f, "V", 3)},${name})=1,FALSE),IFERROR(SEARCH(${S(f, "P", 3)},${name})=1,FALSE),IFERROR(SEARCH(${S(f, "W", 3)},${codeK})=1,FALSE)),3,2))`;
  };

  fillFinder(
    wsU,
    nm,
    {
      title: "ユーザー検索",
      label1: "名前・電話番号・コード・特記",
      hint1: "名前（漢字・ひらがな・カタカナ）、会社名の一部、電話番号、コード、特記",
      hint2: "例：「運送」＋「田中」",
      data: nm.userData,
      rankCol: "K",
      filter: { label: "担当で絞る（数字）", dataCol: "B" },
      cols: [
        { head: "ユーザーコード", col: "A", width: 15, mono: true, center: true },
        { head: "担当", col: "B", width: 7, center: true, big: 22 },
        { head: "正式名称", col: "C", width: 36 },
        { head: "電話番号", col: "E", width: 18, big: 15 },
        { head: "特記", col: "F", width: 66, wrap: true, note: true },
      ],
    },
    protect,
  );
  fillFinder(
    wsP,
    nm,
    {
      title: "部品検索",
      label1: "部品名・部品番号",
      hint1: "部品名の一部（ひらがなもOK）、部品番号の一部（ハイフンなしもOK）",
      hint2: "例：「ブレーキ」＋「パッド」",
      data: nm.partData,
      rankCol: "H",
      cols: [
        { head: "部品番号", col: "A", width: 24, mono: true, center: true },
        { head: "部品名", col: "B", width: 56, big: 16 },
      ],
    },
    protect,
  );
  let users = [];
  let parts = [];
  if (opt.sample) {
    const cos = [["見本自動車", "みほんじどうしゃ"], ["サンプル商事", "さんぷるしょうじ"], ["テスト整備工場", "てすとせいびこうじょう"], ["例示モータース", "れいじもーたーす"], ["ダミー運送", "だみーうんそう"], ["仮名車体", "かめいしゃたい"]];
    const people = [["山田", "やまだ"], ["鈴木", "すずき"], ["佐藤", "さとう"], ["田中", "たなか"], ["高橋", "たかはし"]];
    const notes = ["", "", "月末締め・翌月末払い。納品は午前中のみ", "", "", "電話は担当直通へ。折り返しは17時まで", "", "", "支払いは現金のみ（掛け売り不可）", "", "", "夜間は携帯へ（080から始まる番号）", "", "", ""];
    for (let i = 0; i < 30; i++) {
      const [co, cy] = cos[i % cos.length];
      const [pe, py] = people[(i * 2) % people.length];
      users.push([`Z${String(i + 1).padStart(3, "0")}`, 1 + ((i * 7) % 12), `${co}${Math.floor(i / cos.length) + 1}号店 ${pe}さん`, `${cy}${Math.floor(i / cos.length) + 1}ごうてん ${py}さん`, `000-0000-${String(1000 + i * 3)}`, notes[i % notes.length]]);
    }
    const pn = [["六角ボルト", "ろっかくぼると"], ["ワッシャー", ""], ["ベアリング", ""], ["エンジンオイル", ""], ["オイルフィルター", ""], ["ブレーキパッド", ""], ["点火プラグ", "てんかぷらぐ"], ["燃料ポンプ", "ねんりょうぽんぷ"], ["冷却水ホース", "れいきゃくすいほーす"], ["ワイパーゴム", ""]];
    for (let i = 0; i < 30; i++) {
      const [n, y] = pn[i % pn.length];
      parts.push([`S0000-${String(1000 + i * 5)}`, `${n}${Math.floor(i / pn.length) ? "（" + (Math.floor(i / pn.length) + 1) + "）" : ""}`, y]);
    }
  }
  fillData(
    wsUD,
    nm,
    { inHeads: ["ユーザーコード", "担当", "正式名称（会社名 担当者名）", "よみ（ひらがな・なくてもOK）", "電話番号", "特記"], inWidths: [14, 7, 36, 30, 18, 60], hayCol: "H", flagCol: "I", scoreCol: "J", rankCol: "K", finder: nm.user, filterCol: "B", hay: userHay, score: userScore, note: "ここに、ユーザーの一覧を貼ります（A〜F列、見出しは1行目）。担当は、1〜2桁の数字。よみは、なくてもOKです" },
    users,
  );
  fillData(
    wsPD,
    nm,
    { inHeads: ["部品番号", "部品名", "よみ（ひらがな・なくてもOK）"], inWidths: [20, 40, 30], hayCol: "E", flagCol: "F", scoreCol: "G", rankCol: "H", finder: nm.part, hay: partHay, score: partScore, note: "ここに、部品番号と部品名を貼ります（A〜C列、見出しは1行目）。よみは、なくてもOKです" },
    parts,
  );
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list: [nm.user, nm.part, nm.userData, nm.partData, nm.guide], names: nm };
}
