// 部品番号ラベルを、関数だけで作れるExcel（オイル・ケミカルのExcelと同じデザイン）
//  ・ラベル作成：部品番号を1つ入れると、同じラベルが28枚（A4いっぱい）、すぐ印刷できる
//  ・まとめて印刷：部品番号を何件かリストに入れると、1件＝1ページで、まとめて印刷できる
//  ・ラベルデータ：在庫リストのExcelを、そのまま丸ごとコピー貼り付けして使う（列の並びは、在庫リストと同じ）
import { fillGuide } from "./oilxlsx.js?v=20261022p";

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const GRAY = "FF969696"; // ラベルの地色（NISSAN純正ラベルの灰色）
const font = { name: "Yu Gothic UI", size: 11 };
const labelFont = (size, o = {}) => ({ name: "ＭＳ Ｐゴシック", size, ...o });
const quote = (n) => `'${n.replace(/'/g, "''")}'`;
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const colLetter = (n) => String.fromCharCode(64 + n);
export const DATA_MAX = 6000; // 在庫リストを見る行数（2行目〜）
export const BATCH_MAX = 30; // 「まとめて印刷」で、一度に作れる件数

// 在庫リストと、まったく同じ列の並び（丸ごとコピー貼り付けできるように）
const STOCK_HEADS = ["拠点コード", "ロケーション．", "常備", "部品番号 (表示)", "部品名称", "取得", "商区", "メーカーロケ", "実在庫数", "委託数", "移管中数", "自送Ｎ", "自送Ｎ＋１", "MAX", "O/P", "MIN", "拠点３Ｄ", "拠点６Ｄ", "原価", "販売単位"];
const LOC_COL = "B"; // ロケーション．
const PN_COL = "D"; // 部品番号 (表示)
// 部品番号を、探しやすく整えた列（検索用）。データシートの中に置くと、在庫リストの
// 元のシートが見出し（20列）より広い範囲を含んでいた場合、「シート全体コピー」の
// 貼り付けで巻き込まれて消えてしまうため、貼り付けの影響がぜったいに及ばない、
// 別の隠しシート（nm.calc）に置く
const NORM_COL = "A";

// ラベル1枚ぶんの行の高さ（純正ラベルの実寸比率がもと。2・6行目は、小さい文字が
// 印刷でつぶれないよう、4行目のあき・7行目の余白から少し分けて広げてある）
const ROW_H = [12.75, 9, 14.25, 4.25, 21.75, 10, 13.5];

// ---- ラベル1枚（3列×7行）を描く ----
// r0,c0：そのラベルの左上（1始まり）。partF／locF：部品番号・ロケーションの中身にする数式（同じシート内）
function paintLabel(ws, r0, c0, partF, locF) {
  const c1 = c0,
    c2 = c0 + 1,
    c3 = c0 + 2;
  const white = (t) => ({ ...labelFont(t) });
  // NISSAN / GENUINE PARTS / Made in japan（灰色・白文字）
  ws.mergeCells(r0, c1, r0 + 1, c1);
  Object.assign(ws.getCell(r0, c1), { value: "NISSAN" });
  ws.getCell(r0, c1).font = labelFont(24, { bold: true, color: { argb: "FFFFFFFF" } });
  ws.getCell(r0, c1).alignment = { horizontal: "center", vertical: "center" };
  ws.getCell(r0, c2).value = "GENUINE PARTS";
  ws.getCell(r0, c2).font = labelFont(7, { bold: true, color: { argb: "FFFFFFFF" } });
  ws.getCell(r0, c2).alignment = { vertical: "center" };
  ws.getCell(r0 + 1, c2).value = "                 Made in japan";
  ws.getCell(r0 + 1, c2).font = labelFont(7, { bold: true, color: { argb: "FFFFFFFF" } });
  ws.getCell(r0 + 1, c2).alignment = { vertical: "center" };
  for (const rr of [r0, r0 + 1]) for (const cc of [c1, c2, c3]) ws.getCell(rr, cc).fill = solid(GRAY);
  // ロケーション（白地）
  const loc = ws.getCell(r0 + 2, c1);
  loc.value = { formula: locF };
  loc.font = labelFont(12);
  loc.alignment = { horizontal: "left", vertical: "center" };
  // 部品番号（白地・大きく）
  ws.mergeCells(r0 + 4, c1, r0 + 4, c2);
  const pn = ws.getCell(r0 + 4, c1);
  pn.value = { formula: partF };
  pn.font = labelFont(22, { bold: true });
  pn.alignment = { horizontal: "distributed", vertical: "distributed" };
  // NISSAN MOTOR CO.LTD（灰色・白文字）
  ws.getCell(r0 + 5, c1).fill = solid(GRAY);
  ws.getCell(r0 + 5, c2).fill = solid(GRAY);
  ws.getCell(r0 + 5, c2).value = "NISSAN MOTOR CO.LTD";
  ws.getCell(r0 + 5, c2).font = labelFont(7, { bold: true, color: { argb: "FFFFFFFF" } });
  ws.getCell(r0 + 5, c2).alignment = { vertical: "center" };
  // 右端の仕切り（縦の灰色の帯）
  for (let rr = r0; rr <= r0 + 6; rr++) ws.getCell(rr, c3).fill = solid(GRAY);
}
// 1ページ＝4列×7行＝28枚。pageTop：そのページの先頭の行
function paintPage(ws, pageTop, partF, locF) {
  for (let col = 0; col < 4; col++) for (let blk = 0; blk < 7; blk++) paintLabel(ws, pageTop + blk * 7, 2 + col * 3, partF, locF);
  for (let blk = 0; blk < 7; blk++) for (let i = 0; i < 7; i++) ws.getRow(pageTop + blk * 7 + i).height = ROW_H[i];
}
function setLabelCols(ws) {
  const w = [];
  for (let i = 0; i < 4; i++) w.push({ width: 18 }, { width: 13.9 }, { width: 3.1 });
  ws.columns = [{ width: 2 }, ...w, { width: 12 }]; // 最後の1列（N）は、画面の「使い方」リンク用（印刷はされない）
}

const guideBlocks = (protect, nm) => [
  ["title", "部品番号ラベル作成の使い方"],
  ["sub", "在庫リストから、部品番号・ロケーションの入った、純正ラベル（NISSAN GENUINE PARTS）を作って印刷するシートです。"],
  ["h", "まずはこれだけ（データを入れる）"],
  ["step", "①", `「${nm.data}」シートに、在庫リストのExcelの中身を、そのまま丸ごとコピーして、貼り付けます（A1から。見出しの1行目も、そのまま上書きしてOKです）`],
  ["b", `列の並びは、在庫リストと同じです（${STOCK_HEADS.slice(0, 4).join("・")}…）。並びを変えずに、そのまま貼り付けてください`],
  ["h", `「${nm.find}」シート（1件ずつ、すぐ印刷）`],
  ["step", "①", "黄色いところに、部品番号を入れます（ハイフンあり・なし、大文字・小文字、どちらでもOK）"],
  ["step", "②", "下に、同じラベルが28枚（A4いっぱい）出ます。そのまま印刷します"],
  ["step", "③", "次の部品番号に入れ替えると、ラベルも自動で変わります"],
  ["b", "見つからないときは、ロケーション・部品番号の欄が「見つかりません」になります。品番か、データの貼り付けを確認してください"],
  ["h", `「${nm.batch}」シート（まとめて、何件か印刷）`],
  ["step", "①", "上の表の、黄色い「部品番号」欄に、必要な分だけ、部品番号を入れます（1行に1件）"],
  ["step", "②", `${BATCH_MAX}件まで入れられます。右に、ロケーションと、印刷ページの番号が出ます`],
  ["step", "③", "下に、入れた件数ぶんのページが、自動でできます（1件＝1ページ＝28枚）"],
  ["step", "④", "印刷するときは、印刷の画面で「ページ指定」に、入れた件数のページ数だけを指定してください（上の表に、ページ数が出ています）。指定しないと、空っぽのページまで印刷されます"],
  ["h", "困ったとき"],
  ["b", "何も出ない・計算されない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", `見つからない → 「${nm.data}」シートに、在庫リストのデータが、正しく貼り付けてあるか確認してください`],
  ["b", "データが古い → 新しい在庫リストを、また丸ごとコピーして、貼り付け直してください（前のデータは上書きされます）"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "ラベルのシートは、見た目を合わせるため、印刷の範囲をあらかじめ決めてあります（B〜L列）。M列より右や、上の入力欄は、印刷されません"],
  ...(protect
    ? [
        ["b", "数式を間違って消さないように、ラベル・使い方のシートは保護してあります。入力できるのは黄色い欄だけです（データのシートは、保護していません）"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

function fillGuideStyle(ws, nm) {
  const t = ws.getCell("B2");
  t.value = nm.title;
  t.font = { ...font, size: 18, bold: true };
}
const label2 = (ws, addr, text, o = {}) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, bold: true, ...o.font };
  c.alignment = { vertical: "middle", ...o.alignment };
};
const yellow = (ws, addr, o = {}) => {
  const c = ws.getCell(addr);
  c.fill = solid(IN);
  c.border = { top: { style: "thin", color: { argb: "FF8A8F98" } }, left: { style: "thin", color: { argb: "FF8A8F98" } }, bottom: { style: "thin", color: { argb: "FF8A8F98" } }, right: { style: "thin", color: { argb: "FF8A8F98" } } };
  c.protection = { locked: false };
  c.numFmt = "@"; // 文字として受ける（0で始まる品番もそのまま）
  c.font = { ...font, size: o.size || 14, bold: true };
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
};
const titleBar = (ws, title, lastCol, guideName) => {
  ws.mergeCells(1, 2, 1, lastCol - 1);
  const t = ws.getCell(1, 2);
  t.value = title;
  t.font = { ...font, size: 18, bold: true, color: { argb: DK } };
  t.alignment = { vertical: "middle", indent: 1 };
  for (let i = 1; i <= lastCol; i++) ws.getCell(1, i).fill = solid(AC);
  const g = ws.getCell(1, lastCol);
  g.value = { formula: `HYPERLINK("#"&"${quote(guideName)}!A1","使い方 →")` };
  g.font = { ...font, color: { argb: "FF1A4FA0" }, underline: true };
  g.alignment = { vertical: "middle", horizontal: "right" };
  ws.getRow(1).height = 32;
};

// ---- データシート（在庫リストと同じ列。丸ごと貼り付け用） ----
function fillData(ws) {
  ws.columns = STOCK_HEADS.map((h, i) => ({ width: i === 1 || i === 3 ? 16 : i === 4 ? 22 : 10 }));
  STOCK_HEADS.forEach((h, i) => {
    const c = ws.getCell(1, i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
  });
  const info = ws.getCell(1, STOCK_HEADS.length + 3);
  info.value = `← 在庫リストのExcelを、A1から、見出しごと、そのまま丸ごとコピーして貼り付けます（${DATA_MAX.toLocaleString()}行目まで見ます）`;
  info.font = { ...font, size: 9, bold: true, color: { argb: "FFD93025" } };
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// 全角／半角の違い（ハイフン・スペースなど）を吸収して、見つかりやすくする
const norm = (addr) => `SUBSTITUTE(SUBSTITUTE(UPPER(TRIM(ASC(${addr}))),"-","")," ","")`;

// ---- 検索用シート（隠し。部品番号を、探しやすく整えた列だけを持つ） ----
// データシートとは別のシートに置くことで、在庫リストの貼り付けが、どれだけ広い範囲でも
// 巻き込まれないようにしている
function fillCalc(ws, dataName) {
  ws.getCell(`${NORM_COL}1`).value = "検索用（消さない・触らない）";
  for (let r = 2; r <= DATA_MAX + 1; r++)
    ws.getCell(`${NORM_COL}${r}`).value = { formula: `IF(${quote(dataName)}!$${PN_COL}${r}="","",${norm(`${quote(dataName)}!$${PN_COL}${r}`)})` };
}

// 入力セル addr の中身から、見つかった行番号を返す数式（正規化して、部品番号列と突き合わせ）
const rowIdxFormula = (calcName, addr) => `IFERROR(MATCH(${norm(addr)},${quote(calcName)}!$${NORM_COL}$2:$${NORM_COL}$${DATA_MAX + 1},0),"")`;
const pnFromIdx = (dataName, idxAddr, inputAddr) => `IF(${inputAddr}="","",IF(${idxAddr}="","見つかりません",INDEX(${quote(dataName)}!$${PN_COL}$2:$${PN_COL}$${DATA_MAX + 1},${idxAddr})&""))`;
const locFromIdx = (dataName, idxAddr, inputAddr) => `IF(${inputAddr}="","",IF(${idxAddr}="","見つかりません",INDEX(${quote(dataName)}!$${LOC_COL}$2:$${LOC_COL}$${DATA_MAX + 1},${idxAddr})&""))`;

// ---- ラベル作成（1件ずつ） ----
function fillFind(ws, nm) {
  setLabelCols(ws);
  titleBar(ws, "ラベル作成（1件ずつ・すぐ印刷）", 14, nm.guide);
  label2(ws, "B3", "部品番号", { font: { size: 12 } });
  ws.mergeCells(3, 3, 3, 7);
  yellow(ws, "C3", { size: 16 });
  ws.getRow(3).height = 32;
  ws.getCell("B4").value = { formula: `IF($C$3="","← 部品番号を入れてください",IF($P$3="","見つかりません（データ・品番を確認してください）","見つかりました："&$Q$3&"　／　"&$R$3))` };
  ws.getCell("B4").font = { ...font, size: 12, bold: true, color: { argb: "FFB06A00" } };
  ws.mergeCells(4, 2, 4, 10);
  ws.getRow(4).height = 22;
  // P：見つかった行番号（隠し）／Q：部品番号／R：ロケーション
  ws.getCell("P3").value = { formula: rowIdxFormula(nm.calc, "$C$3") };
  ws.getCell("Q3").value = { formula: pnFromIdx(nm.data, "$P$3", "$C$3") };
  ws.getCell("R3").value = { formula: locFromIdx(nm.data, "$P$3", "$C$3") };
  for (const c of ["P", "Q", "R"]) ws.getColumn(c.charCodeAt(0) - 64).hidden = true;
  ws.getRow(5).height = 10;
  const GRID_TOP = 6;
  paintPage(ws, GRID_TOP, "$Q$3", "$R$3");
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 }, printArea: `B$${GRID_TOP}:L$${GRID_TOP + 48}` };
  return ws;
}

// まとめて印刷の、入力欄の並び（ラベルの列幅をそのまま流用するため、狭い仕切り列（D・G）は使わない）
const BATCH_COLS = { no: 2, input: 3, pn: 5, loc: 6, status: 8 }; // B・C・E・F・H
const thinBorder = { style: "thin", color: { argb: "FFB7BBC2" } };
const box = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };

// ---- まとめて印刷 ----
function fillBatch(ws, nm) {
  setLabelCols(ws);
  titleBar(ws, `まとめて印刷（1件＝1ページ・最大${BATCH_MAX}件）`, 14, nm.guide);
  const HR = 3;
  [
    [BATCH_COLS.no, "No"],
    [BATCH_COLS.input, "部品番号"],
    [BATCH_COLS.pn, "部品番号"],
    [BATCH_COLS.loc, "ロケーション"],
    [BATCH_COLS.status, "印刷ページ"],
  ].forEach(([col, h]) => {
    const c = ws.getCell(HR, col);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle" };
  });
  ws.getRow(HR).height = 22;
  const LIST_TOP = HR + 1;
  for (let i = 0; i < BATCH_MAX; i++) {
    const r = LIST_TOP + i;
    const inp = `$C$${r}`;
    const no = ws.getCell(r, BATCH_COLS.no);
    no.value = i + 1;
    no.alignment = { horizontal: "center" };
    no.font = { ...font, color: { argb: "FF8A8F98" } };
    no.border = box;
    yellow(ws, `C${r}`, { size: 12 });
    ws.getCell(`P${r}`).value = { formula: rowIdxFormula(nm.calc, inp) };
    const pn = ws.getCell(r, BATCH_COLS.pn);
    pn.value = { formula: pnFromIdx(nm.data, `$P$${r}`, inp) };
    pn.font = font;
    pn.border = box;
    pn.alignment = { vertical: "middle" };
    const loc = ws.getCell(r, BATCH_COLS.loc);
    loc.value = { formula: locFromIdx(nm.data, `$P$${r}`, inp) };
    loc.font = font;
    loc.border = box;
    loc.alignment = { vertical: "middle" };
    const st = ws.getCell(r, BATCH_COLS.status);
    st.value = { formula: `IF(${inp}="","",IF($${colLetter(BATCH_COLS.pn)}$${r}="見つかりません","見つかりません",${i + 1}))` };
    st.font = { ...font, color: { argb: "FFB06A00" }, bold: true };
    st.alignment = { horizontal: "center", vertical: "middle" };
    st.border = box;
    ws.getRow(r).height = 18;
  }
  ws.getColumn(16).hidden = true; // P列
  ws.mergeCells(2, 2, 2, 11);
  ws.getCell("B2").value = { formula: `IF(COUNTIF($C$${LIST_TOP}:$C$${LIST_TOP + BATCH_MAX - 1},"<>")=0,"部品番号を入れると、下にページができます","入力 "&COUNTIF($C$${LIST_TOP}:$C$${LIST_TOP + BATCH_MAX - 1},"<>")&"件　→　印刷は、ページ指定で「1-"&COUNTIF($C$${LIST_TOP}:$C$${LIST_TOP + BATCH_MAX - 1},"<>")&"」")` };
  ws.getCell("B2").font = { ...font, size: 11, bold: true, color: { argb: "FFB06A00" } };
  ws.getRow(LIST_TOP + BATCH_MAX).height = 10;
  const GRID_TOP = LIST_TOP + BATCH_MAX + 1;
  for (let i = 0; i < BATCH_MAX; i++) paintPage(ws, GRID_TOP + i * 49, `$${colLetter(BATCH_COLS.pn)}$${LIST_TOP + i}`, `$${colLetter(BATCH_COLS.loc)}$${LIST_TOP + i}`);
  ws.views = [{ showGridLines: false, state: "frozen", ySplit: HR }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 }, printArea: `B$${GRID_TOP}:L$${GRID_TOP + 49 * BATCH_MAX - 1}` };
  return ws;
}

// opt: { protect, names?: {find,batch,data,guide,calc}, sample?:boolean, rows?: 在庫リストの行（配列の配列。STOCK_HEADSの並び） }
export function buildLabelWorkbook(ExcelJS, opt = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = {
    find: opt.names?.find || "ラベル作成",
    batch: opt.names?.batch || "まとめて印刷",
    data: opt.names?.data || "ラベルデータ",
    guide: opt.names?.guide || "ラベルの使い方",
    calc: opt.names?.calc || "検索用（隠し）",
  };
  const protect = !!opt.protect;
  const wsFind = wb.addWorksheet(nm.find, { views: [{ showGridLines: false, tabSelected: true }] });
  const wsBatch = wb.addWorksheet(nm.batch, { views: [{ showGridLines: false }] });
  const wsData = wb.addWorksheet(nm.data, { views: [{ state: "frozen", ySplit: 1 }] });
  const wsGuide = wb.addWorksheet(nm.guide, { views: [{ showGridLines: false }] });
  const wsCalc = wb.addWorksheet(nm.calc, { state: "veryHidden" });
  const list = [nm.find, nm.batch, nm.data, nm.guide, nm.calc];

  fillData(wsData);
  const rows = opt.rows || (opt.sample ? sampleRows() : []);
  rows.forEach((row, i) => row.forEach((v, j) => v !== null && v !== undefined && (wsData.getCell(i + 2, j + 1).value = v)));
  fillCalc(wsCalc, nm.data);

  fillFind(wsFind, nm);
  fillBatch(wsBatch, nm);
  if (opt.sample) wsFind.getCell("C3").value = rows[0]?.[3] || "";

  fillGuide(wsGuide, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  fillGuideStyle(wsGuide, { title: "ラベル作成の使い方" });

  if (protect) {
    wsFind.sheetProtection = { sheet: true, formatColumns: true, formatRows: true };
    wsBatch.sheetProtection = { sheet: true, formatColumns: true, formatRows: true };
    wsCalc.sheetProtection = { sheet: true };
  }
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list, names: nm };
}

function sampleRows() {
  const pn = ["AY100-NS004", "KA230-00891", "24319-C9910", "AY080-00090", "081B4-2405M"];
  const loc = ["1A02-15C", "1A02-11A", "1A01-16P", "1A01-25D", "1A01-14G"];
  return pn.map((p, i) => {
    const row = Array(STOCK_HEADS.length).fill("");
    row[1] = loc[i];
    row[3] = p;
    row[4] = "見本データ";
    return row;
  });
}
