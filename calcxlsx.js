// 売価計算・原価計算・パッド／ローター を、Excelの計算表として書き出す（オイル・ケミカルのExcelと同じデザイン）
// 入力は黄色いセルだけ。計算はぜんぶ数式（Excel 2016 までの関数のみ）
import { fillGuide, freeName } from "./oilxlsx.js";

export const SALE_MODES = ["掛け率で出す", "仕切りに利益を上乗せ"];
export const ROUND_LABELS = ["1円↑", "1円↓", "5円↑", "5円↓", "10円↑", "10円↓"];
export const COST_MODES = ["税抜きのまま足す", "税込みを税抜きにして足す（スズキ）"];
const ROWS = 20; // 最初から見やすく整えている行数
const MAXR = 100; // 数式を入れてある行数
const FIRST = 10; // 表の最初の行
const LAST = FIRST + MAXR - 1;

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const hair = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const font = { name: "Yu Gothic UI", size: 11 };
const quote = (n) => `'${n.replace(/'/g, "''")}'`;
const YEN = '"¥"#,##0;[Red]-"¥"#,##0';

const guideBlocks = (protect, nm) => [
  ["title", "計算シートの使い方"],
  ["sub", "売価・原価・パッド／ローターの計算を、Excelでぱぱっとやるためのブックです。黄色いところに数字を入れるだけ。むずかしくないので、さくっと読んでみてください！"],
  ["h", "まずはこれだけ"],
  ["step", "①", "下のシート名（左下のタブ）から、使いたい計算を選びます"],
  ["step", "②", "黄色いところに数字を入れると、答えは自動で出ます（Enterで次の行へ）"],
  ["step", "③", "答えのところ（白い欄）は、数式が入っているので書き換えないでください"],
  ["b", "表は100行目まで使えます。最初の20行は枠をつけてあって、21行目から下も同じように入れられます"],
  ["b", "何も入れていない行は、空のままでOKです"],
  ["h", `「${nm.sale}」シート`],
  ["step", "設定", "上の黄色いところで、「売価の出し方」「掛け率（%）」「上乗せ（%）」「端数」を選びます"],
  ["b", "掛け率で出す：定価 × 掛け率 → 端数をそろえて売価に。定価だけ入れればOK（仕切りは入れなくても出ます）"],
  ["b", "仕切りに利益を上乗せ：仕切り ×（100＋上乗せ）% → 端数をそろえて売価に"],
  ["b", "端数：「5円↑」なら5円単位に切り上げ、「10円↓」なら10円単位に切り捨てです"],
  ["step", "入力", "「定価」と、「仕切り（円）」か「掛け」のどちらかを入れます。掛けは「75」でも「7.5」でも75掛として扱います。両方入れたら仕切りが優先です"],
  ["step", "答え", "仕入れ掛（何掛けで入ったか）・売価・定価の何掛け・粗利・粗利率が出ます。原価割れの行は赤くなります"],
  ["b", "いちばん上の合計のところに、売価計・粗利計・平均仕入れ掛（全体）が出ます"],
  ["h", `「${nm.cost}」シート`],
  ["step", "設定", "「金額の入れ方」を選びます。税抜きの金額を足すだけなら「税抜きのまま足す」、税込みの金額を入れるなら「税込みを税抜きにして足す（スズキ）」"],
  ["step", "入力", "「単価」と「数量」を入れます（数量が空なら1個として計算します）"],
  ["b", "税込みから税抜きにするときは、消費税 ＝ 税込み ÷ 11 を四捨五入、税抜き ＝ 税込み − 消費税 で計算します（スズキの伝票と同じやり方）"],
  ["b", "いちばん上に、合計（税抜き）・税込み計・消費税計が出ます"],
  ["h", `「${nm.pad}」シート`],
  ["step", "設定", "上の黄色いところに、普通車・軽自動車のセット価格、原価割れのときのローター利益、値引きが必要なときのパッド上乗せ（案1・案2）を入れておきます"],
  ["step", "入力", "1行に1台ぶん：パッド通常売価・パッド原価・ローター原価（1枚）を入れます。ローター定価（1枚）は、お得額を出したいときだけ入れます"],
  ["b", "ローター1枚 ＝（セット価格 − パッド通常売価）÷ 2（割り切れないときは1円下げてパッドに回します）"],
  ["b", "ローターが原価割れするときは、ローター ＝ 原価＋利益、パッド ＝ セット価格 − ローター×2 になります"],
  ["b", "それでもパッドが原価を下回るときは「値引きが必要」と出て、右の「提案」にパッド売価の案（原価＋案1／案2）と、「ローターはローター定価の○掛け＋パッドは通常売価」の案（ローター定価を入れたとき）が出ます"],
  ["h", "困ったとき"],
  ["b", "何も出ない・計算されない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", "答えが出ない行がある → 数字の途中に文字（円、カンマ以外）が入っていないか見てください"],
  ["b", "100行をこえて使いたい → 100行目の行をコピーして、その下に貼り付けてください（保護をかけている場合は先に解除）"],
  ["b", "おかしくなった → 元のファイルをもう一回コピーすれば、元どおりです"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "数字を入れたら、保存してもOKですが、共有するときは取引先の仕切り（原価）が入っていないか気をつけてください"],
  ...(protect
    ? [
        ["b", "計算の数式を間違って消さないように、シートは保護してあります。入力できるのは黄色い欄（とプルダウン）だけです"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

function titleBand(ws, text, lastCol, guideName) {
  ws.mergeCells(1, 1, 1, lastCol - 2);
  const c = ws.getCell(1, 1);
  c.value = text;
  c.font = { ...font, size: 18, bold: true, color: { argb: DK } };
  c.alignment = { vertical: "middle", indent: 1 };
  for (let i = 1; i <= lastCol; i++) ws.getCell(1, i).fill = solid(AC);
  ws.mergeCells(1, lastCol - 1, 1, lastCol);
  const g = ws.getCell(1, lastCol - 1);
  g.value = { formula: `HYPERLINK("#"&"${quote(guideName)}!A1","使い方はこちら →")`, result: "使い方はこちら →" };
  g.font = { ...font, color: { argb: "FF1A4FA0" }, underline: true };
  g.alignment = { vertical: "middle", horizontal: "right" };
  ws.getRow(1).height = 36;
  ws.getRow(2).height = 8;
}
const label = (ws, addr, text, o = {}) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, bold: true, ...o.font };
  c.alignment = { vertical: "middle", ...o.alignment };
};
const inputCell = (ws, addr, value, o = {}) => {
  const c = ws.getCell(addr);
  c.value = value;
  c.fill = solid(IN);
  c.border = box;
  c.font = { ...font, size: 13, bold: true };
  c.alignment = { vertical: "middle", horizontal: o.left ? "left" : "right" };
  if (o.fmt) c.numFmt = o.fmt;
  if (o.dv) c.dataValidation = o.dv;
  c.protection = { locked: false };
  return c;
};
const note = (ws, addr, text) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
  c.alignment = { vertical: "middle" };
};
const listDv = (items) => ({ type: "list", allowBlank: false, formulae: [`"${items.join(",")}"`] });
const header = (ws, row, cols) => {
  cols.forEach((h, i) => {
    const c = ws.getRow(row).getCell(i + 2);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(row).height = 34;
};
// 合計の帯（見出し1行＋数式1行）
const sumBand = (ws, row, items) => {
  items.forEach(([h, f, result, fmt], i) => {
    const hc = ws.getCell(row, i + 2);
    hc.value = h;
    hc.font = { ...font, size: 9, bold: true, color: { argb: "FF4A4F57" } };
    hc.fill = solid("FFEDEFF2");
    hc.alignment = { horizontal: "center", vertical: "middle" };
    hc.border = box;
    const c = ws.getCell(row + 1, i + 2);
    c.value = { formula: f, result };
    c.font = { ...font, size: 13, bold: true };
    c.alignment = { horizontal: "center", vertical: "middle" };
    c.border = box;
    if (fmt) c.numFmt = fmt;
  });
  ws.getRow(row + 1).height = 26;
};
// 表の1行ぶんの見た目（入力セットは黄色／下の予備行は白）
const rowStyle = (ws, r, inCols, nCols, o = {}) => {
  const strong = r < FIRST + ROWS;
  for (let i = 2; i <= nCols + 1; i++) {
    const c = ws.getCell(r, i);
    const isIn = inCols.includes(i);
    c.font = { ...font, ...(o.fonts?.[i] || {}) };
    c.border = strong ? box : hair;
    c.alignment = { vertical: "middle", horizontal: o.left?.includes(i) ? "left" : "right", wrapText: !!o.left?.includes(i) };
    if (isIn) {
      if (strong) c.fill = solid(IN);
      c.protection = { locked: false };
    }
    if (o.fmts?.[i]) c.numFmt = o.fmts[i];
  }
};
const redIf = (ws, ref, formula) =>
  ws.addConditionalFormatting({
    ref,
    rules: [{ type: "expression", formulae: [formula], style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFAD4D0" } } } }],
  });

// ================= 売価計算 =================
function fillSale(ws, nm, sale, protect) {
  // 列：B定価 C仕切り(入力) D掛け(入力) E仕切り F仕入れ掛 G売価 H定価の何掛け I粗利 J粗利率 ／ 隠し K〜O
  ws.columns = [
    { width: 2 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 13 }, { width: 14 }, { width: 15 }, { width: 13 }, { width: 11 },
    { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 },
  ];
  titleBand(ws, "売価計算", 10, nm.guide);
  const modeIdx = sale.mode === "mark" ? 1 : 0;
  const roundLabel = ROUND_LABELS[ROUND_LABELS.indexOf(sale.roundLabel)] || "5円↑";
  label(ws, "B3", "売価の出し方");
  ws.mergeCells("C3:E3");
  inputCell(ws, "C3", SALE_MODES[modeIdx], { left: true, dv: listDv(SALE_MODES) });
  for (const a of ["D3", "E3"]) ws.getCell(a).border = box;
  label(ws, "G3", "端数のしかた");
  inputCell(ws, "H3", roundLabel, { left: true, dv: listDv(ROUND_LABELS) });
  label(ws, "B4", "掛け率（%）");
  inputCell(ws, "C4", sale.rate ?? 85, { fmt: "0.##" });
  note(ws, "D4", "← 掛け率で出すとき");
  label(ws, "G4", "上乗せ（%）");
  inputCell(ws, "H4", sale.mark ?? 15, { fmt: "0.##" });
  note(ws, "I4", "← 上乗せのとき");
  // 計算用（隠し列）：端数の単位・向き／出し方／かける率
  const unit = parseInt(roundLabel, 10);
  const up = roundLabel.endsWith("↑") ? 1 : 0;
  const rateMode = modeIdx === 0 ? 1 : 0;
  const mult = Math.round((rateMode ? sale.rate ?? 85 : 100 + (sale.mark ?? 15)) * 100);
  ws.getCell("L3").value = { formula: `IF(OR($H$3="1円↑",$H$3="1円↓"),1,IF(OR($H$3="5円↑",$H$3="5円↓"),5,10))`, result: unit };
  ws.getCell("M3").value = { formula: `IF(RIGHT($H$3,1)="↑",1,0)`, result: up };
  ws.getCell("N3").value = { formula: `IF($C$3="${SALE_MODES[0]}",1,0)`, result: rateMode };
  ws.getCell("L5").value = {
    formula: `IF($N$3=1,ROUND($C$4*100,0),ROUND((100+$H$4)*100,0))`,
    result: mult,
  };
  const rng = (c) => `$${c}$${FIRST}:$${c}$${LAST}`;
  sumBand(ws, 6, [
    ["件数", `COUNT(${rng("G")})`, 0, '0"件"'],
    ["定価計", `SUM(${rng("N")})`, 0, YEN],
    ["仕切り計", `SUM(${rng("O")})`, 0, YEN],
    ["平均仕入れ掛", `IF(SUM(${rng("L")})>0,ROUND(SUM(${rng("M")})/SUM(${rng("L")})*1000,0)/10&"掛","")`, "", null],
    ["売価計", `SUM(${rng("G")})`, 0, YEN],
    ["粗利計", `SUM(${rng("I")})`, 0, YEN],
    ["粗利率", `IF(SUM(${rng("G")})>0,SUM(${rng("I")})/SUM(${rng("G")}),"")`, "", "0.0%"],
    ["原価割れ", `COUNTIF(${rng("I")},"<0")`, 0, '0"件"'],
  ]);
  ws.getRow(8).height = 8;
  header(ws, 9, ["定価", "仕切り（円）", "掛け", "仕切り", "仕入れ掛", "売価", "定価の何掛け", "粗利", "粗利率"]);
  ws.getCell("C9").fill = solid(AC);
  for (let r = FIRST; r <= LAST; r++) {
    rowStyle(ws, r, [2, 3, 4], 9, { fmts: { 2: "#,##0", 3: "#,##0", 5: YEN, 7: YEN, 9: YEN, 10: "0.0%" }, fonts: { 7: { bold: true } } });
    ws.getRow(r).height = 22;
    const E = `IF(ISNUMBER(C${r}),C${r},IF(AND(ISNUMBER(B${r}),ISNUMBER(D${r})),IF(D${r}>0,ROUND(B${r}*IF(D${r}<=10,D${r}*10,D${r})/100,0),""),""))`;
    ws.getCell(`E${r}`).value = { formula: E, result: "" };
    ws.getCell(`F${r}`).value = {
      formula: `IF(NOT(ISNUMBER($E${r})),"",IF(ISNUMBER(C${r}),IF(AND(ISNUMBER(B${r}),B${r}>0),ROUND(C${r}/B${r}*1000,0)/10&"掛",""),IF(D${r}<=10,ROUND(D${r}*100,0)/10,ROUND(D${r}*10,0)/10)&"掛"))`,
      result: "",
    };
    const R = `ROUND($K${r},0)*$L$5/(10000*$L$3)`;
    ws.getCell(`G${r}`).value = { formula: `IF(ISNUMBER($K${r}),IF($M$3=1,ROUNDUP(${R},0),ROUNDDOWN(${R},0))*$L$3,"")`, result: "" };
    ws.getCell(`H${r}`).value = {
      formula: `IF($N$3=1,"",IF(AND(ISNUMBER(G${r}),ISNUMBER(B${r})),IF(B${r}>0,ROUND(G${r}/B${r}*1000,0)/10&"掛",""),""))`,
      result: "",
    };
    ws.getCell(`I${r}`).value = { formula: `IF(AND(ISNUMBER(G${r}),ISNUMBER(E${r})),G${r}-E${r},"")`, result: "" };
    ws.getCell(`J${r}`).value = { formula: `IF(AND(ISNUMBER(I${r}),ISNUMBER(G${r})),IF(G${r}>0,I${r}/G${r},""),"")`, result: "" };
    ws.getCell(`F${r}`).alignment = { vertical: "middle", horizontal: "center" };
    ws.getCell(`H${r}`).alignment = { vertical: "middle", horizontal: "center" };
    // 計算用（隠し列）
    ws.getCell(`K${r}`).value = { formula: `IF($N$3=1,IF(ISNUMBER(B${r}),B${r},""),IF(ISNUMBER(E${r}),E${r},""))`, result: "" };
    ws.getCell(`L${r}`).value = { formula: `IF(AND(ISNUMBER(B${r}),ISNUMBER(E${r})),IF(B${r}>0,B${r},0),0)`, result: 0 };
    ws.getCell(`M${r}`).value = { formula: `IF(L${r}>0,E${r},0)`, result: 0 };
    ws.getCell(`N${r}`).value = { formula: `IF(ISNUMBER(G${r}),N(B${r}),0)`, result: 0 };
    ws.getCell(`O${r}`).value = { formula: `IF(ISNUMBER(G${r}),N(E${r}),0)`, result: 0 };
  }
  redIf(ws, `B${FIRST}:J${LAST}`, `AND(ISNUMBER($I${FIRST}),$I${FIRST}<0)`);
  for (const c of ["K", "L", "M", "N", "O"]) ws.getColumn(c).hidden = true;
  ws.views = [{ state: "frozen", ySplit: 9, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "9:9" };
  if (protect) ws.sheetProtection = { sheet: true };
}

// ================= 原価計算 =================
function fillCost(ws, nm, cost, protect) {
  // 列：B単価(入力) C数量(入力) D税抜き単価 E小計(税抜き) F消費税計(表示のみ) ／ 隠し G(数量) H(税込み小計)
  ws.columns = [{ width: 2 }, { width: 18 }, { width: 12 }, { width: 18 }, { width: 20 }, { width: 14 }, { width: 10 }, { width: 10 }];
  titleBand(ws, "原価計算", 5, nm.guide);
  const incl = cost.mode === "in" ? 1 : 0;
  label(ws, "B3", "金額の入れ方");
  ws.mergeCells("C3:E3");
  inputCell(ws, "C3", COST_MODES[incl], { left: true, dv: listDv(COST_MODES) });
  for (const a of ["D3", "E3"]) ws.getCell(a).border = box;
  ws.getCell("G3").value = { formula: `IF($C$3="${COST_MODES[1]}",1,0)`, result: incl };
  const rng = (c) => `$${c}$${FIRST}:$${c}$${LAST}`;
  sumBand(ws, 6, [
    ["行数", `COUNT(${rng("E")})`, 0, '0"行"'],
    ["点数", `SUM(${rng("G")})`, 0, '0"点"'],
    ["合計（税抜き）", `SUM(${rng("E")})`, 0, YEN],
    [incl ? "税込み計" : "税込み（×1.1）", `IF($G$3=1,SUM(${rng("H")}),INT(SUM(${rng("E")})*11/10))`, 0, YEN],
  ]);
  ws.getCell("E6").value = { formula: `IF($G$3=1,"税込み計","税込み（×1.1）")`, result: incl ? "税込み計" : "税込み（×1.1）" };
  // 消費税計（税込みで入れたときだけ）
  const hc = ws.getCell("F6");
  hc.value = { formula: `IF($G$3=1,"消費税計","")`, result: incl ? "消費税計" : "" };
  hc.font = { ...font, size: 9, bold: true, color: { argb: "FF4A4F57" } };
  hc.alignment = { horizontal: "center", vertical: "middle" };
  const vc = ws.getCell("F7");
  vc.value = { formula: `IF($G$3=1,SUM(${rng("H")})-SUM(${rng("E")}),"")`, result: incl ? 0 : "" };
  vc.font = { ...font, size: 13, bold: true };
  vc.alignment = { horizontal: "center", vertical: "middle" };
  vc.numFmt = YEN;
  ws.getRow(8).height = 8;
  header(ws, 9, ["単価", "数量", "税抜き単価", "小計（税抜き）"]);
  for (let r = FIRST; r <= LAST; r++) {
    rowStyle(ws, r, [2, 3], 5, { fmts: { 2: "#,##0", 3: "#,##0", 4: YEN, 5: YEN }, fonts: { 5: { bold: true } } });
    ws.getRow(r).height = 22;
    ws.getCell(`D${r}`).value = {
      formula: `IF(ISNUMBER(B${r}),IF($G$3=1,ROUND(B${r},0)-ROUND(ROUND(B${r},0)/11,0),ROUND(B${r},0)),"")`,
      result: "",
    };
    ws.getCell(`E${r}`).value = { formula: `IF(ISNUMBER(B${r}),D${r}*IF(ISNUMBER(C${r}),C${r},1),"")`, result: "" };
    // 計算用（隠し列）：数量・税込み小計
    ws.getCell(`G${r}`).value = { formula: `IF(ISNUMBER(B${r}),IF(ISNUMBER(C${r}),C${r},1),0)`, result: 0 };
    ws.getCell(`H${r}`).value = { formula: `IF(ISNUMBER(B${r}),ROUND(B${r},0)*G${r},0)`, result: 0 };
  }
  // G（数量）・H（税込み小計）は計算用なので隠す。F列は消費税計の表示用
  ws.getColumn(7).hidden = true;
  ws.getColumn(8).hidden = true;
  ws.views = [{ state: "frozen", ySplit: 9, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "9:9" };
  if (protect) ws.sheetProtection = { sheet: true };
}

// ================= パッド・ローター =================
function fillPad(ws, nm, pad, protect) {
  // 列：B車種メモ C〜F入力 ／ 普通車 G〜K ／ 軽 L〜P ／ 提案 Q〜S
  ws.columns = [{ width: 2 }, { width: 23 }, ...Array(4).fill({ width: 13 }), ...Array(10).fill({ width: 12.5 }), ...Array(3).fill({ width: 13 }), { width: 16 }, { width: 15 }];
  titleBand(ws, "パッド・ローター", 21, nm.guide);
  const s = pad;
  label(ws, "B3", "普通車のセット価格");
  inputCell(ws, "C3", s.normal, { fmt: YEN });
  label(ws, "B4", "軽自動車のセット価格");
  inputCell(ws, "C4", s.kei, { fmt: YEN });
  label(ws, "E3", "原価割れのとき ローター利益（円）", { alignment: { horizontal: "left" } });
  inputCell(ws, "I3", s.margin, { fmt: "#,##0" });
  label(ws, "E4", "値引き提案①（パッド原価＋）", { alignment: { horizontal: "left" } });
  inputCell(ws, "I4", s.plus1, { fmt: "#,##0" });
  label(ws, "E5", "値引き提案②（パッド原価＋）", { alignment: { horizontal: "left" } });
  inputCell(ws, "I5", s.plus2, { fmt: "#,##0" });
  label(ws, "E6", "値引き提案③（ローター定価の何掛け）", { alignment: { horizontal: "left" } });
  inputCell(ws, "I6", s.kake ?? 45, { fmt: "0.##" });
  note(ws, "B5", "※ ローター定価は任意（お得額用）");
  // 見出し（グループ）
  ws.getRow(7).height = 8;
  const grp = (c1, c2, text, argb) => {
    ws.mergeCells(`${c1}8:${c2}8`);
    const c = ws.getCell(`${c1}8`);
    c.value = text;
    c.font = { ...font, bold: true, color: { argb: argb === DK ? "FFFFFFFF" : DK } };
    c.alignment = { horizontal: "center", vertical: "middle" };
    for (const col of colRange(c1, c2)) {
      ws.getCell(`${col}8`).fill = solid(argb);
      ws.getCell(`${col}8`).border = box;
    }
  };
  const colRange = (a, b) => {
    const out = [];
    for (let i = a.charCodeAt(0); i <= b.charCodeAt(0); i++) out.push(String.fromCharCode(i));
    return out;
  };
  grp("C", "F", "入力", AC);
  grp("G", "K", "普通車セット", DK);
  grp("L", "P", "軽自動車セット", DK);
  grp("Q", "U", "値引きが必要なとき", AC);
  ws.getCell("G8").value = { formula: `"普通車セット　"&TEXT($C$3,"#,##0")&"円"`, result: `普通車セット　${s.normal.toLocaleString("en-US")}円` };
  ws.getCell("L8").value = { formula: `"軽自動車セット　"&TEXT($C$4,"#,##0")&"円"`, result: `軽自動車セット　${s.kei.toLocaleString("en-US")}円` };
  ws.getRow(8).height = 24;
  const head = ["車種・メモ（任意）", "パッド通常売価", "パッド原価", "ローター原価（1枚）", "ローター定価（任意）"];
  const grpHead = ["ローター 1枚", "パッド", "セット利益", "判定", "通常より お得"];
  header(ws, 9, [...head, ...grpHead, ...grpHead, "ローター 1枚", "パッド 提案①", "パッド 提案②", "ローター 提案③（定価の掛け）", "セット合計 提案③"]);
  ws.getRow(9).height = 50;
  const fmts = {};
  for (const i of [3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]) fmts[i] = i >= 7 ? YEN : "#,##0";
  fmts[3] = fmts[4] = fmts[5] = fmts[6] = "#,##0";
  for (let r = FIRST; r <= LAST; r++) {
    rowStyle(ws, r, [2, 3, 4, 5, 6], 20, { left: [2], fmts, fonts: { 7: { bold: true }, 12: { bold: true } } });
    ws.getRow(r).height = 22;
    const valid = `AND(ISNUMBER($C${r}),ISNUMBER($D${r}),ISNUMBER($E${r}))`;
    const sets = [
      ["G", "H", "I", "J", "K", "$C$3"],
      ["L", "M", "N", "O", "P", "$C$4"],
    ];
    for (const [cr, cp, cg, cj, ck, SET] of sets) {
      const half = `INT((${SET}-$C${r})/2)`;
      ws.getCell(`${cr}${r}`).value = { formula: `IF(${valid},IF(${half}>$E${r},${half},$E${r}+$I$3),"")`, result: "" };
      ws.getCell(`${cp}${r}`).value = { formula: `IF(${valid},${SET}-2*${cr}${r},"")`, result: "" };
      ws.getCell(`${cg}${r}`).value = { formula: `IF(${valid},(${cp}${r}-$D${r})+(${cr}${r}-$E${r})*2,"")`, result: "" };
      ws.getCell(`${cj}${r}`).value = {
        formula: `IF(${valid},IF(${half}>$E${r},"そのままでOK",IF(${cp}${r}<$D${r},"値引きが必要","パッドで調整")),"")`,
        result: "",
      };
      ws.getCell(`${cj}${r}`).alignment = { vertical: "middle", horizontal: "center" };
      ws.getCell(`${ck}${r}`).value = {
        formula: `IF(AND(${valid},ISNUMBER($F${r})),IF($C${r}+$F${r}*2>${SET},$C${r}+$F${r}*2-${SET},""),"")`,
        result: "",
      };
    }
    const need = `OR($J${r}="値引きが必要",$O${r}="値引きが必要")`;
    ws.getCell(`Q${r}`).value = { formula: `IF(${need},$E${r}+$I$3,"")`, result: "" };
    ws.getCell(`R${r}`).value = { formula: `IF(${need},$D${r}+$I$4,"")`, result: "" };
    ws.getCell(`S${r}`).value = { formula: `IF(${need},$D${r}+$I$5,"")`, result: "" };
    // 提案③：ローター＝ローター定価×掛け率（10円単位に切り上げ）、パッド＝通常売価。ローター定価を入れたときだけ
    ws.getCell(`T${r}`).value = {
      formula: `IF(AND(${need},ISNUMBER($F${r})),ROUNDUP(ROUND($F${r},0)*ROUND($I$6*100,0)/(10000*10),0)*10,"")`,
      result: "",
    };
    ws.getCell(`U${r}`).value = { formula: `IF(ISNUMBER($T${r}),$C${r}+$T${r}*2,"")`, result: "" };
    for (const c of ["Q", "R", "S", "T", "U"]) {
      ws.getCell(`${c}${r}`).numFmt = YEN;
      ws.getCell(`${c}${r}`).border = r < FIRST + ROWS ? box : hair;
      ws.getCell(`${c}${r}`).font = font;
      ws.getCell(`${c}${r}`).alignment = { vertical: "middle", horizontal: "right" };
    }
  }
  for (const col of ["J", "O"]) {
    ws.addConditionalFormatting({
      ref: `${col}${FIRST}:${col}${LAST}`,
      rules: [
        { type: "expression", formulae: [`${col}${FIRST}="値引きが必要"`], style: { font: { bold: true, color: { argb: "FFD93025" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFAD4D0" } } } },
        { type: "expression", formulae: [`${col}${FIRST}="パッドで調整"`], style: { font: { bold: true, color: { argb: "FFB06A00" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFF3C4" } } } },
        { type: "expression", formulae: [`${col}${FIRST}="そのままでOK"`], style: { font: { bold: true, color: { argb: "FF188038" } } } },
      ],
    });
  }
  ws.views = [{ state: "frozen", ySplit: 9, xSplit: 2, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "8:9" };
  if (protect) ws.sheetProtection = { sheet: true };
}

// opt: { protect, names?:{sale,cost,pad,guide}, sale:{mode,rate,mark,roundLabel}, cost:{mode}, pad:{normal,kei,margin,plus1,plus2} }
export function buildCalcWorkbook(ExcelJS, opt = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = {
    sale: opt.names?.sale || "売価計算",
    cost: opt.names?.cost || "原価計算",
    pad: opt.names?.pad || "パッド・ローター",
    guide: opt.names?.guide || "使い方",
  };
  const wsS = wb.addWorksheet(nm.sale, { views: [{ showGridLines: false, tabSelected: true }] });
  const wsC = wb.addWorksheet(nm.cost, { views: [{ showGridLines: false }] });
  const wsP = wb.addWorksheet(nm.pad, { views: [{ showGridLines: false }] });
  const wsG = wb.addWorksheet(nm.guide, { views: [{ showGridLines: false }] });
  const protect = !!opt.protect;
  fillSale(wsS, nm, opt.sale || {}, protect);
  fillCost(wsC, nm, opt.cost || {}, protect);
  fillPad(wsP, nm, opt.pad || { normal: 12000, kei: 10000, margin: 10, plus1: 500, plus2: 1000 }, protect);
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list: [nm.sale, nm.cost, nm.pad, nm.guide], names: nm };
}
export { freeName };
