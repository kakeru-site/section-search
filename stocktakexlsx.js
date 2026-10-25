// 出庫分の循環棚卸を、Excelの数式だけで回せるブックにして書き出す（オイル・ケミカルのExcelと同じデザイン）
// マクロは使わない（Excel 2016 までの関数だけ）。入力は黄色いセルだけ。
//  ① 棚卸検索：部品番号を貼ると、ロケーション・実在庫・N+1が出る（重複・未登録は印が付く）
//  ② 場所順リスト：①を、ロケーション順に自動で並べた、印刷用の数え表
//  ③ ランダム抽出：DBから、割合（例：20%）だけ、ランダムに抜き出し、場所順に並べる
//  ④ DB：最新の在庫リストを貼る場所
//  ⑤ 誤差記録と目標：誤差を記録し、誤差率が目標以内かを見る
import { fillGuide } from "./oilxlsx.js?v=20261021u";

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const hair = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const font = { name: "Yu Gothic UI", size: 11 };
const quote = (n) => `'${n.replace(/'/g, "''")}'`;

export const DB_MAX = 3000; // DBシートで数式が見る行数（2行目〜）
const SR = 1000; // ①棚卸検索の行数
const ER = 600; // ③ランダム抽出で並べられる行数
const GR = 200; // ⑤誤差記録の行数
const STRONG = 20; // 最初から枠を整えてある行数

const DB_HEADS = ["拠点コード", "ロケーション", "常備", "部品番号", "部品名称", "取得", "商区", "メーカーコード", "実在庫数", "委託数", "移管中数", "自送N", "自送N+1", "MAX", "O/P", "MIN", "拠点3D", "拠点6D", "原価", "販売単位"];

const guideBlocks = (protect, nm) => [
  ["title", "循環棚卸ツールの使い方"],
  ["sub", "出庫分の循環棚卸を、Excelの数式だけで回すブックです（マクロなし）。在庫リストをDBに貼って、数える部品番号を貼るだけ。黄色いところに入れてください。"],
  ["h", "いつもの流れ"],
  ["step", "①", `「${nm.db}」シートに、最新の在庫リストを貼ります（A列〜T列。見出しは1行目）`],
  ["step", "②", `「${nm.search}」シートの部品番号の列に、数える部品番号を貼ります。ロケーション・実在庫数・N+1が出ます`],
  ["step", "③", `「${nm.sorted}」シートを印刷します。ロケーション順に並んだ、数え表です（重複・未登録は入りません）`],
  ["step", "④", `数えた数を、「${nm.search}」シートの「現在個数」に入れます。差数と判定が出ます`],
  ["step", "⑤", `差が出た部品は、「${nm.err}」シートに記録します。誤差率が、目標以内かが出ます`],
  ["h", `「${nm.search}」シート`],
  ["b", "部品番号の列（黄色）に、貼り付けるだけ。ロケーションは、DBから自動で出ます。DBにない番号は「未登録」と出ます"],
  ["b", "同じ部品番号を2回入れると、2つ目に「重複」と出ます（元のマクロのように消えないので、見てから消してください）"],
  ["b", "判定：OK＝差なし／差あり＝実在庫と違う／未入力＝まだ数えていない／未登録・重複"],
  ["b", "いちばん上の帯に、件数・未登録・重複・差ありの数が出ます"],
  ["h", `「${nm.sorted}」シート`],
  ["b", `「${nm.search}」の部品を、ロケーション順に、自動で並べ替えます（未登録・重複は除きます）`],
  ["b", "「数えた数」の欄は、紙に書くために空けてあります。A4たてに、1ページに収まります"],
  ["b", "印刷は、「印刷範囲を選択」か、ページを指定して印刷してください（空の行まで印刷されるのを避けるため）"],
  ["h", `「${nm.rand}」シート`],
  ["b", "DBの中から、割合（初期値20%）の件数を、ランダムに抜き出し、ロケーション順に並べます"],
  ["b", "「ばらつきの種」の数字を変えると、別の抽出になります。同じ数字なら、何度開いても、同じ抽出です（勝手に変わりません）"],
  ["b", "「数えた数」を入れると、差数が出ます"],
  ["h", `「${nm.db}」シート`],
  ["b", "最新の在庫リストを、そのまま貼り付けます（見出しの並びは、元のDBと同じ）。部品番号・ロケーション・実在庫数・自送N+1を使います"],
  ["b", `最初に入っている行は、見本です。使う前に、A列〜T列の見本を消して、本物を貼ってください（右の灰色の列は数式なので、消さないでください）`],
  ["b", `数式は、DBの${DB_MAX}行目まで見ます`],
  ["h", `「${nm.err}」シート`],
  ["b", "ロケーションと名称は、部品番号を入れると、DBから自動で出ます"],
  ["b", "総部品点数は、空のままなら、DBの件数になります。目標の誤差率（初期値0.5%）も、変えられます"],
  ["b", "誤差率が、目標を超えると、赤く出ます"],
  ["h", "困ったとき"],
  ["b", "何も出ない・計算されない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", "ロケーションが「未登録」になる → DBに、その部品番号があるか、全角・半角、ハイフンが同じかを見てください"],
  ["b", "並び順がおかしい → ロケーションの文字（全角・半角、ハイフン）が揃っているか見てください"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "在庫数・原価を含む表は、社外に出さないでください"],
  ...(protect
    ? [
        ["b", "数式を間違って消さないように、シートは保護してあります。入力できるのは黄色い欄だけです（DBは保護していません）"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

const titleBand = (ws, text, lastCol, guideName, sub) => {
  ws.mergeCells(1, 2, 1, lastCol - 2);
  const c = ws.getCell(1, 2);
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
  if (sub) {
    ws.mergeCells(2, 2, 2, lastCol);
    ws.getRow(2).height = 18;
    const s = ws.getCell(2, 2);
    s.value = sub;
    s.font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
    s.alignment = { vertical: "middle", indent: 1 };
  }
};
const label = (ws, addr, text, o = {}) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, bold: true, ...o.font };
  c.alignment = { vertical: "middle", shrinkToFit: true, ...o.alignment };
};
const inputCell = (ws, addr, value, o = {}) => {
  const c = ws.getCell(addr);
  c.value = value;
  c.fill = solid(IN);
  c.border = box;
  c.font = { ...font, size: 13, bold: true };
  c.alignment = { vertical: "middle", horizontal: o.left ? "left" : "right" };
  if (o.fmt) c.numFmt = o.fmt;
  c.protection = { locked: false };
  return c;
};
const note = (ws, addr, text) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
  c.alignment = { vertical: "middle" };
};
const header = (ws, row, c0, cols) => {
  cols.forEach((h, i) => {
    const c = ws.getRow(row).getCell(c0 + i);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(row).height = 32;
};
const sumBand = (ws, row, c0, items) => {
  items.forEach(([h, f, result, fmt], i) => {
    const hc = ws.getCell(row, c0 + i);
    hc.value = h;
    hc.font = { ...font, size: 9, bold: true, color: { argb: "FF4A4F57" } };
    hc.fill = solid("FFEDEFF2");
    hc.alignment = { horizontal: "center", vertical: "middle" };
    hc.border = box;
    const c = ws.getCell(row + 1, c0 + i);
    c.value = { formula: f, result: result === 0 ? undefined : result };
    c.font = { ...font, size: 13, bold: true };
    c.alignment = { horizontal: "center", vertical: "middle" };
    c.border = box;
    if (fmt) c.numFmt = fmt;
  });
  ws.getRow(row + 1).height = 26;
};
const cfFill = (argb, font2) => ({ fill: { type: "pattern", pattern: "solid", bgColor: { argb } }, ...(font2 ? { font: font2 } : {}) });
const protectOpt = { sheet: true, formatColumns: true, formatRows: true, autoFilter: false, sort: false };

// DB の範囲（絶対参照）
const dbr = (col, db) => `${quote(db)}!$${col}$2:$${col}$${DB_MAX + 1}`;

// ================= ① 棚卸検索 =================
function fillSearch(ws, nm, protect, sample) {
  // B ロケーション C 部品番号(入力) D 部品名称 E 実在庫数 F N+1 G 現在個数(入力) H 差数 I 判定 ／ 隠し K 有効 L 順位
  ws.columns = [{ width: 2 }, { width: 14 }, { width: 18 }, { width: 28 }, { width: 11 }, { width: 9 }, { width: 11 }, { width: 9 }, { width: 11 }, { width: 3 }, { width: 6 }, { width: 8 }];
  titleBand(ws, "出庫分 循環棚卸　ロケーション検索", 9, nm.guide);
  const F = 8;
  const L = F + SR - 1;
  const rg = (c) => `$${c}$${F}:$${c}$${L}`;
  sumBand(ws, 3, 2, [
    ["件数", `COUNTIF(${rg("C")},"?*")`, 0, '0"件"'],
    ["未登録", `COUNTIF(${rg("I")},"未登録")`, 0, '0"件"'],
    ["重複", `COUNTIF(${rg("I")},"重複")`, 0, '0"件"'],
    ["数え済み", `COUNT(${rg("G")})`, 0, '0"件"'],
    ["差あり", `COUNTIF(${rg("I")},"差あり")`, 0, '0"件"'],
    ["差の合計", `IF(COUNT(${rg("H")})=0,"",SUMIF(${rg("H")},">0")-SUMIF(${rg("H")},"<0"))`, "", '#,##0'],
  ]);
  ws.getRow(5).height = 8;
  note(ws, "B6", "↓ 黄色い「部品番号」の列に、数える部品番号を貼ります。数え終わったら、「現在個数」に入れます");
  header(ws, 7, 2, ["ロケーション", "部品番号", "部品名称", "実在庫数", "N+1", "現在個数", "差数", "判定"]);
  ws.getCell("C7").fill = solid(AC);
  const d = nm.db;
  for (let r = F; r <= L; r++) {
    const strong = r < F + STRONG;
    const m = `MATCH($C${r},${dbr("D", d)},0)`;
    ws.getCell(`B${r}`).value = {
      formula: `IF($C${r}="","",IFERROR(IF(TRIM(INDEX(${dbr("B", d)},${m})&"")="","未登録",TRIM(INDEX(${dbr("B", d)},${m})&"")),"未登録"))`,
      result: "",
    };
    ws.getCell(`D${r}`).value = { formula: `IF($C${r}="","",IFERROR(INDEX(${dbr("E", d)},${m})&"",""))`, result: "" };
    ws.getCell(`E${r}`).value = { formula: `IF($C${r}="","",IFERROR(VLOOKUP($C${r},${quote(d)}!$D$2:$I$${DB_MAX + 1},6,FALSE),"未登録"))`, result: "" };
    ws.getCell(`F${r}`).value = { formula: `IF($C${r}="","",IFERROR(VLOOKUP($C${r},${quote(d)}!$D$2:$M$${DB_MAX + 1},10,FALSE),"未登録"))`, result: "" };
    ws.getCell(`H${r}`).value = { formula: `IF(OR($C${r}="",$G${r}=""),"",IF(ISNUMBER($E${r}),$G${r}-$E${r},""))`, result: "" };
    ws.getCell(`I${r}`).value = {
      formula: `IF($C${r}="","",IF(COUNTIF($C$${F}:$C${r},$C${r})>1,"重複",IF($B${r}="未登録","未登録",IF($G${r}="","未入力",IF($H${r}=0,"OK","差あり")))))`,
      result: "",
    };
    ws.getCell(`K${r}`).value = { formula: `IF(OR($C${r}="",$I${r}="重複",$I${r}="未登録"),0,1)` };
    ws.getCell(`L${r}`).value = {
      formula: `IF($K${r}=1,COUNTIFS(${rg("K")},1,${rg("B")},"<"&$B${r})+COUNTIFS($K$${F}:$K${r},1,$B$${F}:$B${r},$B${r}),"")`,
      result: "",
    };
    for (const col of ["B", "C", "D", "E", "F", "G", "H", "I"]) {
      const c = ws.getCell(`${col}${r}`);
      c.font = col === "C" ? { name: "Consolas", size: 12, bold: true } : font;
      c.border = strong ? box : hair;
      c.alignment = { vertical: "middle", horizontal: ["E", "F", "G", "H"].includes(col) ? "right" : col === "I" ? "center" : "left" };
      if ((col === "C" || col === "G") && strong) c.fill = solid(IN);
      if (col === "C" || col === "G") c.protection = { locked: false };
      if (["E", "F", "G", "H"].includes(col)) c.numFmt = "#,##0;[Red]-#,##0";
    }
    if (strong) ws.getRow(r).height = 22;
  }
  ws.getColumn(11).hidden = true;
  ws.getColumn(12).hidden = true;
  // 判定の色
  ws.addConditionalFormatting({
    ref: `I${F}:I${L}`,
    rules: [
      { type: "expression", formulae: [`$I${F}="OK"`], style: cfFill("FFE6F4EA", { bold: true, color: { argb: "FF188038" } }) },
      { type: "expression", formulae: [`$I${F}="差あり"`], style: cfFill("FFFAD4D0", { bold: true, color: { argb: "FFD93025" } }) },
      { type: "expression", formulae: [`OR($I${F}="未登録",$I${F}="重複")`], style: cfFill("FFFFE0B2", { bold: true, color: { argb: "FFB06A00" } }) },
      { type: "expression", formulae: [`$I${F}="未入力"`], style: cfFill("FFEDEFF2", { color: { argb: "FF7A7F87" } }) },
    ],
  });
  ws.addConditionalFormatting({
    ref: `B${F}:H${L}`,
    rules: [{ type: "expression", formulae: [`OR($I${F}="重複",$I${F}="未登録")`], style: { font: { color: { argb: "FFB06A00" } } } }],
  });
  if (sample) {
    sample.parts.forEach((p, i) => {
      ws.getCell(`C${F + i}`).value = p;
    });
    sample.counts.forEach(([i, n]) => (ws.getCell(`G${F + i}`).value = n));
  }
  ws.views = [{ state: "frozen", ySplit: 7, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "7:7" };
  if (protect) ws.sheetProtection = protectOpt;
  return { first: F, last: L };
}

// ================= ② 場所順リスト（印刷用） =================
function fillSorted(ws, nm, protect, S) {
  // B No C ロケーション D 部品番号 E 部品名称 F 実在庫数 G N+1 H 数えた数 I 差 J ✓ ／ 隠し L 行番号
  ws.columns = [{ width: 2 }, { width: 5 }, { width: 14 }, { width: 18 }, { width: 28 }, { width: 9 }, { width: 8 }, { width: 11 }, { width: 8 }, { width: 6 }, { width: 3 }, { width: 6 }];
  titleBand(ws, "場所順リスト（印刷して、数える）", 10, nm.guide);
  ws.getCell("B3").value = { formula: `"全 "&MAX(${quote(nm.search)}!$L$${S.first}:$L$${S.last})&" 件　（未登録・重複は除いています）"`, result: "" };
  ws.getCell("B3").font = { ...font, bold: true, color: { argb: "FFB06A00" } };
  ws.getRow(3).height = 22;
  ws.getRow(4).height = 6;
  const HR = 5;
  header(ws, HR, 2, ["No", "ロケーション", "部品番号", "部品名称", "実在庫", "N+1", "数えた数", "差", "確認"]);
  const F = HR + 1;
  const sq = quote(nm.search);
  for (let k = 1; k <= SR; k++) {
    const r = F + k - 1;
    ws.getCell(`L${r}`).value = { formula: `IFERROR(MATCH(${k},${sq}!$L$${S.first}:$L$${S.last},0),"")`, result: "" };
    ws.getCell(`B${r}`).value = { formula: `IF($L${r}="","",${k})`, result: "" };
    const pick = (col) => `IF($L${r}="","",INDEX(${sq}!$${col}$${S.first}:$${col}$${S.last},$L${r})&"")`;
    ws.getCell(`C${r}`).value = { formula: pick("B"), result: "" };
    ws.getCell(`D${r}`).value = { formula: pick("C"), result: "" };
    ws.getCell(`E${r}`).value = { formula: pick("D"), result: "" };
    ws.getCell(`F${r}`).value = { formula: `IF($L${r}="","",INDEX(${sq}!$E$${S.first}:$E$${S.last},$L${r}))`, result: "" };
    ws.getCell(`G${r}`).value = { formula: `IF($L${r}="","",INDEX(${sq}!$F$${S.first}:$F$${S.last},$L${r}))`, result: "" };
    for (const col of ["B", "C", "D", "E", "F", "G", "H", "I", "J"]) {
      const c = ws.getCell(`${col}${r}`);
      c.font = col === "D" ? { name: "Consolas", size: 12, bold: true } : col === "C" ? { ...font, size: 12, bold: true } : font;
      c.alignment = { vertical: "middle", horizontal: ["B", "J"].includes(col) ? "center" : ["F", "G", "H", "I"].includes(col) ? "right" : "left" };
      if (["F", "G"].includes(col)) c.numFmt = "#,##0";
    }
    ws.getRow(r).height = 24;
  }
  ws.getColumn(12).hidden = true;
  // 行がある所だけ、枠線をつける（空の行は、何も見えない）
  ws.addConditionalFormatting({
    ref: `B${F}:J${F + SR - 1}`,
    rules: [{ type: "expression", formulae: [`$L${F}<>""`], style: { border: { top: thin, left: thin, bottom: thin, right: thin } } }],
  });
  ws.views = [{ state: "frozen", ySplit: HR, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: `${HR}:${HR}`, horizontalCentered: true, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.55, header: 0.25, footer: 0.25 } };
  ws.headerFooter = { oddFooter: "&C&9&P ページ" };
  if (protect) ws.sheetProtection = { sheet: true };
}

// ================= ③ ランダム抽出 =================
function fillRandom(ws, nm, protect) {
  // B No C ロケーション D 部品番号 E 部品名称 F 実在庫 G N+1 H 数えた数(入力) I 差 ／ 隠し K DB行
  ws.columns = [{ width: 2 }, { width: 5 }, { width: 14 }, { width: 18 }, { width: 28 }, { width: 9 }, { width: 8 }, { width: 11 }, { width: 8 }, { width: 3 }, { width: 6 }];
  titleBand(ws, "ランダム抽出（DBから、割合だけ、ランダムに）", 9, nm.guide);
  label(ws, "B3", "抽出する割合（%）");
  ws.mergeCells("B3:D3");
  inputCell(ws, "E3", 20, { fmt: "0.##", left: true });
  label(ws, "B4", "ばらつきの種");
  ws.mergeCells("B4:D4");
  inputCell(ws, "E4", 1, { fmt: "0", left: true });
  note(ws, "F4", "← 数字を変えると、別の抽出になります");
  const cnt = `COUNTIF(${dbr("D", nm.db)},"?*")`;
  label(ws, "B5", "DBの件数");
  ws.mergeCells("B5:D5");
  ws.getCell("E5").value = { formula: cnt };
  ws.getCell("E5").numFmt = '0"件"';
  ws.getCell("E5").font = { ...font, bold: true };
  ws.getCell("E5").alignment = { horizontal: "left" };
  label(ws, "B6", "抽出する件数");
  ws.mergeCells("B6:D6");
  ws.getCell("E6").value = { formula: `IF($E$5=0,0,MIN(${ER},MAX(1,ROUNDDOWN($E$5*$E$3/100,0))))` };
  ws.getCell("E6").numFmt = '0"件"';
  ws.getCell("E6").font = { ...font, size: 13, bold: true, color: { argb: "FFB06A00" } };
  ws.getCell("E6").alignment = { horizontal: "left" };
  ws.getRow(7).height = 8;
  const HR = 8;
  header(ws, HR, 2, ["No", "ロケーション", "部品番号", "部品名称", "実在庫", "N+1", "数えた数", "差"]);
  const F = HR + 1;
  const d = nm.db;
  for (let k = 1; k <= ER; k++) {
    const r = F + k - 1;
    ws.getCell(`K${r}`).value = { formula: `IF(${k}>$E$6,"",IFERROR(MATCH(${k},${dbr("Z", d)},0),""))`, result: "" };
    ws.getCell(`B${r}`).value = { formula: `IF($K${r}="","",${k})`, result: "" };
    const pick = (col) => `IF($K${r}="","",INDEX(${dbr(col, d)},$K${r})&"")`;
    ws.getCell(`C${r}`).value = { formula: pick("B"), result: "" };
    ws.getCell(`D${r}`).value = { formula: pick("D"), result: "" };
    ws.getCell(`E${r}`).value = { formula: pick("E"), result: "" };
    ws.getCell(`F${r}`).value = { formula: `IF($K${r}="","",INDEX(${dbr("I", d)},$K${r}))`, result: "" };
    ws.getCell(`G${r}`).value = { formula: `IF($K${r}="","",INDEX(${dbr("M", d)},$K${r}))`, result: "" };
    ws.getCell(`I${r}`).value = { formula: `IF(OR($K${r}="",$H${r}=""),"",IF(ISNUMBER($F${r}),$H${r}-$F${r},""))`, result: "" };
    for (const col of ["B", "C", "D", "E", "F", "G", "H", "I"]) {
      const c = ws.getCell(`${col}${r}`);
      c.font = col === "D" ? { name: "Consolas", size: 12, bold: true } : font;
      c.alignment = { vertical: "middle", horizontal: col === "B" ? "center" : ["F", "G", "H", "I"].includes(col) ? "right" : "left" };
      if (["F", "G", "H", "I"].includes(col)) c.numFmt = "#,##0;[Red]-#,##0";
      if (col === "H") {
        c.protection = { locked: false };
        c.fill = solid(IN);
      }
    }
    ws.getRow(r).height = 24;
  }
  ws.getColumn(11).hidden = true;
  ws.addConditionalFormatting({
    ref: `B${F}:I${F + ER - 1}`,
    rules: [{ type: "expression", formulae: [`$K${F}<>""`], style: { border: { top: thin, left: thin, bottom: thin, right: thin } } }],
  });
  ws.addConditionalFormatting({
    ref: `I${F}:I${F + ER - 1}`,
    rules: [{ type: "expression", formulae: [`AND($I${F}<>"",$I${F}<>0)`], style: cfFill("FFFAD4D0", { bold: true, color: { argb: "FFD93025" } }) }],
  });
  // 空の行は、黄色い入力欄の色を出さない（行がある所だけ、黄色）
  ws.views = [{ state: "frozen", ySplit: HR, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: `${HR}:${HR}`, horizontalCentered: true };
  ws.headerFooter = { oddFooter: "&C&9&P ページ" };
  if (protect) ws.sheetProtection = { sheet: true };
}

// ================= ④ DB =================
function fillDb(ws, nm, sampleRows) {
  const widths = [10, 14, 6, 16, 26, 6, 6, 12, 10, 8, 10, 8, 10, 7, 7, 7, 9, 9, 9, 9];
  ws.columns = [...widths.map((width) => ({ width })), { width: 3 }, { width: 12 }, { width: 14 }, { width: 8 }, { width: 6 }, { width: 8 }];
  DB_HEADS.forEach((h, i) => {
    const c = ws.getCell(1, i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ["並び用の場所", "乱数", "乱数順", "抽出", "抽出の順"].forEach((h, i) => {
    const c = ws.getCell(1, 22 + i);
    c.value = h;
    c.font = { ...font, size: 9, bold: true, color: { argb: "FF7A7F87" } };
    c.fill = solid("FFEDEFF2");
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(1).height = 32;
  const rnd = quote(nm.rand);
  const N = DB_MAX + 1;
  for (let r = 2; r <= N; r++) {
    // V 並び用の場所／W 乱数（種が同じなら、いつも同じ）／X 乱数の順位／Y 抽出する(1)／Z 抽出の中の場所順
    ws.getCell(`V${r}`).value = { formula: `IF($D${r}="","",IF(TRIM($B${r}&"")="","~未設定",TRIM($B${r}&"")))`, result: "" };
    ws.getCell(`W${r}`).value = { formula: `IF($D${r}="","",ABS(MOD(SIN(ROW()*${rnd}!$E$4)*43758.5453,1))+ROW()*0.000000001)`, result: "" };
    ws.getCell(`X${r}`).value = { formula: `IF($W${r}="","",RANK($W${r},$W$2:$W$${N},1))`, result: "" };
    ws.getCell(`Y${r}`).value = { formula: `IF($X${r}="","",IF($X${r}<=${rnd}!$E$6,1,0))`, result: "" };
    ws.getCell(`Z${r}`).value = {
      formula: `IF($Y${r}=1,COUNTIFS($Y$2:$Y$${N},1,$V$2:$V$${N},"<"&$V${r})+COUNTIFS($Y$2:$Y${r},1,$V$2:$V${r},$V${r}),"")`,
      result: "",
    };
  }
  sampleRows.forEach((row, i) => {
    row.forEach((v, j) => {
      if (v !== null && v !== undefined) ws.getCell(i + 2, j + 1).value = v;
    });
  });
  const c1 = ws.getCell("B1");
  c1.note = "ここに、最新の在庫リストを、コピーして貼り付けます（A列〜T列、見出しは1行目）";
  for (let col = 22; col <= 26; col++) ws.getColumn(col).outlineLevel = 0;
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 20 } };
  ws.views = [{ state: "frozen", ySplit: 1, showGridLines: true }];
  ws.getCell("AB1").value = "← 見本の行は、本物を貼る前に、消してください（右の灰色の列は数式です）";
  ws.getCell("AB1").font = { ...font, size: 9, color: { argb: "FFD93025" }, bold: true };
}

// ================= ⑤ 誤差記録と目標 =================
function fillErr(ws, nm, protect, sample) {
  // B ロケーション C 部品番号(入力) D 部品名称 E 誤差数(入力) F 発見日(入力) G 調査結果(入力)
  ws.columns = [{ width: 2 }, { width: 14 }, { width: 18 }, { width: 28 }, { width: 11 }, { width: 13 }, { width: 44 }];
  titleBand(ws, "誤差記録と目標", 7, nm.guide);
  const F = 13;
  const L = F + GR - 1;
  const d = nm.db;
  ws.mergeCells("B3:C3");
  label(ws, "B3", "総部品点数（空ならDBの件数）");
  inputCell(ws, "D3", null, { fmt: "#,##0", left: true });
  note(ws, "E3", "← 棚卸の対象の、総点数");
  ws.mergeCells("B4:C4");
  label(ws, "B4", "目標の誤差率（%）");
  inputCell(ws, "D4", 0.5, { fmt: "0.##", left: true });
  const total = `IF(ISNUMBER($D$3),$D$3,COUNTIF(${dbr("D", d)},"?*"))`;
  const rg = (c) => `$${c}$${F}:$${c}$${L}`;
  sumBand(ws, 6, 2, [
    ["総部品点数", total, 0, "#,##0"],
    ["部品差異数", `COUNTIF(${rg("C")},"?*")`, 0, '0"件"'],
    ["誤差率", `IF(${total}=0,"",COUNTIF(${rg("C")},"?*")/${total})`, "", "0.00%"],
    ["目標以内の件数", `${total}*$D$4/100`, 0, '0.0"件"'],
    ["判定", `IF(${total}=0,"",IF(COUNTIF(${rg("C")},"?*")/${total}<=$D$4/100,"目標クリア","目標オーバー"))`, "", null],
    ["プラス／マイナス", `COUNTIF(${rg("E")},">0")&"件 ／ "&COUNTIF(${rg("E")},"<0")&"件"`, "", null],
  ]);
  ws.getRow(8).height = 8;
  note(ws, "B9", "一点あたりの誤差率（1件の差が、全体に与える割合）");
  ws.getCell("E9").value = { formula: `IF(${total}=0,"",1/${total})`, result: "" };
  ws.getCell("E9").numFmt = "0.000%";
  ws.getCell("E9").font = { ...font, bold: true };
  ws.getCell("E9").alignment = { horizontal: "left" };
  ws.getRow(10).height = 8;
  note(ws, "B11", "↓ 差が出た部品を、記録します（ロケーションと名称は、部品番号から自動で出ます）");
  header(ws, 12, 2, ["ロケーション", "部品番号", "部品名称", "誤差数", "発見日", "調査結果"]);
  for (let r = F; r <= L; r++) {
    const strong = r < F + STRONG;
    const m = `MATCH($C${r},${dbr("D", d)},0)`;
    ws.getCell(`B${r}`).value = { formula: `IF($C${r}="","",IFERROR(INDEX(${dbr("B", d)},${m})&"","未登録"))`, result: "" };
    ws.getCell(`D${r}`).value = { formula: `IF($C${r}="","",IFERROR(INDEX(${dbr("E", d)},${m})&"",""))`, result: "" };
    for (const col of ["B", "C", "D", "E", "F", "G"]) {
      const c = ws.getCell(`${col}${r}`);
      c.font = col === "C" ? { name: "Consolas", size: 12, bold: true } : font;
      c.border = strong ? box : hair;
      c.alignment = { vertical: "middle", horizontal: col === "E" ? "right" : "left", wrapText: col === "G" };
      if (["C", "E", "F", "G"].includes(col)) {
        c.protection = { locked: false };
        if (strong) c.fill = solid(IN);
      }
      if (col === "E") c.numFmt = "#,##0;[Red]-#,##0";
      if (col === "F") c.numFmt = "yyyy/m/d";
    }
    if (strong) ws.getRow(r).height = 22;
  }
  ws.addConditionalFormatting({
    ref: "F7",
    rules: [
      { type: "expression", formulae: [`$F$7="目標クリア"`], style: cfFill("FFE6F4EA", { bold: true, color: { argb: "FF188038" } }) },
      { type: "expression", formulae: [`$F$7="目標オーバー"`], style: cfFill("FFFAD4D0", { bold: true, color: { argb: "FFD93025" } }) },
    ],
  });
  ws.addConditionalFormatting({
    ref: `E${F}:E${L}`,
    rules: [
      { type: "expression", formulae: [`AND(ISNUMBER($E${F}),$E${F}>0)`], style: cfFill("FFFFF3C4", { bold: true }) },
      { type: "expression", formulae: [`AND(ISNUMBER($E${F}),$E${F}<0)`], style: cfFill("FFFAD4D0", { bold: true }) },
    ],
  });
  if (sample) sample.forEach(([i, part, n, day, memo]) => {
    ws.getCell(`C${F + i}`).value = part;
    ws.getCell(`E${F + i}`).value = n;
    ws.getCell(`F${F + i}`).value = day;
    ws.getCell(`G${F + i}`).value = memo;
  });
  ws.views = [{ state: "frozen", ySplit: 12, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "12:12" };
  if (protect) ws.sheetProtection = protectOpt;
}

// opt: { protect, sample:boolean, names? }
export function buildStocktakeWorkbook(ExcelJS, opt = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = {
    search: opt.names?.search || "棚卸検索",
    sorted: opt.names?.sorted || "場所順リスト",
    rand: opt.names?.rand || "ランダム抽出",
    db: opt.names?.db || "DB",
    err: opt.names?.err || "誤差記録と目標",
    guide: opt.names?.guide || "使い方",
  };
  const v = (name, first) => wb.addWorksheet(name, { views: [{ showGridLines: false, ...(first ? { tabSelected: true } : {}) }] });
  const wsS = v(nm.search, true);
  const wsO = v(nm.sorted);
  const wsR = v(nm.rand);
  const wsD = wb.addWorksheet(nm.db);
  const wsE = v(nm.err);
  const wsG = v(nm.guide);
  const protect = !!opt.protect;
  // 見本（使う前に消すもの）
  let sampleDb = [];
  let sampleSearch = null;
  let sampleErr = null;
  if (opt.sample) {
    const locs = ["1A01-01A", "1A01-02A", "1A02-01A", "1B01-01B", "1B01-02B", "1C03-05A", "2A01-01A", "2A01-03B", "2B02-02A", "2B02-04C", "3A01-01A", "3A02-02B", "3C01-04C", "4B01-12D", "4B01-90D", "4B02-33D", "5A01-01A", "5A01-02A", "5B03-01B", "5C02-02C"];
    for (let i = 0; i < 40; i++) {
      const row = new Array(20).fill(null);
      row[1] = locs[i % locs.length].replace(/A$/, String.fromCharCode(65 + (i % 3)));
      row[3] = `G0000-${String(1000 + i * 7)}`;
      row[4] = `見本の部品${String(i + 1).padStart(2, "0")}`;
      row[8] = 5 + ((i * 37) % 90);
      row[12] = 3 + ((i * 11) % 40);
      row[17] = 10 + ((i * 53) % 400);
      sampleDb.push(row);
    }
    sampleSearch = {
      parts: [sampleDb[3][3], sampleDb[17][3], sampleDb[8][3], sampleDb[17][3], "G9999-9999", sampleDb[25][3], sampleDb[1][3], sampleDb[30][3]],
      counts: [[0, sampleDb[3][8]], [1, sampleDb[17][8] + 2], [2, sampleDb[8][8]], [5, sampleDb[25][8] - 1]],
    };
    sampleErr = [[0, sampleDb[17][3], 2, new Date(2026, 9, 20), "見本：数え間違い"], [1, sampleDb[25][3], -1, new Date(2026, 9, 20), "見本：出庫の入力もれ"]];
  }
  const S = fillSearch(wsS, nm, protect, sampleSearch);
  fillSorted(wsO, nm, protect, S);
  fillRandom(wsR, nm, protect);
  fillDb(wsD, nm, sampleDb);
  fillErr(wsE, nm, protect, sampleErr);
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list: [nm.search, nm.sorted, nm.rand, nm.db, nm.err, nm.guide], names: nm };
}
