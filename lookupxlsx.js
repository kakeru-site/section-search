// ユーザー（得意先）検索と、部品名→部品番号の検索を、関数だけのExcelにして書き出す（オイル・ケミカルのExcelと同じデザイン）
// 入力は黄色いセルだけ。マクロなし・Excel 2016 までの関数だけ。
//  ・ユーザー検索：名前（漢字・ひらがな・カタカナ・会社名の一部・担当者名）か、電話番号で、ユーザーコードを出す
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
const STRONG = 20; // 最初から枠を整えてある行数

// ひらがな⇄カタカナ（SUBSTITUTEを30個ずつ、3つのセルに分けて重ねる。関数のネストの上限（64）を超えないため）
const HIRA = [];
for (let c = 0x3041; c <= 0x3096; c++) HIRA.push(String.fromCharCode(c));
const toKata = (ch) => String.fromCharCode(ch.charCodeAt(0) + 0x60);
const chain = (pairs, inner) => pairs.reduce((acc, [a, b]) => `SUBSTITUTE(${acc},"${a}","${b}")`, inner);
const chunks = (arr, n) => {
  const out = [];
  for (let i = 0; i < arr.length; i += Math.ceil(arr.length / n)) out.push(arr.slice(i, i + Math.ceil(arr.length / n)));
  return out;
};
const H2K = chunks(HIRA.map((h) => [h, toKata(h)]), 3);
const K2H = chunks(HIRA.map((h) => [toKata(h), h]), 3);

const guideBlocks = (protect, nm) => [
  ["title", "ユーザー検索・部品検索の使い方"],
  ["sub", "ユーザー（得意先）の名前や電話番号から「ユーザーコード」を、部品名や部品番号の一部から「部品番号」を、すぐに探すブックです。黄色いところに言葉を入れるだけ。"],
  ["h", "まずはこれだけ"],
  ["step", "①", `「${nm.user}」か「${nm.part}」のシートを開きます`],
  ["step", "②", "上の黄色いところに、探す言葉を入れます。入れたそばから、下に候補が出ます"],
  ["step", "③", "もっと絞りたいときは、「もう1つの言葉」にも入れます（両方に当てはまるものだけ出ます）"],
  ["h", `「${nm.user}」シート`],
  ["b", "名前（漢字・ひらがな・カタカナ）、会社名の一部、担当者名、ユーザーコード、電話番号（ハイフンなしでもOK）で探せます"],
  ["b", "半角・全角、大文字・小文字は、気にしなくてOKです。ひらがなで入れても、カタカナで入れても、同じように探せます"],
  ["b", `漢字の名前を「ひらがなで」探したいときは、「${nm.userData}」シートの「よみ」の列に、よみ（ひらがな）を入れておきます。入っていない人は、漢字か、一部の言葉で探してください`],
  ["h", `「${nm.part}」シート`],
  ["b", "部品名の一部、ひらがな（カタカナの名前もOK）、部品番号の一部で探せます。ハイフンなしでもOKです"],
  ["b", `漢字の部品名を「ひらがなで」探したいときは、「${nm.partData}」シートの「よみ」の列に入れておきます`],
  ["h", "データを入れる（更新する）"],
  ["step", "①", `「${nm.userData}」シートのA〜D列に、ユーザーコード・正式名称・よみ（なくてもOK）・電話番号を貼ります（見出しは1行目）`],
  ["step", "②", `「${nm.partData}」シートのA〜C列に、部品番号・部品名・よみ（なくてもOK）を貼ります`],
  ["step", "③", "右の灰色の列は、数式です。消さないでください。貼り付けるだけで、検索に反映されます"],
  ["b", `最初に入っている行は、見本です。本物を貼る前に、A〜D列（部品はA〜C列）の見本を消してください`],
  ["b", `どちらのデータも、${LK_MAX.toLocaleString()}行目まで見ます`],
  ["b", `候補は、上から${SHOW}件まで出ます。多すぎるときは、言葉を足して絞ってください`],
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

const label = (ws, addr, text, o = {}) => {
  const c = ws.getCell(addr);
  c.value = text;
  c.font = { ...font, bold: true, ...o.font };
  c.alignment = { vertical: "middle", shrinkToFit: true, ...o.alignment };
};
const dbr = (sheet, col) => `${quote(sheet)}!$${col}$2:$${col}$${LK_MAX + 1}`;

// ---- 検索シート ----
// cfg: { title, sheet, data, cols:[{head,col,width}], hayCol, flagCol, runCol, promptNote }
function fillFinder(ws, nm, cfg, protect) {
  const nCols = cfg.cols.length;
  ws.columns = [{ width: 2 }, { width: 6 }, ...cfg.cols.map((c) => ({ width: c.width })), { width: 3 }];
  // 隠しの作業用セル（P列〜）：言葉を整えたもの
  for (let i = 0; i < 12; i++) ws.getColumn(16 + i).hidden = true;
  const lastCol = 2 + nCols;
  // タイトル
  ws.mergeCells(1, 2, 1, lastCol - 1);
  const t = ws.getCell(1, 2);
  t.value = cfg.title;
  t.font = { ...font, size: 18, bold: true, color: { argb: DK } };
  t.alignment = { vertical: "middle", indent: 1 };
  for (let i = 1; i <= lastCol; i++) ws.getCell(1, i).fill = solid(AC);
  const g = ws.getCell(1, lastCol);
  g.value = { formula: `HYPERLINK("#"&"${quote(nm.guide)}!A1","使い方 →")`, result: "使い方 →" };
  g.font = { ...font, color: { argb: "FF1A4FA0" }, underline: true };
  g.alignment = { vertical: "middle", horizontal: "right" };
  ws.getRow(1).height = 36;
  ws.getRow(2).height = 8;
  // 入力
  const inputs = [
    [3, cfg.label1],
    [4, "もう1つの言葉（絞り込み）"],
  ];
  for (const [r, text] of inputs) {
    ws.mergeCells(r, 2, r, 3);
    label(ws, `B${r}`, text);
    const endCol = lastCol;
    if (endCol > 4) ws.mergeCells(r, 4, r, endCol);
    const c = ws.getCell(r, 4);
    c.fill = solid(IN);
    c.font = { ...font, size: 15, bold: true };
    c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    c.protection = { locked: false };
    for (let i = 4; i <= endCol; i++) {
      ws.getCell(r, i).border = box;
      ws.getCell(r, i).fill = solid(IN);
    }
    ws.getRow(r).height = 32;
  }
  ws.mergeCells(5, 2, 5, 3);
  label(ws, "B5", "見つかった数");
  if (lastCol > 4) ws.mergeCells(5, 4, 5, lastCol);
  const run = dbr(cfg.data, cfg.runCol);
  ws.getCell("D5").value = {
    formula: `IF(MAX(${run})=0,"見つかりませんでした",IF(MAX(${run})>${SHOW},MAX(${run})&"件（上の${SHOW}件を表示。言葉を足して絞ってください）",MAX(${run})&"件"))`,
  };
  ws.getCell("D5").font = { ...font, size: 12, bold: true, color: { argb: "FFB06A00" } };
  ws.getCell("D5").alignment = { vertical: "middle", indent: 1 };
  ws.getRow(5).height = 24;
  ws.getRow(6).height = 8;
  // 作業用：整えた言葉（DBCSで全角にそろえ、小文字・空白なし）／カタカナ版／ひらがな版／ハイフンなし版
  for (const r of [3, 4]) {
    const raw = `$D$${r}`;
    ws.getCell(`P${r}`).value = { formula: `LOWER(SUBSTITUTE(SUBSTITUTE(${FN}(${raw}&""),"　","")," ",""))` };
    // カタカナ版 Q,R,S ／ ひらがな版 T,U,V ／ ハイフン・長音を除いた数字版 W
    H2K.forEach((pairs, i) => {
      ws.getCell(`${["Q", "R", "S"][i]}${r}`).value = { formula: chain(pairs, i === 0 ? `$P${r}` : `$${["Q", "R", "S"][i - 1]}${r}`) };
    });
    K2H.forEach((pairs, i) => {
      ws.getCell(`${["T", "U", "V"][i]}${r}`).value = { formula: chain(pairs, i === 0 ? `$P${r}` : `$${["T", "U", "V"][i - 1]}${r}`) };
    });
    ws.getCell(`W${r}`).value = { formula: `SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($P${r},"－",""),"‐",""),"−","")` };
  }
  // 見出し
  const HR = 7;
  ["No", ...cfg.cols.map((c) => c.head)].forEach((h, i) => {
    const c = ws.getCell(HR, 2 + i);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(HR).height = 28;
  const F = HR + 1;
  for (let k = 1; k <= SHOW; k++) {
    const r = F + k - 1;
    const strong = k <= STRONG;
    ws.getCell(`Y${r}`).value = { formula: `IFERROR(MATCH(${k},${run},0),"")` };
    ws.getCell(`B${r}`).value = { formula: `IF($Y${r}="","",${k})` };
    cfg.cols.forEach((col, i) => {
      const c = ws.getCell(r, 3 + i);
      c.value = { formula: `IF($Y${r}="","",INDEX(${dbr(cfg.data, col.col)},$Y${r})&"")` };
      c.font = col.mono ? { name: "Consolas", size: 12, bold: true } : font;
      c.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      c.border = strong ? box : hair;
    });
    const b = ws.getCell(`B${r}`);
    b.font = { ...font, color: { argb: "FF7A7F87" } };
    b.alignment = { vertical: "middle", horizontal: "center" };
    b.border = strong ? box : hair;
    ws.getRow(r).height = 24;
  }
  ws.getColumn(25).hidden = true;
  ws.views = [{ state: "frozen", ySplit: HR, showGridLines: false }];
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: `${HR}:${HR}` };
  if (protect) ws.sheetProtection = { sheet: true, formatColumns: true, formatRows: true };
}

// ---- データシート ----
// inCols: 入力の見出し（A〜）／ hay の式を作る関数 ／ 作業用の列の位置
function fillData(ws, nm, cfg, rows) {
  const n = cfg.inHeads.length;
  ws.columns = [...cfg.inWidths.map((width) => ({ width })), { width: 3 }, { width: 26 }, { width: 7 }, { width: 8 }];
  cfg.inHeads.forEach((h, i) => {
    const c = ws.getCell(1, i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = box;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  [["検索用（消さない）", cfg.hayCol], ["一致", cfg.flagCol], ["通し番号", cfg.runCol]].forEach(([h, col]) => {
    const c = ws.getCell(`${col}1`);
    c.value = h;
    c.font = { ...font, size: 9, bold: true, color: { argb: "FF7A7F87" } };
    c.fill = solid("FFEDEFF2");
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(1).height = 30;
  const S = quote(cfg.finder);
  const match = (r) => {
    const one = (row) =>
      `OR(${S}!$P$${row}="",ISNUMBER(SEARCH(${S}!$P$${row},$${cfg.hayCol}${r})),ISNUMBER(SEARCH(${S}!$S$${row},$${cfg.hayCol}${r})),ISNUMBER(SEARCH(${S}!$V$${row},$${cfg.hayCol}${r})),ISNUMBER(SEARCH(${S}!$W$${row},$${cfg.hayCol}${r})))`;
    return `AND(${one(3)},${one(4)})`;
  };
  for (let r = 2; r <= LK_MAX + 1; r++) {
    ws.getCell(`${cfg.hayCol}${r}`).value = { formula: cfg.hay(r) };
    ws.getCell(`${cfg.flagCol}${r}`).value = { formula: `IF($${cfg.hayCol}${r}="","",IF(${match(r)},1,0))` };
    ws.getCell(`${cfg.runCol}${r}`).value = { formula: `IF($${cfg.flagCol}${r}=1,1,0)+N(${cfg.runCol}${r - 1})` };
  }
  rows.forEach((row, i) => row.forEach((v, j) => v !== null && v !== undefined && (ws.getCell(i + 2, j + 1).value = v)));
  ws.getCell(`A1`).note = cfg.note;
  ws.views = [{ state: "frozen", ySplit: 1, showGridLines: true }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: n } };
  const noteCol = n + 5;
  const c = ws.getCell(1, noteCol);
  c.value = "← 見本の行は、本物を貼る前に消してください。右の灰色の列は数式です（消さないでください）";
  c.font = { ...font, size: 9, bold: true, color: { argb: "FFD93025" } };
}

// opt: { protect, sample:boolean, names? }
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

  // ユーザー：A コード B 正式名称 C よみ D 電話番号 ／ F 検索用 G 一致 H 通し番号
  const userHay = (r) => `IF($A${r}&$B${r}="","",${FN}(LOWER(SUBSTITUTE(SUBSTITUTE($A${r}&$B${r}&$C${r},"　","")," ","")))&"｜"&${FN}(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($D${r},"-","")," ",""),"　","")))`;
  // 部品：A 部品番号 B 部品名 C よみ ／ E 検索用 F 一致 G 通し番号
  const partHay = (r) => `IF($A${r}&$B${r}="","",${FN}(LOWER(SUBSTITUTE(SUBSTITUTE($A${r}&$B${r}&$C${r},"　","")," ","")))&"｜"&${FN}(SUBSTITUTE(SUBSTITUTE($A${r},"-","")," ","")))`;
  fillFinder(
    wsU,
    nm,
    {
      title: "ユーザー検索",
      label1: "名前・電話番号・コード",
      data: nm.userData,
      runCol: "H",
      cols: [
        { head: "ユーザーコード", col: "A", width: 16, mono: true },
        { head: "正式名称", col: "B", width: 44 },
        { head: "電話番号", col: "D", width: 22 },
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
      data: nm.partData,
      runCol: "G",
      cols: [
        { head: "部品番号", col: "A", width: 22, mono: true },
        { head: "部品名", col: "B", width: 50 },
      ],
    },
    protect,
  );
  let users = [];
  let parts = [];
  if (opt.sample) {
    const cos = [["見本自動車", "みほんじどうしゃ"], ["サンプル商事", "さんぷるしょうじ"], ["テスト整備工場", "てすとせいびこうじょう"], ["例示モータース", "れいじもーたーす"], ["ダミー運送", "だみーうんそう"], ["仮名車体", "かめいしゃたい"]];
    const people = [["山田", "やまだ"], ["鈴木", "すずき"], ["佐藤", "さとう"], ["田中", "たなか"], ["高橋", "たかはし"]];
    for (let i = 0; i < 30; i++) {
      const [co, cy] = cos[i % cos.length];
      const [pe, py] = people[(i * 2) % people.length];
      users.push([`Z${String(i + 1).padStart(3, "0")}`, `${co}${Math.floor(i / cos.length) + 1}号店 ${pe}さん`, `${cy}${Math.floor(i / cos.length) + 1}ごうてん ${py}さん`, `000-0000-${String(1000 + i * 3)}`]);
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
    { inHeads: ["ユーザーコード", "正式名称（会社名 担当者名）", "よみ（ひらがな・なくてもOK）", "電話番号"], inWidths: [14, 40, 30, 18], hayCol: "F", flagCol: "G", runCol: "H", finder: nm.user, hay: userHay, note: "ここに、ユーザーの一覧を貼ります（A〜D列、見出しは1行目）。よみは、なくてもOKです" },
    users,
  );
  fillData(
    wsPD,
    nm,
    { inHeads: ["部品番号", "部品名", "よみ（ひらがな・なくてもOK）"], inWidths: [20, 40, 30], hayCol: "E", flagCol: "F", runCol: "G", finder: nm.part, hay: partHay, note: "ここに、部品番号と部品名を貼ります（A〜C列、見出しは1行目）。よみは、なくてもOKです" },
    parts,
  );
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list: [nm.user, nm.part, nm.userData, nm.partData, nm.guide], names: nm };
}
