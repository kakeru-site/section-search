// 宛名シールを、関数だけで作れるExcel（部品番号ラベルと同じ構造）
//  ・シール作成（自由入力）：顧客マスターを使わず、上・下の文字をそのまま打ってシールにする
//  ・まとめて作成：顧客名・コードを何件か入れると、1件＝1ページで、まとめて印刷できる（あいまい検索・候補選びに対応）
//  ・顧客データ：顧客マスターのExcelを、そのまま丸ごとコピー貼り付けして使う（列の並びは、顧客マスターと同じ）
import { fillGuide } from "./oilxlsx.js?v=20261022y";
import { KANA_H2K, KANA_K2H, kanaChain, normExpr as normExprBase } from "./lookupxlsx.js?v=20261022y";
// 顧客コードは、半角にそろえる（部品番号と同じ考え方）
const codeNormExpr = (x) => normExprBase(x, "ASC");
// 名称は、全角にそろえる（ひらがな⇔カタカナの変換表が、全角カタカナ前提のため）
let FN = "_xlfn.DBCS"; // 確認用に、別の名前に替えられる（setAtenaFn）
export const setAtenaFn = (x) => {
  FN = x || "_xlfn.DBCS";
};
const nameNormExpr = (x) => normExprBase(x, FN);

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const font = { name: "Yu Gothic UI", size: 11 };
const atenaFont = (size, o = {}) => ({ name: "HGP創英角ゴシックUB", size, bold: true, ...o });
const quote = (n) => `'${n.replace(/'/g, "''")}'`;
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };

export const DATA_MAX = 2000; // 顧客マスターを見る行数（2行目〜）
export const BATCH_MAX = 30; // 「まとめて作成」で、一度に作れる件数

// 顧客マスターと、まったく同じ列の並び（丸ごとコピー貼り付けできるように）
const MASTER_HEADS = ["売上担当", "顧客コード", "取", "顧客名称（漢字）", "郵便番号", "住所（漢字）", "業種", "電話番号", "ＦＡＸ番号", "代表者名（漢字）", "支払日", "請求締日", "サマリー区分", "取引停止区分", "よみ（ひらがな・なくてもOK）"];
const CODE_COL = "B"; // 顧客コード
const NAME_COL = "D"; // 顧客名称（漢字）
const ADDR_COL = "F"; // 住所（漢字）
const YOMI_COL = "O"; // よみ（任意）

// ---- ラベル1枚ぶんの列・行（実際の宛名シール用紙と、まったく同じ寸法。A5横・2列×3行＝6枚） ----
// 実物のシール用紙のExcel（カーロード様シート）から、そのまま写した数値
const COL_W = [6.25, 39.875, 22.5, 39.875, 6.375]; // A・B・C・D・E（B,Dが2枚のラベル、C・Eは余白。印刷範囲はA〜E）
const ROW_H = [52, 57, 75]; // 住所の行・名称の行・次のラベルとの余白（3行×3ブロック＝9行）
const PRINT_TOP_MARGIN = 0; // 上の余白なし
const PRINT_SCALE = 77; // 実物と同じ、印刷の拡大縮小

function setAtenaCols(ws) {
  ws.columns = [...COL_W.map((width) => ({ width })), { width: 12 }]; // 最後の列は、画面の「使い方」リンク用（印刷はされない）
}
// シールに印刷するときだけ、「株式会社」→「㈱」「有限会社」→「㈲」に短くする（検索・顧客データ側はそのまま）
const abbrevCo = (x) => `SUBSTITUTE(SUBSTITUTE(${x},"有限会社","㈲"),"株式会社","㈱")`;
// シールに印刷するときだけ、文字の間の空白（全角・半角とも）を詰める（検索・顧客データ側はそのまま）
const squeezeSpaces = (x) => `SUBSTITUTE(SUBSTITUTE(${x}," ",""),"　","")`;
// ラベル1枚（B列 or D列、1つぶん）を描く。r0：そのラベルの先頭行。col：B=2 or D=4。addrF／nameF：中身の数式
// opts.addrSize／opts.addrAlign：上の欄（住所／自由入力の上の文字）の、文字の大きさ・寄せを変えたいときに指定
function paintAtena(ws, r0, col, addrF, nameF, opts = {}) {
  const a = ws.getCell(r0, col);
  a.value = { formula: squeezeSpaces(addrF) };
  a.font = atenaFont(opts.addrSize || 15);
  a.alignment = opts.addrAlign || { horizontal: "left", vertical: "center", wrapText: true };
  const n = ws.getCell(r0 + 1, col);
  n.value = { formula: abbrevCo(squeezeSpaces(nameF)) };
  n.font = atenaFont(36); // 実物のシール（カーロード様シート）と同じ大きさ
  n.alignment = { horizontal: "left", vertical: "center", shrinkToFit: true };
}
// 1ページ＝2列×3行＝6枚。pageTop：そのページの先頭の行
// withTrailingGap：最後（3つ目）のブロックの後ろにも、余白の行を確保するか。
// 「まとめて作成」のように、お客様ごとに次のページがすぐ続く場合は false にして、
// 1人分の行（実質8行）が、物理的に1ページぶんの高さを超えないようにする
function paintAtenaPage(ws, pageTop, addrF, nameF, withTrailingGap = true, opts = {}) {
  for (let blk = 0; blk < 3; blk++) {
    const r0 = pageTop + blk * 3;
    paintAtena(ws, r0, 2, addrF, nameF, opts);
    paintAtena(ws, r0, 4, addrF, nameF, opts);
    ws.getRow(r0).height = ROW_H[0];
    ws.getRow(r0 + 1).height = ROW_H[1];
    if (blk < 2 || withTrailingGap) ws.getRow(r0 + 2).height = ROW_H[2];
  }
}

const guideBlocks = (protect, nm) => [
  ["title", "宛名シール作成の使い方"],
  ["sub", "顧客マスターから、住所・名称入りの宛名シール（実際のシール用紙・A5・6枚/1ページ）を作って印刷するシートです。"],
  ["h", "まずはこれだけ（データを入れる）"],
  ["step", "①", `「${nm.data}」シートに、顧客マスターのExcelの中身を、そのまま丸ごとコピーして、貼り付けます（A1から。見出しの1行目も、そのまま上書きしてOKです）`],
  ["b", `列の並びは、顧客マスターと同じです（${MASTER_HEADS.slice(0, 4).join("・")}…）。並びを変えずに、そのまま貼り付けてください`],
  ["b", "ラベルの名称は、印刷のときだけ「株式会社」→「㈱」「有限会社」→「㈲」に短くなります（顧客データ・検索は、元の名称のままです）"],
  ["h", `「${nm.free}」シート（顧客マスターを使わず、好きな文字でシールを作る）`],
  ["step", "①", "上の黄色いところに1行目（例：「釧路本店様分」）、下の黄色いところに2行目（例：「ビッグオートボディ」）を、そのまま打ちます"],
  ["step", "②", "下に、同じラベルが6枚、すぐ出ます。そのまま印刷します"],
  ["b", "上の文字は、長ければ2行まで折り返します。下の文字は、1行に収まるよう自動で縮みます"],
  ["h", `「${nm.batch}」シート（まとめて、何件か印刷）`],
  ["step", "①", "上の表の、黄色い「顧客名・コード」欄に、必要な分だけ、入れます（1行に1件）。お客様名の一部（あいまい検索）でも、顧客コードでもOKです"],
  ["b", "お客様名であいまいに探すと、同じ言葉を含む顧客が複数いる場合、ひとまずどれか1件に自動で決まり、「名称」に「（候補◯件・左で選べます）」と出ます。左の「候補」欄をクリックすると▽が出るので、一覧から選び直せます"],
  ["step", "②", `${BATCH_MAX}件まで入れられます。右に、名称と、印刷ページの番号が出ます`],
  ["step", "③", "下に、入れた件数ぶんのページが、自動でできます（1件＝1ページ＝6枚）"],
  ["step", "④", "印刷するときは、印刷の画面で「ページ指定」に、入れた件数のページ数だけを指定してください（上の表に、ページ数が出ています）。指定しないと、空っぽのページまで印刷されます"],
  ["h", "文字の大きさ・太さを変えたいとき"],
  ["b", "ラベルの文字（住所・名称）は、クリックして選んでから、普通のExcelと同じように、上のリボンの「フォントサイズ」「Ｂ（太字）」で、その場で自由に変えられます（中身の文字そのものは、保護されていて直接は打てません）"],
  ["h", "困ったとき"],
  ["b", "何も出ない・計算されない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", `見つからない → 「${nm.data}」シートに、顧客マスターのデータが、正しく貼り付けてあるか確認してください`],
  ["b", "住所が長くて3行目が切れる → 住所の欄は、2行までしか表示されません（シール用紙の大きさに合わせてあります）"],
  ["b", "データが古い → 新しい顧客マスターを、また丸ごとコピーして、貼り付け直してください（前のデータは上書きされます）"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "顧客の名称・住所は、個人情報・取引先情報を含みます。社外に出さないでください"],
  ["b", "ラベルのシートは、見た目を合わせるため、印刷の範囲をあらかじめ決めてあります。右の列や、上の入力欄は、印刷されません"],
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
  c.border = box;
  c.protection = { locked: false };
  c.numFmt = "@"; // 文字として受ける
  c.font = { ...font, size: o.size || 14, bold: true };
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
};
// セルの書式変更を許可している（文字の大きさ等をその場で変えられるように）ため、
// 別の場所からの貼り付けで、黄色い塗りが上書きされて消えることがある。
// 条件付き書式で「常に黄色」を重ねておき、貼り付けで消えても、見た目は黄色に戻るようにする
const yellowGuard = (ws, ref) => {
  ws.addConditionalFormatting({ ref, rules: [{ type: "expression", formulae: ["TRUE"], style: { fill: solid(IN) } }] });
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

// ---- データシート（顧客マスターと同じ列。丸ごと貼り付け用） ----
function fillData(ws) {
  ws.columns = MASTER_HEADS.map((h, i) => ({ width: i === 3 || i === 4 ? 30 : i === 5 ? 36 : i === 14 ? 22 : 11 }));
  MASTER_HEADS.forEach((h, i) => {
    const c = ws.getCell(1, i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
  });
  const info = ws.getCell(1, MASTER_HEADS.length + 3);
  info.value = `← 顧客マスターのExcelを、A1から、見出しごと、そのまま丸ごとコピーして貼り付けます（${DATA_MAX.toLocaleString()}行目まで見ます）`;
  info.font = { ...font, size: 9, bold: true, color: { argb: "FFD93025" } };
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// ---- 検索用シート（隠し）：顧客データの写し＋あいまい検索の計算。
// データシートとは別のシートに置くことで、顧客マスターの貼り付けが、どれだけ広い範囲でも巻き込まれないようにしている ----
function fillCalc(ws, nm) {
  const D = quote(nm.data);
  ["コード", "名称", "住所", "よみ", "検索用コード", "検索用（名称）"].forEach((h, i) => {
    const c = ws.getCell(1, i + 1);
    c.value = h;
    c.font = { ...font, size: 9, bold: true };
  });
  const codeNorm = (addr) => `SUBSTITUTE(SUBSTITUTE(${codeNormExpr(addr)},"-",""),"ー","")`;
  for (let r = 2; r <= DATA_MAX + 1; r++) {
    const A = `$A${r}`,
      B = `$B${r}`,
      Dd = `$D${r}`;
    ws.getCell(`A${r}`).value = { formula: `${D}!$${CODE_COL}${r}&""` };
    ws.getCell(`B${r}`).value = { formula: `${D}!$${NAME_COL}${r}&""` };
    ws.getCell(`C${r}`).value = { formula: `${D}!$${ADDR_COL}${r}&""` };
    ws.getCell(`D${r}`).value = { formula: `${D}!$${YOMI_COL}${r}&""` };
    ws.getCell(`E${r}`).value = { formula: `IF(${A}="","",${codeNorm(A)})` };
    ws.getCell(`F${r}`).value = { formula: `IF(${B}&${A}="","",${nameNormExpr(`${B}&${A}&${Dd}`)})` };
  }
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// ---- 自由入力ラベル（顧客マスターを使わず、上・下、好きな文字を直接打ってシールにする） ----
function fillFree(ws, nm) {
  setAtenaCols(ws);
  titleBar(ws, "シール作成（自由入力・すぐ印刷）", 6, nm.guide);
  label2(ws, "B3", "上の文字", { font: { size: 12 } });
  ws.mergeCells("C3:D3");
  yellow(ws, "C3", { size: 14 });
  yellowGuard(ws, "C3:D3");
  ws.getCell("C3").alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
  ws.getRow(3).height = 34;
  label2(ws, "B4", "下の文字", { font: { size: 12 } });
  ws.mergeCells("C4:D4");
  yellow(ws, "C4", { size: 16 });
  yellowGuard(ws, "C4:D4");
  ws.getRow(4).height = 30;
  ws.getRow(5).height = 10;
  const GRID_TOP = 6;
  paintAtenaPage(ws, GRID_TOP, "$C$3", "$C$4", true, { addrSize: 18, addrAlign: { horizontal: "center", vertical: "middle", wrapText: true } });
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = { paperSize: 11, orientation: "landscape", scale: PRINT_SCALE, margins: { left: 0, right: 0, top: PRINT_TOP_MARGIN, bottom: 0, header: 0, footer: 0 }, printArea: `A$${GRID_TOP}:E$${GRID_TOP + 7}` };
  return ws;
}

// まとめて作成の、各お客様のページの先頭行（お客様ごとに8行）。改ページを手動で入れる場所の計算にも使う
const BATCH_HR = 3;
const BATCH_LIST_TOP = BATCH_HR + 1;
export const BATCH_GRID_TOP = BATCH_LIST_TOP + BATCH_MAX + 1;
// Excelの自動改ページ（行の高さの合計から計算）にまかせると、端数のずれで、お客様の境目とページの境目がズレることがあるため、
// 1件＝1ページになるよう、お客様の最後の行のあとに、改ページを手動で指定する
export const atenaBatchBreakRows = () => Array.from({ length: BATCH_MAX - 1 }, (_, i) => BATCH_GRID_TOP + i * 8 + 7);

// まとめて作成の、入力欄の並び
const BATCH_COLS = { no: 2, code: 3, cand: 4, name: 5, status: 6 }; // B・C・D・E・F
const thinBorder = { style: "thin", color: { argb: "FFB7BBC2" } };
const batchBox = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };
const colLetter = (n) => String.fromCharCode(64 + n);
// A・B・C…Z・AA・AB… という、Excelの列名（26を超えてもOK）
const colName = (n) => {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};
const BATCH_CAND_N = 50; // あいまい検索で、お客様ごとに出す候補の数（まとめて作成。これだけ当てはまることは、まず無いはずの件数）
// 候補の行番号（CI）・候補の表示名（CN）・候補の件数（CNT）を置く、隠し列（P〜Xの続きから）
const CI_COL = (k) => colName(24 + k); // 1〜BATCH_CAND_N
const CN_COL = (k) => colName(24 + BATCH_CAND_N + k); // 1〜BATCH_CAND_N
const CNT_COL_IDX = 24 + 2 * BATCH_CAND_N + 1;
const CNT_COL = colName(CNT_COL_IDX);

// ---- まとめて作成 ----
function fillBatch(ws, nm) {
  setAtenaCols(ws);
  titleBar(ws, `まとめて作成（1件＝1ページ・最大${BATCH_MAX}件）`, 6, nm.guide);
  const HR = BATCH_HR;
  [
    [BATCH_COLS.no, "No"],
    [BATCH_COLS.code, "顧客名・コード"],
    [BATCH_COLS.cand, "候補（複数のとき）"],
    [BATCH_COLS.name, "名称"],
    [BATCH_COLS.status, "印刷ページ"],
  ].forEach(([col, h]) => {
    const c = ws.getCell(HR, col);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = solid(AC);
    c.border = batchBox;
    c.alignment = { horizontal: "center", vertical: "middle" };
  });
  ws.getRow(HR).height = 22;
  const LIST_TOP = HR + 1;
  const calcCode = `${quote(nm.calc)}!$E$2:$E$${DATA_MAX + 1}`;
  const calcName = `${quote(nm.calc)}!$F$2:$F$${DATA_MAX + 1}`;
  const calcNameAnchor = `${quote(nm.calc)}!$F$2`;
  const calcNameDisp = `${quote(nm.calc)}!$B$2:$B$${DATA_MAX + 1}`;
  const calcCodeDisp = `${quote(nm.calc)}!$A$2:$A$${DATA_MAX + 1}`;
  const hyphenStrip = (expr) => ["－", "‐", "−", "-", "ー", "―", "–", "—", "‑", "ｰ", " ", "　"].reduce((a, h) => `SUBSTITUTE(${a},"${h}","")`, expr);
  for (let i = 0; i < BATCH_MAX; i++) {
    const r = LIST_TOP + i;
    const inp = `$C$${r}`;
    const pick = `$${colLetter(BATCH_COLS.cand)}$${r}`;
    const no = ws.getCell(r, BATCH_COLS.no);
    no.value = i + 1;
    no.alignment = { horizontal: "center" };
    no.font = { ...font, color: { argb: "FF8A8F98" } };
    no.border = batchBox;
    yellow(ws, `C${r}`, { size: 12 });
    // あいまい検索（1件ずつのラベル作成と同じ考え方）：①コード完全一致 ②名称完全一致 ③名称の一部（ひらがな・カタカナ・ハイフン抜きを順に試す）
    const codeNormF = `SUBSTITUTE(SUBSTITUTE(${codeNormExpr(inp)},"-",""),"ー","")`;
    ws.getCell(`Q${r}`).value = { formula: nameNormExpr(`${inp}&""`) };
    const letters1 = ["R", "S", "T"];
    KANA_H2K.forEach((pairs, i2) => (ws.getCell(`${letters1[i2]}${r}`).value = { formula: kanaChain(pairs, i2 === 0 ? `$Q${r}` : `$${letters1[i2 - 1]}${r}`) }));
    const letters2 = ["U", "V", "W"];
    KANA_K2H.forEach((pairs, i2) => (ws.getCell(`${letters2[i2]}${r}`).value = { formula: kanaChain(pairs, i2 === 0 ? `$Q${r}` : `$${letters2[i2 - 1]}${r}`) }));
    ws.getCell(`X${r}`).value = { formula: hyphenStrip(`$Q${r}`) };
    // 候補（お客様名があいまいに何件かに当てはまるとき、1件ずつ・前の候補の続きから探す）
    for (let k = 1; k <= BATCH_CAND_N; k++) {
      const ciCell = `$${CI_COL(k)}$${r}`;
      const prevCi = k === 1 ? "" : `$${CI_COL(k - 1)}$${r}`;
      const ciF =
        k === 1
          ? `IFERROR(MATCH("*"&$Q${r}&"*",${calcName},0),"")`
          : `IF(${prevCi}="","",IFERROR(${prevCi}+MATCH("*"&$Q${r}&"*",OFFSET(${calcNameAnchor},${prevCi},0,${DATA_MAX}-${prevCi},1),0),""))`;
      ws.getCell(`${CI_COL(k)}${r}`).value = { formula: ciF };
      ws.getCell(`${CN_COL(k)}${r}`).value = { formula: `IF(${ciCell}="","",INDEX(${calcNameDisp},${ciCell})&"（"&INDEX(${calcCodeDisp},${ciCell})&"）")` };
    }
    const cnRange = `$${CN_COL(1)}$${r}:$${CN_COL(BATCH_CAND_N)}$${r}`;
    const ciRangeFull = `$${CI_COL(1)}$${r}:$${CI_COL(BATCH_CAND_N)}$${r}`;
    ws.getCell(`${CNT_COL}${r}`).value = { formula: `COUNT(${ciRangeFull})` };
    const cand = ws.getCell(r, BATCH_COLS.cand);
    cand.numFmt = "@";
    cand.font = { ...font, size: 10 };
    cand.fill = solid(IN);
    cand.border = batchBox;
    cand.alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
    cand.protection = { locked: false };
    cand.dataValidation = { type: "list", allowBlank: true, formulae: [cnRange], showErrorMessage: false };
    // 選んだ候補（CHOOSEは引数254個までのため、件数を増やせるようINDEX+MATCHにしている）
    // 「候補」欄は、入力欄をあとで消しても文字が残るので、今の候補の一覧に実際にある（ISNUMBER(MATCH…)）ときだけ使う。
    // 残ったままの古い候補は、ここで無視され、自動のあいまい検索に戻る
    const pickValidF = `AND(${pick}<>"",ISNUMBER(MATCH(${pick},${cnRange},0)))`;
    const idxF = `IF(${inp}="","",IFERROR(MATCH(${codeNormF},${calcCode},0),IF(${pickValidF},INDEX(${ciRangeFull},MATCH(${pick},${cnRange},0)),IFERROR(MATCH($Q${r},${calcName},0),IF($${CI_COL(1)}$${r}<>"",$${CI_COL(1)}$${r},IFERROR(MATCH("*"&$T${r}&"*",${calcName},0),IFERROR(MATCH("*"&$W${r}&"*",${calcName},0),IFERROR(MATCH("*"&$X${r}&"*",${calcName},0),""))))))))`;
    ws.getCell(`P${r}`).value = { formula: idxF };
    const name = ws.getCell(r, BATCH_COLS.name);
    name.value = {
      formula: `IF(${inp}="","",IF($P$${r}="","見つかりません",INDEX(${quote(nm.calc)}!$B$2:$B$${DATA_MAX + 1},$P$${r})&IF(AND($${CNT_COL}$${r}>1,NOT(${pickValidF})),"　（候補"&$${CNT_COL}$${r}&"件・左で選べます）","")))`,
    };
    name.font = font;
    name.border = batchBox;
    name.alignment = { vertical: "middle", wrapText: true };
    const st = ws.getCell(r, BATCH_COLS.status);
    st.value = { formula: `IF(${inp}="","",IF($${colLetter(BATCH_COLS.name)}$${r}="見つかりません","見つかりません",${i + 1}))` };
    st.font = { ...font, color: { argb: "FFB06A00" }, bold: true };
    st.alignment = { horizontal: "center", vertical: "middle" };
    st.border = batchBox;
    ws.getRow(r).height = 18;
  }
  for (let n = 16; n <= CNT_COL_IDX; n++) ws.getColumn(n).hidden = true; // P〜（検索・候補の計算用）
  // 「候補」欄は普通の入力欄なので、顧客名・コードを消しても文字は残る。見た目だけでも分かるよう、
  // 顧客名・コードが空のときは、候補欄の文字色を背景色と同じにして、見えなくする（計算は上のpickValidFで無視済み）
  const candCol = colLetter(BATCH_COLS.cand);
  const codeCol = colLetter(BATCH_COLS.code);
  ws.addConditionalFormatting({
    ref: `${candCol}${LIST_TOP}:${candCol}${LIST_TOP + BATCH_MAX - 1}`,
    rules: [{ type: "expression", formulae: [`$${codeCol}${LIST_TOP}=""`], style: { font: { color: { argb: IN } } } }],
  });
  // 一覧へのコピー貼り付けで、黄色い塗りが上書きされて消えることがあるため、条件付き書式で黄色を重ねておく
  yellowGuard(ws, `${codeCol}${LIST_TOP}:${codeCol}${LIST_TOP + BATCH_MAX - 1}`);
  yellowGuard(ws, `${candCol}${LIST_TOP}:${candCol}${LIST_TOP + BATCH_MAX - 1}`);
  ws.mergeCells(2, 2, 2, 5);
  ws.getCell("B2").value = { formula: `IF(COUNTIF($C$${LIST_TOP}:$C$${LIST_TOP + BATCH_MAX - 1},"<>")=0,"顧客名（またはコード）を入れると、下にページができます","入力 "&COUNTIF($C$${LIST_TOP}:$C$${LIST_TOP + BATCH_MAX - 1},"<>")&"件　→　印刷は、ページ指定で「1-"&COUNTIF($C$${LIST_TOP}:$C$${LIST_TOP + BATCH_MAX - 1},"<>")&"」")` };
  ws.getCell("B2").font = { ...font, size: 11, bold: true, color: { argb: "FFB06A00" } };
  ws.getRow(LIST_TOP + BATCH_MAX).height = 10;
  const GRID_TOP = BATCH_GRID_TOP;
  for (let i = 0; i < BATCH_MAX; i++) {
    const addrF = `IF($P$${LIST_TOP + i}="","",IF(INDEX(${quote(nm.calc)}!$C$2:$C$${DATA_MAX + 1},$P$${LIST_TOP + i})="","分",INDEX(${quote(nm.calc)}!$C$2:$C$${DATA_MAX + 1},$P$${LIST_TOP + i})))`;
    const nameF = `IF($P$${LIST_TOP + i}="","",INDEX(${quote(nm.calc)}!$B$2:$B$${DATA_MAX + 1},$P$${LIST_TOP + i}))`;
    paintAtenaPage(ws, GRID_TOP + i * 8, addrF, nameF, false);
  }
  ws.views = [{ showGridLines: false, state: "frozen", ySplit: HR }];
  ws.pageSetup = { paperSize: 11, orientation: "landscape", scale: PRINT_SCALE, margins: { left: 0, right: 0, top: PRINT_TOP_MARGIN, bottom: 0, header: 0, footer: 0 }, printArea: `A$${GRID_TOP}:E$${GRID_TOP + 8 * BATCH_MAX - 1}` };
  return ws;
}

// opt: { protect, names?: {find,batch,data,calc,guide}, sample?:boolean, rows?: 顧客マスターの行（配列の配列。MASTER_HEADSの並び） }
export function buildAtenaWorkbook(ExcelJS, opt = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = {
    free: opt.names?.free || "シール作成（自由入力）",
    batch: opt.names?.batch || "まとめて作成",
    data: opt.names?.data || "顧客データ",
    calc: opt.names?.calc || "検索用（隠し）",
    guide: opt.names?.guide || "宛名シールの使い方",
  };
  const protect = !!opt.protect;
  const wsFree = wb.addWorksheet(nm.free, { views: [{ showGridLines: false, tabSelected: true }] });
  wsFree.properties.dyDescent = 0.4;
  const wsBatch = wb.addWorksheet(nm.batch, { views: [{ showGridLines: false }] });
  wsBatch.properties.dyDescent = 0.4;
  const wsData = wb.addWorksheet(nm.data, { views: [{ state: "frozen", ySplit: 1 }] });
  const wsGuide = wb.addWorksheet(nm.guide, { views: [{ showGridLines: false }] });
  const wsCalc = wb.addWorksheet(nm.calc, { state: "veryHidden" });
  const list = [nm.free, nm.batch, nm.data, nm.guide, nm.calc];

  fillData(wsData);
  const rows = opt.rows || (opt.sample ? sampleRows() : []);
  rows.forEach((row, i) => row.forEach((v, j) => v !== null && v !== undefined && (wsData.getCell(i + 2, j + 1).value = v)));
  fillCalc(wsCalc, nm);

  fillFree(wsFree, nm);
  fillBatch(wsBatch, nm);
  if (opt.sample) {
    wsFree.getCell("C3").value = "釧路本店様分";
    wsFree.getCell("C4").value = "ビッグオートボディ";
  }

  fillGuide(wsGuide, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  fillGuideStyle(wsGuide, { title: "宛名シール作成の使い方" });

  if (protect) {
    // formatCells：true にして、ラベルの文字（大きさ・太さなど）を、見ながらその場で自由に変えられるようにしている
    // （数式・中身は、locked（入力不可）のまま変わらない。黄色い欄だけが、中身も変えられる）
    wsFree.sheetProtection = { sheet: true, formatColumns: true, formatRows: true, formatCells: true };
    wsBatch.sheetProtection = { sheet: true, formatColumns: true, formatRows: true, formatCells: true };
    wsCalc.sheetProtection = { sheet: true };
  }
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list, names: nm };
}

function sampleRows() {
  const co = ["見本自動車株式会社", "サンプル商事", "テスト整備工場 釧路店", "例示モータース", "ダミー運送株式会社"];
  const addr = ["釧路市柳町１３番１９号", "釧路郡釧路町中央６丁目１９番地", "釧路市鳥取大通３丁目１１番地１７", "釧路郡釧路町東陽大通西３丁目４番地１４", "釧路市貝塚３丁目５－２７"];
  return co.map((c, i) => {
    const row = Array(MASTER_HEADS.length).fill("");
    row[1] = `Z${String(100 + i).padStart(6, "0")}`;
    row[2] = "1";
    row[3] = c;
    row[5] = addr[i];
    return row;
  });
}
