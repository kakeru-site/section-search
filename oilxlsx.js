// オイル・ケミカルを、関数だけで検索できるExcelにして書き出す（Excel 2016 で動く関数だけを使う）
import { yomiText } from "./yomi.js?v=20261014x";
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
export function layoutFor(items) {
  const val = (o, f) => o[f];
  const maxW = (f) => Math.max(0, ...items.map((o) => dispW(val(o, f))));
  const w = {
    code: clamp(Math.ceil(maxW("code") * 1.15) + 2, 14, 28),
    name: clamp(Math.ceil(maxW("name") * 1.1) + 2, 24, 56),
    cap: clamp(Math.ceil(maxW("cap") * 1.15) + 2, 8, 16),
    use: clamp(Math.ceil(maxW("use") * 1.1) + 2, 24, 110),
    cat: clamp(Math.ceil(maxW("cat") * 1.15) + 2, 12, 26),
  };
  let lines = 1;
  for (const o of items)
    for (const f of Object.keys(w)) lines = Math.max(lines, Math.ceil((dispW(val(o, f)) * 1.1) / (w[f] - 1)));
  lines = Math.min(lines, 21); // Excelの行の高さは409ptまで
  return { w, lines, height: Math.max(24, lines * 19 + 4) };
}


// ---- 使い方シート ----
const GW = { b: 7, c: 28, d: 62 }; // 列の幅（B・C・D）
const guideBlocks = (protect, K) => [
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
  ["h", "困ったときは"],
  ["b", "何も出ない・計算されない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", `あるはずなのに出ない → 全角になっていないか、${K.catH}が「すべて」になっているか見てください`],
  ["b", "追加したのに出ない → 品番（A列）が空になっていないか、行を挿入していないか見てください"],
  ["b", "200件で止まる → 言葉を足して絞り込んでください"],
  ["b", "おかしくなった → 元のファイルをもう一回コピーすれば、元どおりです"],
  ["h", "これだけは守ってください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ["b", "検索欄に入れた言葉は、保存しなくて大丈夫です。閉じるときは「保存しない」でOK！"],
  ["b", "検索結果の表（数式が入ってます）は、消したり上書きしたりしないでください"],
  ...(protect
    ? [
        ["b", "「検索」と「使い方」のシートは、間違って消さないように保護してあります。入力できるのは黄色い欄とプルダウンだけです"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

function fillGuide(ws, { font, protect, K }) {
  const AC = "FFF2C230";
  const DK = "FF17191C";
  ws.columns = [{ width: 2 }, { width: GW.b }, { width: GW.c }, { width: GW.d }];
  const spanW = GW.c + GW.d; // 文章がまたぐ幅（C〜D）
  const lines = (text, width) => Math.max(1, Math.ceil((dispW(text) * 1.12) / (width - 2)));
  const hairs = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
  let r = 1;
  for (const blk of guideBlocks(protect, K)) {
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
  const nameD = opt.names?.data || freeName(wb, `${K.title}データ`);
  const QD = quote(nameD);
  const N = items.length;
  const LAST = N + 1 + SPARE; // データシートの最終行
  const font = { name: "Yu Gothic UI", size: 11 };
  const L = layoutFor(items);
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
    row.getCell(5).value = at("E", o?.use);
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
  fillGuide(wg, { font, protect: !!opt.protect, K });
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
  return { wb, nameS, nameG, nameD };
}
