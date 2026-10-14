// オイル・ケミカルを、関数だけで検索できるExcelにして書き出す（Excel 2016 で動く関数だけを使う）
const kataOf = (s) => s.replace(/[ぁ-ゖ]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0x60));
const hiraOf = (s) => s.replace(/[ァ-ヶ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x60));
const NF = (s) => String(s ?? "").normalize("NFKC").toLowerCase();

// 検索用の文字列：元の文字＋カタカナ版＋ひらがな版＋同義語（読み・別名）
// これをExcel側で SEARCH するので、ひらがな・読み・別名でも当たる（小さい「っ」「ょ」もそのまま）
// groups は「液状=えきじょう」のような同義語を [["液状","えきじょう"],…] にしたもの
export function searchText(o, groups) {
  const base = NF(`${o.code} ${String(o.code || "").replace(/[^0-9A-Za-z]/g, "")} ${o.name} ${o.cap} ${o.use} ${o.cat}`);
  const kata = kataOf(base);
  const extra = [];
  for (const g of groups) {
    const ws = g.map(NF).filter(Boolean);
    if (ws.some((w) => kata.includes(kataOf(w)))) for (const w of ws) extra.push(w, kataOf(w), hiraOf(w));
  }
  const flat = base.replace(/[\s\-・.／/]/g, ""); // 「0w-20」を「0w20」で探しても当たるように
  return [...new Set([base, flat, kata, hiraOf(base), ...extra])].join(" ").slice(0, 3000);
}

const MAXROWS = 200; // 検索シートに出す最大件数
const SPARE = 300; // データシートの空き行（あとから足せる）

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
  const wb = opt.wb || new ExcelJS.Workbook();
  if (!opt.wb) wb.creator = "セクション検索";
  wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nameS = freeName(wb, "検索");
  const QS = quote(nameS);
  const nameD = freeName(wb, "データ");
  const QD = quote(nameD);
  const N = items.length;
  const LAST = N + 1 + SPARE; // データシートの最終行
  const font = { name: "Yu Gothic UI", size: 11 };
  const AC = "FFF2C230";
  const DK = "FF17191C";
  const thin = { style: "thin", color: { argb: "FF8A8F98" } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };

  // 既存のシートが選ばれたままだと、まとめて編集されてしまうので外す
  for (const w of wb.worksheets) w.views = (w.views || []).map((v) => ({ ...v, tabSelected: false }));
  const ws = wb.addWorksheet(nameS, {
    views: [{ state: "frozen", ySplit: 6, showGridLines: false, tabSelected: true }],
  });
  const wd = wb.addWorksheet(nameD, { views: [{ state: "frozen", ySplit: 1 }] });

  // ---- データ ----
  wd.columns = [
    { header: "品番", width: 18 },
    { header: "名称", width: 44 },
    { header: "容量", width: 9 },
    { header: "用途", width: 44 },
    { header: "分類", width: 14 },
    { header: "", width: 3 },
    { header: "検索用（自動・触らない）", width: 18 },
    { header: "一致", width: 6 },
    { header: "順位", width: 6 },
    { header: "", width: 3 },
    { header: "分類一覧", width: 14 },
  ];
  const hr = wd.getRow(1);
  hr.font = { ...font, bold: true, color: { argb: DK } };
  hr.eachCell((c, i) => {
    if (i <= 5) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AC } };
  });
  const cats = [...new Set(items.map((o) => o.cat).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ja"));
  wd.getCell("K2").value = "すべて";
  cats.forEach((c, i) => (wd.getCell(`K${i + 3}`).value = c));
  for (let r = 2; r <= LAST; r++) {
    const o = items[r - 2];
    const row = wd.getRow(r);
    if (o) {
      row.getCell(1).value = o.code;
      row.getCell(2).value = o.name;
      row.getCell(3).value = o.cap || "";
      row.getCell(4).value = o.use || "";
      row.getCell(5).value = o.cat || "";
      row.getCell(7).value = searchText(o, groups);
    } else {
      // 空き行：あとから足した商品も、漢字・カタカナの直接入力で検索できる
      row.getCell(7).value = { formula: `IF($A${r}="","",LOWER($A${r}&" "&$B${r}&" "&$C${r}&" "&$D${r}&" "&$E${r}))` };
    }
    row.getCell(8).value = {
      formula: `IF($A${r}="",0,IF(AND(OR(${QS}!$K$2="",ISNUMBER(SEARCH(${QS}!$K$2,$G${r}))),OR(${QS}!$K$3="",ISNUMBER(SEARCH(${QS}!$K$3,$G${r}))),OR(${QS}!$K$4="",ISNUMBER(SEARCH(${QS}!$K$4,$G${r}))),OR(${QS}!$B$3="すべて",$E${r}=${QS}!$B$3)),1,0))`,
    };
    row.getCell(9).value = { formula: `IF($H${r}=1,COUNTIF($H$2:$H${r},1),"")` };
  }
  wd.autoFilter = { from: "A1", to: "E1" };
  for (const c of ["G", "H", "I"]) wd.getColumn(c).font = { ...font, color: { argb: "FF9AA0A6" } };
  wd.getCell("G1").font = { ...font, color: { argb: "FF9AA0A6" }, italic: true };

  // ---- 検索 ----
  ws.columns = [
    { width: 6 },
    { width: 18 },
    { width: 44 },
    { width: 9 },
    { width: 44 },
    { width: 14 },
    { width: 9 },
    { width: 3 },
    { width: 6 },
    { width: 3 },
    { width: 20 },
  ];
  ws.getCell("A1").value = "オイル・ケミカル検索";
  ws.getCell("A1").font = { ...font, size: 18, bold: true };
  ws.getCell("A2").value = "検索";
  ws.getCell("A3").value = "分類";
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
    formulae: [`${QD}!$K$2:$K$${cats.length + 2}`],
  };
  ws.getCell("C3").value = "← 黄色い欄に入力（品番・名称・用途・容量。空白で区切ると絞り込み。ひらがなもOK）";
  ws.getCell("C3").font = { ...font, size: 9, color: { argb: "FF7A7F87" } };
  ws.getCell("A4").value = {
    formula: `SUM(${QD}!$H$2:$H$${LAST})&" 件ヒット（全 "&COUNTA(${QD}!$A$2:$A$${LAST})&" 件）"&IF(SUM(${QD}!$H$2:$H$${LAST})>${MAXROWS},"　※先頭の${MAXROWS}件まで表示。語を足して絞ってください","")`,
  };
  ws.getCell("A4").font = { ...font, bold: true, color: { argb: "FFB06A00" } };
  // 検索語を3つまでに分ける（計算用・K列は非表示）
  ws.getCell("K1").value = "計算用";
  ws.getCell("K2").value = { formula: `IFERROR(LEFT($K$5,FIND(" ",$K$5)-1),$K$5)` };
  ws.getCell("K3").value = { formula: `IFERROR(LEFT($L$5,FIND(" ",$L$5)-1),$L$5)` };
  ws.getCell("K4").value = { formula: `IFERROR(MID($L$5,FIND(" ",$L$5)+1,200),"")` };
  ws.getCell("K5").value = { formula: `LOWER(TRIM(SUBSTITUTE($B$2,"　"," ")))` };
  ws.getCell("L5").value = { formula: `IFERROR(MID($K$5,FIND(" ",$K$5)+1,200),"")` };
  const H = ["No", "品番", "名称", "容量", "用途", "分類", "Google"];
  H.forEach((h, i) => {
    const c = ws.getRow(6).getCell(i + 1);
    c.value = h;
    c.font = { ...font, bold: true, color: { argb: DK } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AC } };
    c.border = box;
  });
  for (let k = 1; k <= MAXROWS; k++) {
    const r = 6 + k;
    const row = ws.getRow(r);
    row.getCell(9).value = { formula: `IFERROR(MATCH(${k},${QD}!$I$1:$I$${LAST},0),"")` };
    const at = (col) => ({ formula: `IF($I${r}="","",INDEX(${QD}!$${col}$1:$${col}$${LAST},$I${r})&"")` });
    row.getCell(1).value = { formula: `IF($I${r}="","",${k})` };
    row.getCell(2).value = at("A");
    row.getCell(3).value = at("B");
    row.getCell(4).value = at("C");
    row.getCell(5).value = at("D");
    row.getCell(6).value = at("E");
    row.getCell(7).value = {
      formula: `IF($I${r}="","",HYPERLINK("https://www.google.com/search?q="&_xlfn.ENCODEURL($B${r}&" "&$C${r}),"Google"))`,
    };
    for (let i = 1; i <= 7; i++) {
      const c = row.getCell(i);
      c.font = i === 7 ? { ...font, color: { argb: "FF1A73E8" }, underline: true } : font;
      c.border = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
      c.alignment = { vertical: "middle", wrapText: i === 3 || i === 5 };
    }
    row.getCell(2).font = { name: "Consolas", size: 11, bold: true };
  }
  ws.getColumn(9).hidden = true;
  ws.getColumn(11).hidden = true;
  ws.getColumn(12).hidden = true;
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  const idx = wb.worksheets.indexOf(ws);
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: idx, visibility: "visible" }];
  return { wb, nameS, nameD };
}
