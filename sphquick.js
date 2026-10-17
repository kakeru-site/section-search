// スフィアライトの「シリーズ・形状 × 色」の品番早見表（A4たて1枚に収めて印刷する前提）
// 登録されている商品の名前（例：ライジング3 H4 Hi/Lo 6000K）から、シリーズ・形状・色を読み取って表にする

export const QCOLS = [
  { id: "white", head: "6000K", sub: "ホワイト・6500K" },
  { id: "k45", head: "4500K", sub: "サンライト" },
  { id: "yellow", head: "2800K/3000K", sub: "黄色系" },
  { id: "dual", head: "デュアル", sub: "2色切替" },
  { id: "other", head: "その他", sub: "アンバー・電球色" },
];

// 名前の先頭から、シリーズを見分ける（長いものを先に）。shapeHead は、形状の頭につける言葉
const SERIES = [
  ["ライジングα Pro", "ライジングα Pro"],
  ["ライジング3", "ライジング3"],
  ["ライジングα", "ライジングα"],
  ["純正HID用LED 一体型", "純正HID用LED（一体型）"],
  ["純正HID用LED", "純正HID用LED（別体）"],
  ["フォグ用LED", "フォグ用LED"],
  ["純正LEDフォグ パワーアップバルブ", "パワーアップバルブ", "L1B（フォグ用）"],
  ["純正LEDバックランプ パワーアップバルブ", "パワーアップバルブ", "LW5（バックランプ用）"],
  ["サンターン2", "サンターン2（ウインカー用）"],
  ["補修用LEDバルブ", "補修用LEDバルブ"],
];
// 並べる順番（ライジング3 → α Pro → α → HID → フォグ …）
const ORDER = ["ライジング3", "ライジングα Pro", "ライジングα", "純正HID用LED（一体型）", "純正HID用LED（別体）", "フォグ用LED", "パワーアップバルブ", "サンターン2（ウインカー用）", "補修用LEDバルブ"];

const colorOf = (it) => {
  const seg = String(it.cap || "").split("／").pop() || "";
  const text = `${seg} ${it.name}`;
  if (/切替|デュアル/.test(text)) return "dual";
  const k = text.match(/(\d{4})\s*K/i);
  if (k) {
    const v = +k[1];
    if (v >= 5500) return "white";
    if (v >= 4000) return "k45";
    return "yellow";
  }
  if (/アンバー|電球色/.test(text)) return "other";
  return "white"; // 色の記載がないもの（補修用の白など）
};

export function quickModel(items) {
  const groups = new Map();
  for (const it of items) {
    let name = String(it.name || "").replace(/\s+/g, " ").trim();
    name = name.replace(/\s*(\d{4}\s*K|デュアルカラー)$/i, "").trim();
    let series = null;
    let head = "";
    let rest = name;
    for (const [pre, disp, shapeHead] of SERIES) {
      if (name === pre || name.startsWith(pre + " ")) {
        series = disp;
        head = shapeHead || "";
        rest = name.slice(pre.length).trim();
        break;
      }
    }
    if (!series) {
      series = "その他";
      rest = name;
    }
    const shape = head || rest || "—";
    const pi = SERIES.findIndex((s) => s[1] === series && (s[2] || "") === head);
    let g = groups.get(series);
    if (!g) groups.set(series, (g = { series, rows: new Map(), caps: new Set() }));
    const bright = String(it.cap || "").split("／")[0].trim();
    if (/lm/.test(bright)) g.caps.add(bright);
    let row = g.rows.get(shape);
    if (!row) g.rows.set(shape, (row = { shape, cells: {}, key: [pi < 0 ? 99 : pi, String(it.code || "")] }));
    if (String(it.code || "") < row.key[1]) row.key[1] = String(it.code || "");
    const col = colorOf(it);
    (row.cells[col] ||= []).push(it.code);
  }
  const list = [...groups.values()].sort((a, b) => {
    const ia = ORDER.indexOf(a.series);
    const ib = ORDER.indexOf(b.series);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
  });
  return list.map((g) => ({
    series: g.series,
    bright: g.caps.size === 1 ? [...g.caps][0] : "",
    rows: [...g.rows.values()].sort((a, b) => a.key[0] - b.key[0] || a.key[1].localeCompare(b.key[1])),
  }));
}

const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

// アプリの画面用（A4たて1枚の印刷にも使う）
export function quickHtml(model, date) {
  const head = `<tr><th class="q-s">シリーズ</th><th class="q-f">形状</th>${QCOLS.map((c) => `<th><b>${c.head}</b><small>${c.sub}</small></th>`).join("")}</tr>`;
  const body = model
    .map((g, gi) =>
      g.rows
        .map(
          (r, i) =>
            `<tr class="${gi % 2 ? "alt" : ""} ${i === 0 ? "gtop" : ""}">${
              i === 0
                ? `<th class="q-s" rowspan="${g.rows.length}"><b>${esc(g.series)}</b>${g.bright ? `<small>${esc(g.bright)}</small>` : ""}</th>`
                : ""
            }<th class="q-f">${esc(r.shape)}</th>${QCOLS.map(
              (c) => `<td>${r.cells[c.id] ? r.cells[c.id].map((x) => `<code>${esc(x)}</code>`).join("<br>") : '<span class="q-n">－</span>'}</td>`,
            ).join("")}</tr>`,
        )
        .join(""),
    )
    .join("");
  return `<div id="quick" class="quick">
    <div class="q-head"><div><h2>スフィアライト 品番早見表</h2><div class="q-sub">シリーズ・形状・色から、品番をさがす（${esc(date || "")}）</div></div>
      <button class="btn pri" data-qprint="1">A4で印刷する</button></div>
    <div class="q-tw"><table class="q-t"><thead>${head}</thead><tbody>${body}</tbody></table></div>
    <p class="q-note">※ 品番はすべて BS09L- から始まります。※ 明るさは2灯合計の規格値です。※ ヘッドライトに付けられるのは 6000K・4500K（2800K・デュアルカラーは車検非対応）。</p></div>`;
}

// ---- Excel：A4たて1枚 ----
const AC = "FFF2C230";
const DK = "FF17191C";
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const med = { style: "medium", color: { argb: "FF17191C" } };
const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

export function fillQuick(ws, model, { font, protect, date, title }) {
  const W = [1.5, 24, 27, 15.5, 15.5, 15.5, 15.5, 15.5]; // A（余白）, シリーズ, 形状, 色×5
  ws.columns = W.map((width) => ({ width }));
  ws.views = [{ showGridLines: false }];
  const last = W.length;
  // タイトル
  ws.mergeCells(1, 2, 1, last);
  const t = ws.getCell(1, 2);
  t.value = `${title} 品番早見表`;
  t.font = { ...font, size: 20, bold: true, color: { argb: DK } };
  t.alignment = { vertical: "middle", indent: 1 };
  for (let c = 2; c <= last; c++) ws.getCell(1, c).fill = fill(AC);
  ws.getRow(1).height = 38;
  ws.mergeCells(2, 2, 2, last);
  const s = ws.getCell(2, 2);
  s.value = "シリーズ・形状・色から品番をさがす。品番はすべて BS09L- から始まります";
  s.font = { ...font, size: 10, color: { argb: "FF4A4F57" } };
  s.alignment = { vertical: "middle", indent: 1 };
  ws.getRow(2).height = 20;
  // 見出し
  const hr = 3;
  const heads = ["シリーズ", "形状", ...QCOLS.map((c) => c.head)];
  const subs = ["", "", ...QCOLS.map((c) => c.sub)];
  heads.forEach((h, i) => {
    const c = ws.getCell(hr, i + 2);
    c.value = {
      richText: [
        { text: h, font: { ...font, size: 12, bold: true, color: { argb: DK } } },
        ...(subs[i] ? [{ text: `\n${subs[i]}`, font: { ...font, size: 8, color: { argb: "FF3A3F46" } } }] : []),
      ],
    };
    c.fill = fill(AC);
    c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    c.border = { top: med, bottom: med, left: thin, right: thin };
  });
  ws.getRow(hr).height = 42;
  // 本体
  let r = hr + 1;
  const rowH = 20;
  model.forEach((g, gi) => {
    const r0 = r;
    const tint = gi % 2 ? "FFEDEFF2" : "FFFFF3C4";
    g.rows.forEach((row, i) => {
      const ws_r = ws.getRow(r);
      ws_r.height = rowH * Math.max(1, ...QCOLS.map((c) => (row.cells[c.id] || []).length));
      const f = ws.getCell(r, 3);
      f.value = row.shape;
      f.font = { ...font, size: 11, bold: true };
      f.alignment = { vertical: "middle", wrapText: true, indent: 1 };
      f.fill = fill(gi % 2 ? "FFF7F8FA" : "FFFFFAE6");
      QCOLS.forEach((col, k) => {
        const c = ws.getCell(r, 4 + k);
        const v = row.cells[col.id];
        c.value = v ? v.join("\n") : "－";
        c.font = v ? { name: "Consolas", size: 11, bold: true } : { ...font, size: 10, color: { argb: "FFB0B4BA" } };
        c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      });
      for (let c = 3; c <= last; c++) {
        ws.getCell(r, c).border = { top: i === 0 ? med : { style: "hair", color: { argb: "FFB0B4BA" } }, bottom: { style: "hair", color: { argb: "FFB0B4BA" } }, left: thin, right: thin };
      }
      r++;
    });
    // シリーズ名（縦にまとめる）
    const sc = ws.getCell(r0, 2);
    if (r - 1 > r0) ws.mergeCells(r0, 2, r - 1, 2);
    sc.value = {
      richText: [
        { text: g.series, font: { ...font, size: 12, bold: true, color: { argb: DK } } },
        ...(g.bright ? [{ text: `\n${g.bright}（2灯）`, font: { ...font, size: 9, color: { argb: "FF4A4F57" } } }] : []),
      ],
    };
    sc.alignment = { vertical: "middle", horizontal: "left", wrapText: true, indent: 1 };
    for (let k = r0; k < r; k++) {
      ws.getCell(k, 2).fill = fill(tint);
      ws.getCell(k, 2).border = { top: k === r0 ? med : undefined, left: thin, right: thin, bottom: k === r - 1 ? thin : undefined };
    }
  });
  // 下の線・注意書き
  for (let c = 2; c <= last; c++) ws.getCell(r - 1, c).border = { ...ws.getCell(r - 1, c).border, bottom: med };
  ws.mergeCells(r + 1, 2, r + 1, last);
  const n = ws.getCell(r + 1, 2);
  n.value = "※ 明るさは2灯合計の規格値です。※ ヘッドライトに付けられるのは 6000K・4500K（2800K・デュアルカラーは、ヘッドライトだと車検非対応）。";
  n.font = { ...font, size: 9, color: { argb: "FF4A4F57" } };
  n.alignment = { vertical: "middle", wrapText: true, indent: 1 };
  ws.getRow(r + 1).height = 26;
  ws.pageSetup = {
    paperSize: 9,
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.2, footer: 0.25 },
  };
  ws.headerFooter = { oddFooter: `&L&9${title} 品番早見表（${date}）` };
  if (protect) ws.sheetProtection = { sheet: true };
}
