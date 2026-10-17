// スフィアライトの補足資料を、Excelのシートにする（画面で見る用／紙に印刷する用）
// 文章は sphdocs.js のHTMLから読み取って使う（資料を直すと、Excelにもそのまま反映される）

const AC = "FFF2C230";
const DK = "FF17191C";
const IN = "FFFFF3C4";
const thin = { style: "thin", color: { argb: "FF8A8F98" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

// ---- HTML → 読みやすい形（太字・改行つきの文字の並び）にする ----
const mergeRuns = (runs) => {
  const out = [];
  for (const r of runs) {
    if (!r.t) continue;
    const last = out[out.length - 1];
    if (last && !!last.b === !!r.b) last.t += r.t;
    else out.push({ ...r });
  }
  if (out.length) {
    out[0].t = out[0].t.replace(/^[\s]+/, "");
    out[out.length - 1].t = out[out.length - 1].t.replace(/[\s]+$/, "");
  }
  return out.filter((r) => r.t);
};
const runsOf = (node, bold = false, out = []) => {
  for (const n of node.childNodes) {
    if (n.nodeType === 3) out.push({ t: n.textContent.replace(/\s*\n\s*/g, ""), b: bold });
    else if (n.nodeType === 1) {
      if (n.nodeName === "BR") out.push({ t: "\n", b: bold });
      else runsOf(n, bold || n.nodeName === "B" || n.nodeName === "STRONG", out);
    }
  }
  return out;
};
const R = (node) => mergeRuns(runsOf(node));

export function docModel(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const title = (doc.querySelector(".sd-head h2") || {}).textContent || "補足資料";
  const sub = (doc.querySelector(".sd-sub") || {}).textContent || "";
  const sections = [...doc.querySelectorAll("section.sd-sec")].map((s) => {
    const blocks = [];
    for (const el of s.children) {
      const c = el.classList;
      if (el.nodeName === "H3") continue;
      if (el.nodeName === "H4") blocks.push({ t: "h4", runs: R(el) });
      else if (c.contains("sd-lead")) blocks.push({ t: "lead", runs: R(el) });
      else if (c.contains("sd-note")) blocks.push({ t: "note", runs: R(el) });
      else if (c.contains("sd-tw")) {
        const head = [...el.querySelectorAll("thead th")].map(R);
        const rows = [...el.querySelectorAll("tbody tr")].map((tr) => [...tr.children].map(R));
        blocks.push({ t: "table", head, rows });
      } else if (el.nodeName === "UL") blocks.push({ t: "list", items: [...el.children].map(R) });
      else if (c.contains("sd-box") || c.contains("sd-warn")) {
        const b = el.querySelector(":scope > b");
        blocks.push({
          t: "box",
          warn: c.contains("sd-warn"),
          title: b ? R(b) : [],
          items: [...el.querySelectorAll("li")].map(R),
        });
      } else if (c.contains("sd-qa")) {
        blocks.push({ t: "qa", q: R(el.querySelector(".sd-q")), a: R(el.querySelector(".sd-a")) });
      }
    }
    return { title: s.querySelector("h3").textContent, blocks };
  });
  return { title, sub, sections };
}

// ---- 文字の幅・行の高さの見積もり ----
const dispW = (s) => [...String(s)].reduce((a, ch) => a + (ch.charCodeAt(0) < 0x2e80 && !/[＀-￯]/.test(ch) ? 1 : 2), 0);
const text = (runs) => runs.map((r) => r.t).join("");
const linesOf = (runs, width, size) => {
  // 全角1文字 ＝ size×1.33px、1幅 ＝ 7px。余裕を見て1割ふやす
  const per = size * 0.0952 * 1.1;
  const cap = Math.max(4, width - 2);
  return text(runs)
    .split("\n")
    .reduce((n, ln) => n + Math.max(1, Math.ceil((dispW(ln) * per) / cap)), 0);
};
const heightOf = (lines, size) => lines * Math.round(size * 1.5) + 8;
const val = (runs, font) => {
  if (runs.length === 1 && !runs[0].b) return runs[0].t;
  return { richText: runs.map((r) => ({ text: r.t, font: { ...font, bold: font.bold || !!r.b } })) };
};

// ---- 組み立て ----
// cfg: { widths:[A,B,C,…], size, kind:'screen'|'print' }
function layout(model, cfg, o) {
  const W = cfg.widths; // 列の幅（A列から）
  const first = 2; // 記号の列（B）
  const last = W.length; // 最後の列
  const wOf = (c1, c2) => W.slice(c1 - 1, c2).reduce((a, b) => a + b, 0);
  const size = cfg.size;
  const font = { name: o.font.name, size };
  const units = []; // { rows:[{h, draw(ws,r)}], keep, sec }
  const unit = (rows, extra = {}) => units.push({ rows, ...extra });
  const spacer = (h = 8) => ({ h, draw() {} });
  const mergeTo = (ws, r, c1, c2) => c2 > c1 && ws.mergeCells(r, c1, r, c2);
  const paint = (ws, r, c1, c2, argb) => {
    for (let c = c1; c <= c2; c++) ws.getCell(r, c).fill = fill(argb);
  };
  const para = (runs, c1, opt = {}) => {
    const sz = opt.size || size;
    const f = { ...font, size: sz, ...(opt.font || {}) };
    return {
      h: heightOf(linesOf(runs, wOf(c1, last), sz), sz) + (opt.pad || 0),
      draw(ws, r) {
        mergeTo(ws, r, c1, last);
        const c = ws.getCell(r, c1);
        c.value = val(runs, f);
        c.font = f;
        c.alignment = { vertical: opt.marker ? "top" : "middle", wrapText: true, indent: opt.indent ?? 0 };
        if (opt.fill) paint(ws, r, opt.fillFrom || c1, last, opt.fill);
        if (opt.marker) {
          const m = ws.getCell(r, first);
          m.value = opt.marker;
          m.font = { ...f, bold: true, color: { argb: opt.markerColor || "FF7A7F87" } };
          m.alignment = { vertical: "top", horizontal: "center" };
        }
        if (opt.border) for (let c = first; c <= last; c++) ws.getCell(r, c).border = opt.border;
      },
    };
  };
  const bar = (title, sz, argb, color) => ({
    h: sz * 2.2,
    draw(ws, r) {
      mergeTo(ws, r, first, last);
      const c = ws.getCell(r, first);
      c.value = title;
      c.font = { ...font, size: sz, bold: true, color: { argb: color } };
      c.alignment = { vertical: "middle", indent: 1 };
      paint(ws, r, first, last, argb);
    },
  });

  // 表の列を、中身の長さに合わせて列に割り当てる（幅の合計がほぼ比例するように）
  const spansFor = (weights, c1, c2) => {
    const n = weights.length;
    const total = weights.reduce((a, b) => a + b, 0);
    const out = [];
    let c = c1;
    let acc = 0;
    for (let j = 0; j < n; j++) {
      acc += weights[j];
      const target = (acc / total) * wOf(c1, c2) + wOf(1, c1 - 1);
      let e = c;
      const mustLeave = n - 1 - j; // 後ろの列に最低1つずつ残す
      while (e < c2 - mustLeave && wOf(1, e) < target - 1) e++;
      if (j === n - 1) e = c2;
      out.push([c, e]);
      c = e + 1;
    }
    return out;
  };

  // ---- 表：画面用はそのまま表に ----
  const gridTable = (b) => {
    const nCol = b.head.length;
    const weights = b.head.map((h, j) =>
      Math.min(40, Math.max(7, Math.max(dispW(text(h)), ...b.rows.map((r) => Math.min(60, dispW(text(r[j] || []))) * 0.7)))),
    );
    const sp = spansFor(weights, first, last);
    const rows = [];
    rows.push({
      h: Math.max(...b.head.map((h, j) => heightOf(linesOf(h, wOf(...sp[j]), size), size))),
      draw(ws, r) {
        b.head.forEach((h, j) => {
          mergeTo(ws, r, sp[j][0], sp[j][1]);
          const c = ws.getCell(r, sp[j][0]);
          c.value = val(h, { ...font, bold: true });
          c.font = { ...font, bold: true, color: { argb: DK } };
          c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
          for (let k = sp[j][0]; k <= sp[j][1]; k++) {
            ws.getCell(r, k).fill = fill(AC);
            ws.getCell(r, k).border = box;
          }
        });
      },
    });
    for (const row of b.rows)
      rows.push({
        h: Math.max(...row.map((cell, j) => heightOf(linesOf(cell, wOf(...sp[j]), size), size))),
        draw(ws, r) {
          row.forEach((cell, j) => {
            mergeTo(ws, r, sp[j][0], sp[j][1]);
            const c = ws.getCell(r, sp[j][0]);
            const f = { ...font, bold: j === 0 };
            c.value = val(cell, f);
            c.font = f;
            c.alignment = { vertical: "middle", wrapText: true, indent: 0 };
            for (let k = sp[j][0]; k <= sp[j][1]; k++) {
              ws.getCell(r, k).border = box;
              if (j === 0) ws.getCell(r, k).fill = fill("FFFFF8DC");
            }
          });
        },
      });
    rows.push(spacer(6));
    unit(rows);
  };

  // ---- 表：印刷用は、1行ずつカードにして、大きな文字で縦に並べる ----
  const cardTable = (b) => {
    const label = (h) => text(h).replace(/\n/g, " ");
    const shortCells = false; // 項目ごとに1行ずつ並べるほうが読みやすい
    // 3列だけで行が多い表（例：適合できない車の一覧）は、1行にまとめる
    if (b.head.length === 3 && b.rows.length >= 6) {
      const half = W.length >= 4 ? 3 : first + 1; // 左（形状・メーカー）／右（車）
      const rows = [];
      rows.push({
        h: heightOf(1, size) + 4,
        draw(ws, r) {
          mergeTo(ws, r, first, half);
          const a = ws.getCell(r, first);
          a.value = `${label(b.head[0])}／${label(b.head[1])}`;
          mergeTo(ws, r, half + 1, last);
          const c = ws.getCell(r, half + 1);
          c.value = label(b.head[2]);
          for (const x of [a, c]) {
            x.font = { ...font, bold: true, color: { argb: DK } };
            x.alignment = { vertical: "middle", horizontal: "center" };
          }
          paint(ws, r, first, last, AC);
          for (let k = first; k <= last; k++) ws.getCell(r, k).border = box;
        },
      });
      for (const row of b.rows) {
        const l = [{ t: `${text(row[0])}　${text(row[1])}`, b: true }];
        rows.push({
          h: Math.max(heightOf(linesOf(l, wOf(first, half), size), size), heightOf(linesOf(row[2], wOf(half + 1, last), size), size)),
          draw(ws, r) {
            mergeTo(ws, r, first, half);
            const a = ws.getCell(r, first);
            a.value = text(l);
            a.font = { ...font, bold: true };
            a.alignment = { vertical: "middle", wrapText: true };
            mergeTo(ws, r, half + 1, last);
            const c = ws.getCell(r, half + 1);
            c.value = val(row[2], font);
            c.font = font;
            c.alignment = { vertical: "middle", wrapText: true };
            for (let k = first; k <= last; k++) ws.getCell(r, k).border = box;
            paint(ws, r, first, half, "FFFFF8DC");
          },
        });
      }
      rows.push(spacer(10));
      unit(rows);
      return;
    }
    const lc = first + 1; // 見出し項目の列
    for (const row of b.rows) {
      const rows = [];
      rows.push({
        h: heightOf(linesOf([{ t: label(row[0]) }], wOf(first, last), size + 2), size + 2) + 2,
        draw(ws, r) {
          mergeTo(ws, r, first, last);
          const c = ws.getCell(r, first);
          c.value = label(row[0]);
          c.font = { ...font, size: size + 2, bold: true, color: { argb: DK } };
          c.alignment = { vertical: "middle", indent: 1, wrapText: true };
          paint(ws, r, first, last, AC);
          for (let k = first; k <= last; k++) ws.getCell(r, k).border = box;
        },
      });
      if (shortCells) {
        // 短い項目は、1〜2行にまとめて並べる
        const line = row
          .slice(1)
          .map((c, j) => `${label(b.head[j + 1])}：${text(c).replace(/\n/g, " ")}`)
          .join("　");
        const runs = [{ t: line }];
        rows.push({
          h: heightOf(linesOf(runs, wOf(first, last), size), size),
          draw(ws, r) {
            mergeTo(ws, r, first, last);
            const c = ws.getCell(r, first);
            c.value = line;
            c.font = font;
            c.alignment = { vertical: "middle", wrapText: true, indent: 1 };
            for (let k = first; k <= last; k++) ws.getCell(r, k).border = box;
          },
        });
      } else {
        row.slice(1).forEach((cell, j) => {
          rows.push({
            // 項目名（左）が長くて2行になるときも、切れないように高さを合わせる
            h: Math.max(
              heightOf(linesOf(cell, wOf(lc + 1, last), size), size),
              heightOf(linesOf([{ t: label(b.head[j + 1]) }], wOf(first, lc) - 1, size), size),
            ),
            draw(ws, r) {
              mergeTo(ws, r, first, lc);
              const a = ws.getCell(r, first);
              a.value = label(b.head[j + 1]);
              a.font = { ...font, bold: true, color: { argb: "FF4A4F57" } };
              a.alignment = { vertical: "middle", wrapText: true, indent: 1 };
              mergeTo(ws, r, lc + 1, last);
              const c = ws.getCell(r, lc + 1);
              c.value = val(cell, font);
              c.font = font;
              c.alignment = { vertical: "middle", wrapText: true };
              paint(ws, r, first, lc, "FFEDEFF2");
              for (let k = first; k <= last; k++) ws.getCell(r, k).border = box;
            },
          });
        });
      }
      rows.push(spacer(10));
      unit(rows);
    }
  };

  // ---- 全体 ----
  const titleRows = [
    {
      h: size * 3.2,
      draw(ws, r) {
        mergeTo(ws, r, first, last);
        const c = ws.getCell(r, first);
        c.value = model.title;
        c.font = { ...font, size: size + 8, bold: true, color: { argb: DK } };
        c.alignment = { vertical: "middle", indent: 1 };
        paint(ws, r, first, last, AC);
      },
    },
    para([{ t: model.sub }], first, { font: { color: { argb: "FF4A4F57" } }, indent: 1, pad: 2 }),
    spacer(10),
  ];
  unit(titleRows, { keep: true });
  // もくじ（リンクつき）
  const toc = [bar("もくじ", size + 1, DK, "FFFFFFFF")];
  model.sections.forEach((s, i) => {
    toc.push({
      h: heightOf(1, size) + 4,
      tocIndex: i,
      draw(ws, r, ctx) {
        mergeTo(ws, r, first, last);
        const c = ws.getCell(r, first);
        const at = cfg.kind === "print" ? `${ctx.pages[i]}ページ目から` : "";
        const label = `${i + 1}．${s.title}${at ? "　　" + at : ""}`;
        c.value =
          cfg.kind === "print"
            ? label
            : { formula: `HYPERLINK("#"&"'${o.name.replace(/'/g, "''")}'!B${ctx.secRow[i]}","${label.replace(/"/g, '""')}")`, result: label };
        c.font = { ...font, color: { argb: cfg.kind === "print" ? DK : "FF1A4FA0" }, underline: cfg.kind !== "print", bold: cfg.kind === "print" };
        c.alignment = { vertical: "middle", indent: 1 };
        for (let k = first; k <= last; k++) ws.getCell(r, k).border = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
      },
    });
  });
  unit(toc, { keep: true });

  model.sections.forEach((s, si) => {
    let pending = [bar(`${si + 1}．${s.title}`, size + 4, DK, "FFFFFFFF"), spacer(6)];
    let secOpen = true; // 章の最初のまとまりに印をつける（もくじ・改ページ用）
    const mark = (u) => {
      if (secOpen && u) {
        u.sec = si;
        secOpen = false;
      }
    };
    const flush = (rows, extra = {}) => {
      unit([...pending, ...rows], extra);
      mark(units[units.length - 1]);
      pending = [];
    };
    for (const b of s.blocks) {
      if (b.t === "lead") flush([para(b.runs, first, { font: { color: { argb: "FF4A4F57" } }, indent: 1 }), spacer(6)], { keepNext: true });
      else if (b.t === "h4") {
        pending.push(
          {
            h: (size + 2) * 2,
            draw(ws, r) {
              mergeTo(ws, r, first, last);
              const c = ws.getCell(r, first);
              c.value = text(b.runs);
              c.font = { ...font, size: size + 2, bold: true, color: { argb: DK } };
              c.alignment = { vertical: "middle", indent: 1 };
              for (let k = first; k <= last; k++) ws.getCell(r, k).border = { bottom: { style: "medium", color: { argb: AC } } };
            },
          },
          spacer(6),
        );
      } else if (b.t === "note")
        flush([para(b.runs, first, { size: Math.max(10, size - 2), font: { color: { argb: "FF6B7078" } }, indent: 1 }), spacer(8)]);
      else if (b.t === "list") {
        b.items.forEach((it, i) => {
          const rows = [para(it, first + 1, { marker: "●", markerColor: "FFB06A00" })];
          if (i === b.items.length - 1) rows.push(spacer(8));
          flush(rows);
        });
      } else if (b.t === "box") {
        const bg = b.warn ? "FFFDECEA" : IN;
        const tc = b.warn ? "FFD93025" : DK;
        const rows = [];
        rows.push({
          h: heightOf(linesOf(b.title, wOf(first, last), size + 1), size + 1),
          draw(ws, r) {
            mergeTo(ws, r, first, last);
            const c = ws.getCell(r, first);
            c.value = (b.warn ? "⚠ " : "") + text(b.title);
            c.font = { ...font, size: size + 1, bold: true, color: { argb: tc } };
            c.alignment = { vertical: "middle", indent: 1, wrapText: true };
            paint(ws, r, first, last, bg);
          },
        });
        for (const it of b.items) rows.push(para(it, first + 1, { marker: "・", markerColor: tc, fill: bg, fillFrom: first }));
        rows.push(spacer(10));
        flush(rows);
      } else if (b.t === "qa") {
        const rows = [
          {
            h: heightOf(linesOf(b.q, wOf(first + 1, last), size + 1), size + 1),
            draw(ws, r) {
              const m = ws.getCell(r, first);
              m.value = "Q";
              m.font = { ...font, size: size + 2, bold: true, color: { argb: DK } };
              m.alignment = { vertical: "middle", horizontal: "center" };
              mergeTo(ws, r, first + 1, last);
              const c = ws.getCell(r, first + 1);
              c.value = text(b.q);
              c.font = { ...font, size: size + 1, bold: true, color: { argb: DK } };
              c.alignment = { vertical: "middle", wrapText: true, indent: 1 };
              paint(ws, r, first, last, AC);
            },
          },
          {
            h: heightOf(linesOf(b.a, wOf(first + 1, last), size), size),
            draw(ws, r) {
              const m = ws.getCell(r, first);
              m.value = "A";
              m.font = { ...font, size: size + 2, bold: true, color: { argb: "FFB06A00" } };
              m.alignment = { vertical: "top", horizontal: "center" };
              mergeTo(ws, r, first + 1, last);
              const c = ws.getCell(r, first + 1);
              c.value = val(b.a, font);
              c.font = font;
              c.alignment = { vertical: "top", wrapText: true, indent: 1 };
              for (let k = first; k <= last; k++) ws.getCell(r, k).border = { bottom: { style: "hair", color: { argb: "FFB0B4BA" } } };
            },
          },
          spacer(10),
        ];
        flush(rows);
      } else if (b.t === "table") {
        if (cfg.kind === "print") {
          const before = units.length;
          cardTable(b);
          if (pending.length && units.length > before) {
            units[before].rows.unshift(...pending);
            pending = [];
          }
          mark(units[before]);
        } else {
          const before = units.length;
          gridTable(b);
          if (pending.length) {
            units[before].rows.unshift(...pending);
            pending = [];
          }
          mark(units[before]);
        }
      }
    }
    if (pending.length) flush([]);
  });
  return { units };
}

// 章の最初のブロックを見つけるため、章ごとの先頭ユニットの番号を数える
function sectionStarts(model, units) {
  // units[0]=タイトル, units[1]=もくじ。以降は、章の先頭に bar が付いたユニット（sec が入っている）
  const starts = [];
  units.forEach((u, i) => {
    if (u.sec !== undefined && starts[u.sec] === undefined) starts[u.sec] = i;
  });
  return starts;
}

function render(ws, model, cfg, o) {
  const { units } = layout(model, cfg, o);
  const starts = sectionStarts(model, units);
  const ctx = { pages: [], secRow: [] };
  // 1回目：ページと行の位置を決める
  const place = () => {
    let r = 1;
    let used = 0;
    let page = 1;
    const breaks = [];
    const pos = [];
    let i = 0;
    while (i < units.length) {
      // keepNext でつながった塊を1つとして扱う
      let j = i;
      while (units[j].keepNext && j + 1 < units.length) j++;
      const group = units.slice(i, j + 1);
      const h = group.reduce((a, u) => a + u.rows.reduce((b, x) => b + x.h, 0), 0);
      const newSec = starts.includes(i) && cfg.kind === "print";
      if (cfg.kind === "print" && i > 0 && (newSec || used + h > cfg.cap) && used > 0) {
        breaks.push(r - 1);
        page++;
        used = 0;
      }
      group.forEach((u, k) => {
        pos[i + k] = r;
        if (starts.includes(i + k)) {
          ctx.pages[starts.indexOf(i + k)] = page;
          ctx.secRow[starts.indexOf(i + k)] = r;
        }
        r += u.rows.length;
        used += u.rows.reduce((b, x) => b + x.h, 0);
      });
      i = j + 1;
    }
    return { breaks, pos, total: r - 1 };
  };
  const { breaks, pos } = place();
  units.forEach((u, i) => {
    let r = pos[i];
    for (const row of u.rows) {
      ws.getRow(r).height = row.h;
      row.draw(ws, r, ctx);
      r++;
    }
  });
  for (const b of breaks) ws.getRow(b).addPageBreak();
}

// 画面で見る用：横長・12ptの表
export function fillDocScreen(ws, model, { font, protect, name }) {
  const widths = [2, 3, ...Array(12).fill(11)];
  ws.columns = widths.map((width) => ({ width }));
  ws.views = [{ showGridLines: false }];
  render(ws, model, { widths, size: 12, kind: "screen" }, { font, name });
  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true };
  if (protect) ws.sheetProtection = { sheet: true };
}

// 紙に印刷する用：A4たて・14pt・1つずつカード。章ごとに新しいページから
export function fillDocPrint(ws, model, { font, protect, name, date, title }) {
  const widths = [2, 4, 22, 66];
  ws.columns = widths.map((width) => ({ width }));
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = {
    paperSize: 9,
    orientation: "portrait",
    scale: 100,
    horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.55, bottom: 0.7, header: 0.25, footer: 0.3 },
  };
  ws.headerFooter = { oddFooter: `&L&11${title}（${date}）&R&12ページ &P / &N` };
  render(ws, model, { widths, size: 14, kind: "print", cap: 725 }, { font, name });
  if (protect) ws.sheetProtection = { sheet: true };
}
