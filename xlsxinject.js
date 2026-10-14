// 会社のExcelの中身には手を触れず、検索用のシートだけを後から差し込む（ZIPの中身を直接足す）
// ExcelJS などで読み直して書き出すと、元ファイルの機能が壊れることがあるため、この方式にしている

const esc = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const unesc = (s) =>
  String(s)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
const text = async (zip, path) => {
  const f = zip.file(path);
  return f ? f.async("string") : null;
};
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const uniqueName = (names, base) => {
  const low = new Set(names.map((n) => n.toLowerCase()));
  let n = base,
    i = 2;
  while (low.has(n.toLowerCase())) n = `${base}${i++}`;
  return n;
};

// 差し込む前に、元ファイルを読んでシート名・作成者などを調べる
export async function readBase(JSZip, buf) {
  const zip = await JSZip.loadAsync(buf);
  const rels = await text(zip, "_rels/.rels");
  if (!rels) throw new Error("Excelのファイルとして読めませんでした");
  const m = rels.match(/<Relationship\b[^>]*Type="[^"]*\/officeDocument"[^>]*>/);
  const target = m && (m[0].match(/Target="([^"]+)"/) || [])[1];
  const wbPath = target ? target.replace(/^\//, "") : "xl/workbook.xml";
  const wbXml = await text(zip, wbPath);
  if (!wbXml) throw new Error("ブックの本体が見つかりませんでした");
  if (/<fileSharing\b[^>]*reservationPassword|<workbookProtection\b[^>]*workbookPassword/.test(wbXml))
    throw new Error("パスワードで保護されています");
  const names = [...wbXml.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((x) => unesc(x[1]));
  const core = (await text(zip, "docProps/core.xml")) || "";
  const app = (await text(zip, "docProps/app.xml")) || "";
  const pick = (s, re) => {
    const x = s.match(re);
    return x ? unesc(x[1]).trim() : "";
  };
  return {
    zip,
    wbPath,
    names,
    creator: pick(core, /<dc:creator\b[^>]*>([\s\S]*?)<\/dc:creator>/),
    lastModifiedBy: pick(core, /<cp:lastModifiedBy\b[^>]*>([\s\S]*?)<\/cp:lastModifiedBy>/),
    company: pick(app, /<Company\b[^>]*>([\s\S]*?)<\/Company>/),
  };
}

// <fonts>・<fills>… の子を取り出す
const children = (xml, sec, tag) => {
  const m = xml.match(new RegExp(`<${sec}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${sec}>)`));
  if (!m || !m[1]) return [];
  return m[1].match(new RegExp(`<${tag}(?:\\s[^>]*)?/>|<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>`, "g")) || [];
};
// 末尾に足して count を直す。節が無ければ作る
function appendTo(xml, sec, tag, items, before) {
  if (!items.length) return xml;
  const open = new RegExp(`<${sec}\\b([^>]*?)(/>|>)`);
  const m = xml.match(open);
  if (!m) {
    // 節が無い：styleSheet の先頭側に作る（numFmts は最初、他は決まった順なので before を指定）
    const node = `<${sec} count="${items.length}">${items.join("")}</${sec}>`;
    if (before) {
      const i = xml.search(new RegExp(`<${before}\\b`));
      if (i >= 0) return xml.slice(0, i) + node + xml.slice(i);
    }
    return xml.replace(/(<styleSheet\b[^>]*>)/, `$1${node}`);
  }
  const have = children(xml, sec, tag).length;
  let attrs = m[1].replace(/\scount="\d+"/, "");
  let out;
  if (m[2] === "/>") {
    out = xml.replace(open, `<${sec}${attrs} count="${items.length}">${items.join("")}</${sec}>`);
  } else {
    out = xml.replace(open, `<${sec}${attrs} count="${have + items.length}">`);
    const close = out.indexOf(`</${sec}>`, out.search(open));
    out = out.slice(0, close) + items.join("") + out.slice(close);
  }
  return out;
}

// 検索シートのスタイルを、元ファイルの styles.xml に足して、番号の対応表を返す
function mergeStyles(bXml, aXml) {
  const offs = {};
  for (const [sec, tag, before] of [
    ["fonts", "font", "fills"],
    ["fills", "fill", "borders"],
    ["borders", "border", "cellStyleXfs"],
  ]) {
    offs[sec] = children(bXml, sec, tag).length;
    bXml = appendTo(bXml, sec, tag, children(aXml, sec, tag), before);
  }
  // 表示形式（ユーザー定義は164〜）
  const bFmts = [...bXml.matchAll(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"[^>]*\/>/g)];
  let maxId = Math.max(163, ...bFmts.map((x) => +x[1]));
  const fmtMap = {};
  const newFmts = [];
  for (const x of aXml.matchAll(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"[^>]*\/>/g)) {
    const same = bFmts.find((y) => y[2] === x[2]);
    if (same) fmtMap[x[1]] = same[1];
    else {
      fmtMap[x[1]] = String(++maxId);
      newFmts.push(`<numFmt numFmtId="${maxId}" formatCode="${x[2]}"/>`);
    }
  }
  bXml = appendTo(bXml, "numFmts", "numFmt", newFmts, null);
  // セルの書式
  const xfOff = children(bXml, "cellXfs", "xf").length;
  const xfs = children(aXml, "cellXfs", "xf").map((xf) =>
    xf
      .replace(/\sfontId="(\d+)"/, (_, n) => ` fontId="${+n + offs.fonts}"`)
      .replace(/\sfillId="(\d+)"/, (_, n) => ` fillId="${+n + offs.fills}"`)
      .replace(/\sborderId="(\d+)"/, (_, n) => ` borderId="${+n + offs.borders}"`)
      .replace(/\snumFmtId="(\d+)"/, (_, n) => ` numFmtId="${fmtMap[n] || n}"`)
      .replace(/\sxfId="\d+"/, ' xfId="0"'),
  );
  bXml = appendTo(bXml, "cellXfs", "xf", xfs, "cellStyles");
  return { xml: bXml, xfOff };
}

// ExcelJS が作った1枚のシートを、元ファイル用に書き換える（スタイル番号・文字列）
function remapSheet(xml, xfOff, strings) {
  xml = xml.replace(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, (all, attrs, tail, inner) => {
    let a = attrs.replace(/\ss="(\d+)"/, (_, n) => ` s="${+n + xfOff}"`);
    if (tail === "/>") return `<c${a}/>`;
    if (/\st="s"/.test(a)) {
      const idx = +((inner.match(/<v>(\d+)<\/v>/) || [])[1] ?? -1);
      a = a.replace(/\st="s"/, ' t="inlineStr"');
      return `<c${a}><is>${strings[idx] ?? "<t></t>"}</is></c>`;
    }
    return `<c${a}>${inner}</c>`;
  });
  xml = xml.replace(/<row\b[^>]*>/g, (t) => t.replace(/\ss="(\d+)"/, (_, n) => ` s="${+n + xfOff}"`));
  xml = xml.replace(/<col\b[^>]*>/g, (t) => t.replace(/\sstyle="(\d+)"/, (_, n) => ` style="${+n + xfOff}"`));
  return xml;
}

// base: readBase の結果 / ourBuf: ExcelJS が作った「検索」「データ」の2枚だけのxlsx
export async function injectSheets(JSZip, base, ourBuf, { nameS, nameD }) {
  const B = base.zip;
  const put = (p, d) => B.file(p, d, { createFolders: false });
  const A = await JSZip.loadAsync(ourBuf);
  const dir = base.wbPath.includes("/") ? base.wbPath.replace(/[^/]+$/, "") : "";
  const relsPath = `${dir}_rels/${base.wbPath.replace(/^.*\//, "")}.rels`;
  let wb = await text(B, base.wbPath);
  let rels = await text(B, relsPath);
  let ct = await text(B, "[Content_Types].xml");
  let styles = await text(B, `${dir}styles.xml`);
  if (!rels || !ct || !styles) throw new Error("このExcelの構成が想定と違うため、追加できませんでした");

  // --- スタイル ---
  const aStyles = await text(A, "xl/styles.xml");
  const merged = mergeStyles(styles, aStyles);
  put(`${dir}styles.xml`, merged.xml);

  // --- 文字列（共有文字列 → シート内の文字列に） ---
  const sst = (await text(A, "xl/sharedStrings.xml")) || "";
  const strings = [...sst.matchAll(/<si\s*\/>|<si>([\s\S]*?)<\/si>/g)].map((x) => x[1] ?? "<t></t>");

  // --- 新しいシートの置き場所 ---
  const origSheetFiles = Object.keys(B.files).filter((f) => /^xl\/(worksheets|chartsheets)\/[^/]+\.xml$/.test(f));
  const existing = origSheetFiles.filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
  let n = Math.max(0, ...existing.map((f) => +f.match(/sheet(\d+)\.xml/)[1]));
  const sheets = [
    [nameS, "xl/worksheets/sheet1.xml"],
    [nameD, "xl/worksheets/sheet2.xml"],
  ];
  const sheetIds = [...wb.matchAll(/<sheet\b[^>]*\bsheetId="(\d+)"/g)].map((x) => +x[1]);
  let sid = Math.max(0, ...sheetIds);
  let relsAdd = "",
    ctAdd = "",
    sheetsAdd = "";
  // 名前空間はこの要素の上で宣言する（元のファイルの書き方に左右されないように）
  const nsDecl = ` xmlns:r="${REL_NS}"`;
  const firstNew = sheetIds.length;
  for (const [name, src] of sheets) {
    n++;
    sid++;
    const path = `xl/worksheets/sheet${n}.xml`;
    const x = remapSheet(await text(A, src), merged.xfOff, strings);
    put(path, x);
    const rid = `rIdOil${n}`;
    relsAdd += `<Relationship Id="${rid}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${n}.xml"/>`;
    ctAdd += `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    sheetsAdd += `<sheet name="${esc(name)}" sheetId="${sid}"${nsDecl} r:id="${rid}"/>`;
  }
  rels = rels.replace("</Relationships>", relsAdd + "</Relationships>");
  ct = ct.replace("</Types>", ctAdd + "</Types>");
  wb = wb.replace("</sheets>", sheetsAdd + "</sheets>");

  // --- 計算の連鎖は作り直してもらう／開いたら再計算 ---
  const cc = rels.match(/<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*\/>/);
  if (cc) {
    const tg = (cc[0].match(/Target="([^"]+)"/) || [])[1];
    rels = rels.replace(cc[0], "");
    if (tg) {
      const p = tg.startsWith("/") ? tg.slice(1) : dir + tg;
      B.remove(p);
      ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`), "");
    }
  }
  if (/<calcPr\b/.test(wb)) {
    wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (_, a, s) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"${s}>`);
  } else {
    let pos = -1;
    for (const t of ["</definedNames>", "</externalReferences>", "</functionGroups>", "</sheets>"]) {
      const i = wb.indexOf(t);
      if (i >= 0) {
        pos = i + t.length;
        break;
      }
    }
    wb = wb.slice(0, pos) + '<calcPr fullCalcOnLoad="1"/>' + wb.slice(pos);
  }
  // 新しい検索シートを開いた状態にする（元のシートの「選択中」は外す）
  if (!/<bookViews\b/.test(wb) && /<sheets\b/.test(wb))
    wb = wb.replace(/<sheets\b/, `<bookViews><workbookView activeTab="${firstNew}"/></bookViews><sheets`);
  else if (/<workbookView\b/.test(wb))
    wb = wb.replace(/<workbookView\b([^>]*?)(\/?)>/, (_, a, s) => `<workbookView${a.replace(/\sactiveTab="\d+"/, "")} activeTab="${firstNew}"${s}>`);
  for (const f of origSheetFiles) {
    const x = await text(B, f);
    if (x && /\stabSelected="1"/.test(x)) put(f, x.replace(/\stabSelected="1"/g, ""));
  }
  put(base.wbPath, wb);
  put(relsPath, rels);
  put("[Content_Types].xml", ct);
  return B.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}
