// 会社のExcelの中身には手を触れず、検索用のシートだけを後から差し込む（ZIPの中身を直接足す）
// ExcelJS などで読み直して書き出すと、元ファイルの機能が壊れることがあるため、この方式にしている

const esc = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
// &#35211; のような「文字の番号」の書き方も、元の文字に戻す
const numRef = (s) => String(s).replace(/&#(x[0-9a-fA-F]+|\d+);/g, (_, n) => String.fromCodePoint(n[0] === "x" || n[0] === "X" ? parseInt(n.slice(1), 16) : +n));
const unesc = (s) =>
  numRef(s)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
// UTF-8として正しく読めて、書き戻しても1バイトも変わらないことを確かめてから使う（文字化けで壊さないため）
const text = async (zip, path) => {
  const f = zip.file(path);
  if (!f) return null;
  const bytes = await f.async("uint8array");
  let s;
  try {
    s = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch {
    throw new Error(`${path} の文字コードが想定と違うため、追加できませんでした`);
  }
  const back = new TextEncoder().encode(s);
  if (back.length !== bytes.length || back.some((b, i) => b !== bytes[i]))
    throw new Error(`${path} を安全に読み込めなかったため、追加できませんでした`);
  return s;
};
// 関係（rels）の Target を、ZIPの中の場所に直す（"/xl/…" の絶対指定も、"../…" の相対指定も）
const resolveTarget = (fromFile, target) => {
  if (target.startsWith("/")) return target.slice(1);
  const parts = fromFile.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg && seg !== ".") parts.push(seg);
  }
  return parts.join("/");
};
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const uniqueName = (names, base) => {
  const low = new Set(names.map((n) => n.toLowerCase()));
  let n = base,
    i = 2;
  while (low.has(n.toLowerCase())) n = `${base}${i++}`;
  return n;
};

// 自分（ExcelJS）が作ったばかりのxlsxに、指定したシートの「ここで改ページ」を、生のXMLで足す。
// ExcelJSには、手動改ページを作るAPIが無いため、この方式にしている。
// rows：そのシートの「この行の直後で改ページ」という行番号（1始まり）の配列
export async function addRowBreaks(JSZip, buf, sheetName, rows) {
  if (!rows.length) return buf;
  const zip = await JSZip.loadAsync(buf);
  const wbXml = (await text(zip, "xl/workbook.xml")) || "";
  const m = wbXml.match(new RegExp(`<sheet\\b[^>]*\\bname="${reEsc(esc(sheetName))}"[^>]*\\br:id="([^"]+)"`));
  if (!m) return buf; // 見つからなければ、何もしない（壊さない）
  const relsXml = (await text(zip, "xl/_rels/workbook.xml.rels")) || "";
  const rm = relsXml.match(new RegExp(`<Relationship\\b[^>]*\\bId="${m[1]}"[^>]*\\bTarget="([^"]+)"`));
  if (!rm) return buf;
  const path = resolveTarget("xl/workbook.xml", rm[1]);
  const xml = await text(zip, path);
  if (!xml || !/<pageSetup\b[^>]*\/>/.test(xml)) return buf;
  const brk = rows.map((r) => `<brk id="${r}" max="16383" man="1"/>`).join("");
  const node = `<rowBreaks count="${rows.length}" manualBreakCount="${rows.length}">${brk}</rowBreaks>`;
  zip.file(path, xml.replace(/(<pageSetup\b[^>]*\/>)/, `$1${node}`), { createFolders: false });
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

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
  // 条件付き書式の色（dxf）。シート側の dxfId は、足したぶんだけずらす
  const dxfOff = children(bXml, "dxfs", "dxf").length;
  const dxfs = children(aXml, "dxfs", "dxf");
  if (dxfs.length) {
    if (/<dxfs\b/.test(bXml)) bXml = appendTo(bXml, "dxfs", "dxf", dxfs, null);
    else {
      const node = `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>`;
      const at = ["<tableStyles", "<colors", "<extLst", "</styleSheet>"].map((t) => bXml.indexOf(t)).filter((i) => i >= 0);
      const i = Math.min(...at);
      bXml = bXml.slice(0, i) + node + bXml.slice(i);
    }
  }
  return { xml: bXml, xfOff, dxfOff };
}

// ExcelJS が作った1枚のシートを、元ファイル用に書き換える（スタイル番号・文字列）
function remapSheet(xml, xfOff, strings, dxfOff = 0) {
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
  if (dxfOff) xml = xml.replace(/<cfRule\b[^>]*>/g, (t) => t.replace(/\sdxfId="(\d+)"/, (_, n) => ` dxfId="${+n + dxfOff}"`));
  xml = xml.replace(/<row\b[^>]*>/g, (t) => t.replace(/\ss="(\d+)"/, (_, n) => ` s="${+n + xfOff}"`));
  xml = xml.replace(/<col\b[^>]*>/g, (t) => t.replace(/\sstyle="(\d+)"/, (_, n) => ` style="${+n + xfOff}"`));
  return xml;
}

// base: readBase の結果 / ourBuf: ExcelJS が作った「検索」「使い方」「印刷用」「データ」の4枚だけのxlsx（この順）
export async function injectSheets(JSZip, base, ourBuf, { nameS, nameG, nameP, nameD, list, srcPaths, srcIndex, extraNames, renames, report }) {
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
  const existing = Object.keys(B.files).filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
  let n = Math.max(0, ...existing.map((f) => +f.match(/sheet(\d+)\.xml/)[1]));
  // list：差し込むシート名を、ブックに入っている順に並べたもの（省略すると 検索・使い方・印刷用・データ）
  const generic = !!srcPaths; // 自分で作ったExcelから、シートをそのままコピーするとき
  const sheets = list
    ? list.map((nm, i) => [nm, srcPaths?.[i] || `xl/worksheets/sheet${i + 1}.xml`])
    : [
        [nameS, "xl/worksheets/sheet1.xml"],
        [nameG, "xl/worksheets/sheet2.xml"],
        ...(nameP ? [[nameP, "xl/worksheets/sheet3.xml"]] : []),
        [nameD, `xl/worksheets/sheet${nameP ? 4 : 3}.xml`],
      ];
  const sheetIds = [...wb.matchAll(/<sheet\b[^>]*\bsheetId="(\d+)"/g)].map((x) => +x[1]);
  let sid = Math.max(0, ...sheetIds);
  let relsAdd = "",
    ctAdd = "",
    sheetsAdd = "";
  const mediaMap = new Map(); // 絵のファイル：元の場所 → 差し込み先の名前
  let drawN = 0;
  let mediaN = 0;
  const mediaExt = new Set();
  // 名前空間はこの要素の上で宣言する（元のファイルの書き方に左右されないように）
  const nsDecl = ` xmlns:r="${REL_NS}"`;
  // 差し込む側のシートの「隠す／すごく隠す」設定（検索用の隠しシートなど）を、そのまま引き継ぐ
  const srcWbXml = (await text(A, "xl/workbook.xml")) || "";
  const srcStates = [...srcWbXml.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => (m[0].match(/\bstate="([^"]*)"/) || [])[1] || "visible");
  for (const [i, [name, src]] of sheets.entries()) {
    n++;
    sid++;
    const path = `xl/worksheets/sheet${n}.xml`;
    const state = srcStates[i] || "visible";
    // 元のシートの「選択中」には触れない代わりに、こちらの「選択中」を外す（複数選択でまとめて編集されないように）
    let x = remapSheet(await text(A, src), merged.xfOff, strings, merged.dxfOff).replace(/\stabSelected="1"/g, "");
    if (renames?.length) x = renameRefs(x, renames);
    put(path, x);
    // 図（絵）があるシートは、図の部品（drawing・media）も、重ならない名前で一緒に差し込む
    const sRelsPath = src.replace("worksheets/", "worksheets/_rels/") + ".rels";
    const sRels = await text(A, sRelsPath);
    const keepIds = new Set();
    if (sRels) {
      let newRels = "";
      for (const m of sRels.matchAll(/<Relationship\b[^>]*\/>/g)) {
        const attr = (k) => (m[0].match(new RegExp(`\\b${k}="([^"]*)"`)) || [])[1];
        const type = attr("Type") || "";
        const target = attr("Target") || "";
        if (!/\/drawing$/.test(type)) continue; // 図の関係だけを引き継ぐ
        const dSrc = resolveTarget(src, target);
        // グラフ・図形だけの描画は、部品がそろわないためコピーしない（画像だけの描画は、そのままコピー）
        if (generic) {
          const dXml = (await text(A, dSrc)) || "";
          const dRelsX = (await text(A, dSrc.replace("drawings/", "drawings/_rels/") + ".rels")) || "";
          if (/<c:chart\b|\/chart"|<mc:AlternateContent|<xdr:sp\b|<xdr:graphicFrame/.test(dXml + dRelsX)) {
            report?.push(`「${name}」のグラフ・図形は、コピーできませんでした（セルの内容はコピーしました）`);
            continue;
          }
        }
        keepIds.add(attr("Id"));
        const dNew = `xl/drawings/oilgdrawing${++drawN}.xml`;
        put(dNew, await A.file(dSrc).async("uint8array"));
        ctAdd += `<Override PartName="/${dNew}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
        // 図の中の絵への関係
        const dRelsSrc = dSrc.replace("drawings/", "drawings/_rels/") + ".rels";
        const dRels = await text(A, dRelsSrc);
        if (dRels) {
          const fixed = [...dRels.matchAll(/<Relationship\b[^>]*\/>/g)]
            .map((rm) => {
              const t = (rm[0].match(/\bTarget="([^"]*)"/) || [])[1] || "";
              const mSrc = resolveTarget(dSrc, t);
              if (!mediaMap.has(mSrc)) {
                const ext = (mSrc.match(/\.(\w+)$/) || [, "png"])[1].toLowerCase();
                mediaExt.add(ext);
                mediaMap.set(mSrc, `xl/media/oilgimg${++mediaN}.${ext}`);
              }
              return rm[0].replace(/\bTarget="[^"]*"/, `Target="../${mediaMap.get(mSrc).replace(/^xl\//, "")}"`);
            })
            .join("");
          put(dNew.replace("drawings/", "drawings/_rels/") + ".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${fixed}</Relationships>`);
        }
        newRels += m[0].replace(/\bTarget="[^"]*"/, `Target="../drawings/oilgdrawing${drawN}.xml"`);
      }
      if (newRels)
        put(`xl/worksheets/_rels/sheet${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${newRels}</Relationships>`);
    }
    if (generic) put(path, sanitizeSheet(x, keepIds, report, name));
    const rid = `rIdOil${n}`;
    relsAdd += `<Relationship Id="${rid}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${n}.xml"/>`;
    ctAdd += `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    sheetsAdd += `<sheet name="${esc(name)}" sheetId="${sid}"${state !== "visible" ? ` state="${state}"` : ""}${nsDecl} r:id="${rid}"/>`;
  }
  // 印刷範囲・見出し行の繰り返し（シート名に紐づく定義）を、こちらのシートぶんだけ足す
  const aWb = (await text(A, "xl/workbook.xml")) || "";
  const baseCount = [...wb.matchAll(/<sheet\b/g)].length; // 差し込む前のシート数
  const ours = [...aWb.matchAll(/<definedName\b[^>]*\bname="(_xlnm\.Print_(?:Area|Titles))"[^>]*\blocalSheetId="(\d+)"[^>]*>([\s\S]*?)<\/definedName>/g)];
  const order = srcIndex || sheets.map((_, i) => i); // 差し込むシートが、元のブックで何番目だったか
  const namesAdd =
    ours
      .map((x) => {
        const at = order.indexOf(+x[2]);
        return at < 0 ? "" : `<definedName name="${x[1]}" localSheetId="${baseCount + at}">${x[3]}</definedName>`;
      })
      .join("") +
    (extraNames || [])
      .filter((x) => {
        const dup = new RegExp(`<definedName\\b[^>]*\\bname="${reEsc(x.name)}"`).test(wb) && !/localSheetId/.test(x.xml);
        if (dup) report?.push(`名前「${x.name}」は、会社のExcelにすでにあるため、コピーしませんでした`);
        return !dup;
      })
      .map((x) => x.xml)
      .join("");
  if (namesAdd) {
    if (/<definedNames\s*\/>/.test(wb)) wb = wb.replace(/<definedNames\s*\/>/, `<definedNames>${namesAdd}</definedNames>`);
    else if (wb.includes("</definedNames>")) wb = wb.replace("</definedNames>", namesAdd + "</definedNames>");
    else {
      // 決まった順番（sheets → externalReferences → definedNames → calcPr…）の位置に入れる
      const after = ["<calcPr", "<oleSize", "<customWorkbookViews", "<pivotCaches", "<smartTagPr", "<smartTagTypes", "<webPublishing", "<fileRecoveryPr", "<webPublishObjects", "<extLst", "</workbook>"]
        .map((t) => wb.indexOf(t))
        .filter((i) => i >= 0);
      const at = Math.min(...after);
      wb = wb.slice(0, at) + `<definedNames>${namesAdd}</definedNames>` + wb.slice(at);
    }
  }
  // 会社のExcelが「手動計算」のままだと、貼り付けたデータや数式が、開いただけでは反映されない。
  // 開いたときに必ず再計算されるよう、計算方式の設定を上書きする
  if (/<calcPr\b[^>]*\/>/.test(wb)) {
    wb = wb.replace(/<calcPr\b([^>]*)\/>/, (m, attrs) => {
      const a = attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, "").replace(/\s*calcMode="[^"]*"/, "");
      return `<calcPr${a} calcMode="auto" fullCalcOnLoad="1"/>`;
    });
  } else {
    const after = ["<oleSize", "<customWorkbookViews", "<pivotCaches", "<smartTagPr", "<smartTagTypes", "<webPublishing", "<fileRecoveryPr", "<webPublishObjects", "<extLst", "</workbook>"].map((t) => wb.indexOf(t)).filter((i) => i >= 0);
    const at = Math.min(...after);
    wb = wb.slice(0, at) + `<calcPr calcMode="auto" fullCalcOnLoad="1"/>` + wb.slice(at);
  }
  rels = rels.replace("</Relationships>", relsAdd + "</Relationships>");
  // 絵のファイルを書き、種類（png など）の登録が無ければ足す
  for (const [from, to] of mediaMap) put(to, await A.file(from).async("uint8array"));
  for (const ext of mediaExt)
    if (!new RegExp(`<Default\\b[^>]*Extension="${ext}"`, "i").test(ct))
      ctAdd += `<Default Extension="${ext}" ContentType="image/${ext === "jpg" ? "jpeg" : ext}"/>`;
  // Default は Override より前に置く決まりなので、先頭側に入れる
  const defaults = [...ctAdd.matchAll(/<Default\b[^>]*\/>/g)].map((m) => m[0]).join("");
  ctAdd = ctAdd.replace(/<Default\b[^>]*\/>/g, "");
  if (defaults) ct = ct.replace(/(<Types\b[^>]*>)/, `$1${defaults}`);
  ct = ct.replace("</Types>", ctAdd + "</Types>");
  wb = wb.replace("</sheets>", sheetsAdd + "</sheets>");

  put(base.wbPath, wb);
  put(relsPath, rels);
  put("[Content_Types].xml", ct);
  return B.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}


// ---- 自分で作ったExcelから、シートをコピーするための部品 ----
const reEsc = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
// シート名を変えたとき、数式の中の「元の名前!」を新しい名前に直す
const q = (n) => n.replace(/'/g, "''");
// 数式の中で、このシート名を使っているか（'シート名'! と、引用符なしの シート名! の両方）
export const mentionsSheet = (body, n) => body.includes(`'${q(n)}'!`) || new RegExp(`(^|[^\\p{L}\\p{N}_'.])${reEsc(n)}!`, "u").test(body);
export function renameBody(body, renames) {
  let b = numRef(body);
  for (const [from, to] of renames) {
    b = b.split(`'${q(from)}'!`).join(`'${q(to)}'!`);
    b = b.replace(new RegExp(`(^|[^\\p{L}\\p{N}_'.])${reEsc(from)}!`, "gu"), `$1'${q(to)}'!`);
  }
  return b;
}
function renameRefs(xml, renames) {
  return xml.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (all, attrs, body) => {
    const dec = numRef(body); // 文字の番号の書き方を、元の文字に戻してから探す
    const b = renameBody(dec, renames);
    return b === dec ? all : `<f${attrs}>${b}</f>`;
  });
}
// 部品がそろわないもの（表・コメント・外部リンク・グラフなど）を、シートから外す（セルの内容は残す）
function sanitizeSheet(xml, keepIds, report, name) {
  const dropped = new Set();
  const drop = (re, label) => {
    xml = xml.replace(re, () => {
      dropped.add(label);
      return "";
    });
  };
  drop(/<tableParts\b[\s\S]*?<\/tableParts>|<tableParts\b[^>]*\/>/g, "表（テーブル）の設定");
  drop(/<legacyDrawing\b[^>]*\/>/g, "コメント");
  drop(/<legacyDrawingHF\b[^>]*\/>/g, "");
  drop(/<oleObjects\b[\s\S]*?<\/oleObjects>/g, "埋め込みオブジェクト");
  drop(/<controls\b[\s\S]*?<\/controls>/g, "フォームのボタン");
  drop(/<picture\b[^>]*\/>/g, "");
  drop(/<extLst\b[\s\S]*?<\/extLst>/g, "");
  // 他のファイルを参照する数式は、数式をやめて、計算済みの値だけを残す（リンクの部品がないと、Excelが開けなくなるため）
  xml = xml.replace(/<f\b[^>]*>[^<]*\[\d+\][^<]*<\/f>/g, () => {
    dropped.add("他のファイルを参照する数式（値だけ残しました）");
    return "";
  });
  // 外のファイルやページへのリンクは外す（シート内へのリンクは残す）
  xml = xml.replace(/<hyperlink\b[^>]*\/>/g, (t) => (/\br:id="/.test(t) ? (dropped.add("Webページ・ファイルへのリンク"), "") : t));
  xml = xml.replace(/<hyperlinks>\s*<\/hyperlinks>/g, "");
  xml = xml.replace(/<drawing\b[^>]*\br:id="([^"]*)"[^>]*\/>/g, (t, id) => (keepIds.has(id) ? t : ""));
  for (const l of dropped) if (l) report?.push(`「${name}」の${l}は、コピーできませんでした（セルの内容はコピーしました）`);
  return xml;
}

// コピー元のExcelを読んで、シートの一覧・注意点を調べる
export async function readSource(JSZip, buf) {
  const zip = await JSZip.loadAsync(buf);
  const rels = await text(zip, "_rels/.rels");
  if (!rels) throw new Error("Excelのファイルとして読めませんでした");
  const m = rels.match(/<Relationship\b[^>]*Type="[^"]*\/officeDocument"[^>]*>/);
  const target = m && (m[0].match(/Target="([^"]+)"/) || [])[1];
  const wbPath = target ? target.replace(/^\//, "") : "xl/workbook.xml";
  const dir = wbPath.includes("/") ? wbPath.replace(/[^/]+$/, "") : "";
  const wbXml = await text(zip, wbPath);
  if (!wbXml) throw new Error("ブックの本体が見つかりませんでした");
  if (/<workbookProtection\b[^>]*workbookPassword|<fileSharing\b[^>]*reservationPassword/.test(wbXml)) throw new Error("パスワードで保護されています");
  const wbRels = (await text(zip, `${dir}_rels/${wbPath.replace(/^.*\//, "")}.rels`)) || "";
  const relMap = {};
  for (const r of wbRels.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = (r[0].match(/\bId="([^"]*)"/) || [])[1];
    const tg = (r[0].match(/\bTarget="([^"]*)"/) || [])[1];
    if (id && tg) relMap[id] = tg.startsWith("/") ? tg.slice(1) : dir + tg.replace(/^\.\//, "");
  }
  const sheets = [];
  let idx = 0;
  for (const x of wbXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const name = unesc((x[0].match(/\bname="([^"]*)"/) || [])[1] || "");
    const rid = (x[0].match(/\br:id="([^"]*)"/) || [])[1];
    const state = (x[0].match(/\bstate="([^"]*)"/) || [])[1] || "visible";
    const path = relMap[rid];
    if (!path || !/worksheets\//.test(path)) {
      idx++;
      continue; // グラフ専用シートなどは対象外
    }
    const xml = (await text(zip, path)) || "";
    sheets.push({
      name,
      idx: idx++,
      path,
      hidden: state !== "visible",
      formulas: (xml.match(/<f\b/g) || []).length,
      cells: (xml.match(/<c\b/g) || []).length,
      xml,
    });
  }
  // 他のシートを使っている数式
  for (const sh of sheets) {
    sh.refs = new Set();
    for (const f of sh.xml.matchAll(/<f\b[^>]*>([\s\S]*?)<\/f>/g)) {
      const body = unesc(f[1]);
      for (const o of sheets) {
        if (o === sh) continue;
        if (mentionsSheet(body, o.name)) sh.refs.add(o.name);
      }
      if (/\[\d+\]/.test(body) && /!/.test(body)) sh.external = true; // 他のファイルへの参照
    }
  }
  const names = [...wbXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)].map((x) => ({
    name: (x[1].match(/\bname="([^"]*)"/) || [])[1],
    local: (x[1].match(/\blocalSheetId="(\d+)"/) || [])[1],
    hidden: /\bhidden="1"/.test(x[1]),
    body: x[2],
    attrs: x[1],
  }));
  for (const sh of sheets) delete sh.xml;
  return { zip, wbPath, sheets, names, hasMacro: !!zip.file("xl/vbaProject.bin") };
}
