// PITWORK オイルエレメント・燃料エレメントの適合表を、関数だけのExcelにして書き出す（オイル・ケミカルのExcelと同じデザイン）
// 純正品番・他社品番のどれか1つを入れると、PITWORK品番と、各メーカーの品番が出る。
import { fillGuide } from "./oilxlsx.js?v=20261021h";
import { fillFinder, fillData, LK_FONT as font, LK_MAX, setLkFn } from "./lookupxlsx.js?v=20261021h";

export const PIT_MAKERS = ["日産", "トヨタ", "ダイハツ", "三菱", "スズキ", "マツダ", "ホンダ", "スバル", "いすゞ", "日野/UD", "東洋エレメント", "日東工業"];
const quote = (n) => `'${n.replace(/'/g, "''")}'`;

export const PIT_KINDS = {
  oil: { label: "オイルエレメント", find: "オイルエレメント 純正・市販品互換表", data: "オイルエレメントデータ", title: "オイルエレメント 純正・市販品互換表（普通車・中型・大型車 どちらも検索OK）", note: "普通車も、中型・大型車も、同じところで探せます。" },
  fuel: { label: "燃料エレメント", find: "燃料エレメント 純正・市販品互換表", data: "燃料エレメントデータ", title: "燃料エレメント 純正・市販品互換表（普通車・中型・大型車 どちらも検索OK）", note: "普通車も、中型・大型車も、同じところで探せます。" },
};
const GUIDE = "エレメント互換表の使い方";

const guideBlocks = (protect, nm, kinds) => [
  ["title", `${kinds.map((k) => PIT_KINDS[k].label).join("・")} 純正・市販品互換表の使い方`],
  ["sub", `メーカーの純正品番（や他社品番）から、PITWORKの品番を探すシートです。普通車も、中型・大型車も、同じところで探せます。黄色いところに品番を入れるだけ。入れたそばから、下に出ます。`],
  ["h", "まずはこれだけ"],
  ["step", "①", `${kinds.map((k) => `「${nm[k].find}」`).join("か")}のシートを開きます`],
  ["step", "②", "上の黄色いところに、品番（またはその一部）を入れます。例：15208-65F0E、1520865F0E"],
  ["b", "ハイフン・半角全角・大文字小文字は、気にしなくてOKです"],
  ["b", "PITWORKの品番（例：AY100-NS001）を入れても、各メーカーの品番を逆に調べられます"],
  ["b", "候補の並び順：品番がぴったり合うもの → 先頭が合うもの → 途中が合うもの"],
  ["b", "言葉を消すときは、黄色いところを選んで Delete キーを押します"],
  ["b", "品番を何も入れていないときは、候補は出ません（動きを軽くするため）"],
  ["h", "見かた"],
  ["b", "1行が、PITWORKの品番1つです。その行に、同じ部品として使える各メーカーの品番が並びます"],
  ["b", "1つのメーカーに品番が複数あるときは、そのマスの中に縦に並びます（7件以上は、6件まで出して「…ほか○件」と出ます。全部は、データのシートで見られます）"],
  ["b", "品番のうしろの（※1）などは、カタログの注記です。日野（※1）／UDトラックス（※2）は、同じ列に入っています。中・大型車のカタログは、品番のうしろに（日野）（UD）と付けています"],
  ["h", "データを入れる（更新する）"],
  ["step", "①", `${kinds.map((k) => `「${nm[k].data}」`).join("・")}のシートに、PITWORK品番・種別・各メーカーの品番を貼ります（見出しは1行目）`],
  ["step", "②", "1つのマスに品番が複数あるときは、「 / 」（半角スペース・スラッシュ・半角スペース）で区切ります"],
  ["step", "③", "右の灰色の列は、数式です。消さないでください。貼り付けるだけで、検索に反映されます"],
  ["b", `${LK_MAX.toLocaleString()}行目まで見ます`],
  ["h", "困ったとき"],
  ["b", "何も出ない・動かない → 画面上の黄色い帯の「編集を有効にする」を押してください"],
  ["b", "探しても出ない → 品番を短くして試してください（例：「15208-65F0E」→「65F0E」）。それでも出ないときは、カタログに載っていない品番かもしれません"],
  ["h", "これだけは守ってください"],
  ["b", "このデータは、カタログの写真から書き写したものです。注文の前に、必ずカタログで確認してください"],
  ["b", "元のファイルは触らずに、各自コピーして使ってください"],
  ...(protect
    ? [
        ["b", "数式を間違って消さないように、検索のシートは保護してあります。入力できるのは黄色い欄だけです（データのシートは、保護していません）"],
        ["b", "保護を外したいときは「校閲 → シート保護の解除」です（パスワードなし）"],
      ]
    : []),
];

// opt: { protect, names?, dbcsName?, rows:[[pit, series, ...12メーカーの品番（" / "区切り）]] }
// opt: { protect, names?: {oil:{find,data}, fuel:{find,data}, guide}, rows: {oil:[...], fuel:[...]}, kinds?: ["oil","fuel"] }
export function buildPitWorkbook(ExcelJS, opt = {}) {
  setLkFn("ASC"); // 半角にそろえる（長い文字列でも使える）
  const kinds = opt.kinds?.length ? opt.kinds : ["oil", "fuel"];
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = { guide: opt.names?.guide || GUIDE };
  for (const k of kinds) nm[k] = { find: opt.names?.[k]?.find || PIT_KINDS[k].find, data: opt.names?.[k]?.data || PIT_KINDS[k].data };
  const protect = !!opt.protect;
  const list = [];
  const wsF = {};
  const wsD = {};
  for (const k of kinds) {
    wsF[k] = wb.addWorksheet(nm[k].find, { views: [{ showGridLines: false, tabSelected: list.length === 0 }] });
    list.push(nm[k].find);
  }
  for (const k of kinds) {
    wsD[k] = wb.addWorksheet(nm[k].data);
    list.push(nm[k].data);
  }
  const wsG = wb.addWorksheet(nm.guide, { views: [{ showGridLines: false }] });
  list.push(nm.guide);
  const M = PIT_MAKERS.length; // 12
  const cols = [{ head: "PITWORK品番", col: "A", width: 15, mono: true, center: true }, ...PIT_MAKERS.map((m, i) => ({ head: m, col: String.fromCharCode(67 + i), width: 15, nl: true, cap: 6, big: 12 }))]; // A 品番, B 種別, C〜N メーカー
  // 検索用：全品番を「|」でつなぎ、ハイフン・空白・注記をとる
  const raw = `$A{r}&" / "&${Array.from({ length: M }, (_, i) => `$${String.fromCharCode(67 + i)}{r}`).join('&" / "&')}`;
  const strip = (x) => ["(※1)", "(※2)", "(※3)", "(日野)", "(UD)"].reduce((a, t) => `SUBSTITUTE(${a},"${t}","")`, x);
  const hay = (r) => `IF($A${r}="","","|"&${["－", "‐", "−", "ー", "―", "–", "—", "‑"].reduce((a, h) => `SUBSTITUTE(${a},"${h}","")`, `SUBSTITUTE(LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${strip(raw.replaceAll("{r}", r))}," / ","|"),"-","")," ","")),"　","")`)}&"|")`;
  for (const k of kinds) {
    const F = nm[k].find;
    const S = (c) => `${quote(F)}!$${c}$3`;
    fillFinder(wsF[k], { guide: nm.guide }, { title: PIT_KINDS[k].title, label1: "品番（純正・他社・PITWORK）", hint1: "どのメーカーの品番でもOK。ハイフンなし・一部だけでもOK", data: nm[k].data, rankCol: "S", cols, rowH: 108 }, protect);
    const score = (r) => `IF(AND(${S("W")}<>"",ISNUMBER(SEARCH("|"&${S("W")}&"|",$P${r}))),4,IF(AND(${S("W")}<>"",ISNUMBER(SEARCH("|"&${S("W")},$P${r}))),3,2))`;
    fillData(wsD[k], nm[k], { inHeads: ["PITWORK品番", "種別", ...PIT_MAKERS], inWidths: [15, 9, ...PIT_MAKERS.map(() => 22)], hayCol: "P", flagCol: "Q", scoreCol: "R", rankCol: "S", finder: F, hay, score, sample: false }, opt.rows?.[k] || []);
  }
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm, kinds) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list, names: nm, kinds };
}
