// PITWORK オイルエレメント・燃料エレメントの適合表を、関数だけのExcelにして書き出す（オイル・ケミカルのExcelと同じデザイン）
// 純正品番・他社品番のどれか1つを入れると、PITWORK品番と、各メーカーの品番が出る。
import { fillGuide } from "./oilxlsx.js?v=20261023a";
import { fillFinder, fillData, LK_FONT as font, LK_MAX, setLkFn } from "./lookupxlsx.js?v=20261023a";

export const PIT_MAKERS = ["日産", "トヨタ", "ダイハツ", "三菱", "スズキ", "マツダ", "ホンダ", "スバル", "いすゞ", "日野/UD", "東洋エレメント", "日東工業"];
const quote = (n) => `'${n.replace(/'/g, "''")}'`;

export const PIT_NAMES = { find: "オイル・燃料エレメント 純正・市販品互換表", data: "オイル・燃料エレメントデータ", guide: "エレメント互換表の使い方" };
const PIT_TITLE = "オイル・燃料エレメント 純正・市販品互換表(普通車・中型・大型車 どちらも検索OK)";

const guideBlocks = (protect, nm) => [
  ["title", "オイル・燃料エレメント 純正・市販品互換表の使い方"],
  ["sub", "メーカーの純正品番（や他社品番）から、PITWORKの品番を探すシートです。オイルエレメントも、燃料エレメントも、普通車も、中型・大型車も、同じところで探せます。黄色いところに品番を入れるだけ。入れたそばから、下に出ます。"],
  ["h", "まずはこれだけ"],
  ["step", "①", `「${nm.find}」のシートを開きます`],
  ["step", "②", "上の黄色いところに、品番（またはその一部）を入れます。例：15208-65F0E、1520865F0E"],
  ["b", "ハイフン・半角全角・大文字小文字は、気にしなくてOKです"],
  ["b", "PITWORKの品番（例：AY100-KE001）を入れても、各メーカーの品番を逆に調べられます"],
  ["b", "候補の並び順：品番がぴったり合うもの → 先頭が合うもの → 途中が合うもの"],
  ["b", "言葉を消すときは、黄色いところを選んで Delete キーを押します"],
  ["b", "品番を何も入れていないときは、候補は出ません（動きを軽くするため）"],
  ["h", "見かた"],
  ["b", "1行が、PITWORKの品番1つです。品番の下の（　）は、種類です（オイル・普通車／オイル・中大型／燃料・普通車／燃料・中大型）"],
  ["b", "その行に、同じ部品として使える各メーカーの品番が並びます。1つのメーカーに品番が複数あるときは、そのマスの中に縦に並びます（7件以上は、6件まで出して「…ほか○件」と出ます。全部は、データのシートで見られます）"],
  ["b", "品番のうしろの（※1）などは、カタログの注記です。日野（※1）／UDトラックス（※2）は、同じ列に入っています。中・大型車のカタログは、品番のうしろに（日野）（UD）と付けています"],
  ["h", "データを入れる（更新する）"],
  ["step", "①", `「${nm.data}」シートに、PITWORK品番・種類・各メーカーの品番を貼ります（見出しは1行目）`],
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

// opt: { protect, names?: {find, data, guide}, rows: [[pit, 種類, ...12メーカーの品番（" / "区切り）]] }
export function buildPitWorkbook(ExcelJS, opt = {}) {
  setLkFn("ASC"); // 半角にそろえる（長い文字列でも使える）
  const wb = new ExcelJS.Workbook();
  wb.creator = "セクション検索";
  if (!opt.names) wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const nm = { find: opt.names?.find || PIT_NAMES.find, data: opt.names?.data || PIT_NAMES.data, guide: opt.names?.guide || PIT_NAMES.guide };
  const protect = !!opt.protect;
  const wsF = wb.addWorksheet(nm.find, { views: [{ showGridLines: false, tabSelected: true }] });
  const wsD = wb.addWorksheet(nm.data);
  const wsG = wb.addWorksheet(nm.guide, { views: [{ showGridLines: false }] });
  const list = [nm.find, nm.data, nm.guide];
  const M = PIT_MAKERS.length; // 12
  const cols = [{ head: "PITWORK品番", col: "A", tag: "B", width: 15, mono: true, center: true }, ...PIT_MAKERS.map((m, i) => ({ head: m, col: String.fromCharCode(67 + i), width: 15, nl: true, cap: 6, big: 12 }))]; // A 品番, B 種類, C〜N メーカー
  fillFinder(wsF, nm, { title: PIT_TITLE, label1: "品番（純正・他社・PITWORK）", hint1: "どのメーカーの品番でもOK。ハイフンなし・一部だけでもOK", data: nm.data, rankCol: "S", cols, rowH: 108 }, protect);
  // 検索用：全品番を「|」でつなぎ、ハイフン・空白・注記をとる
  const raw = `$A{r}&" / "&${Array.from({ length: M }, (_, i) => `$${String.fromCharCode(67 + i)}{r}`).join('&" / "&')}`;
  const strip = (x) => ["(※1)", "(※2)", "(※3)", "(日野)", "(UD)"].reduce((a, t) => `SUBSTITUTE(${a},"${t}","")`, x);
  const hay = (r) => `IF($A${r}="","","|"&${["－", "‐", "−", "ー", "―", "–", "—", "‑"].reduce((a, h) => `SUBSTITUTE(${a},"${h}","")`, `SUBSTITUTE(LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${strip(raw.replaceAll("{r}", r))}," / ","|"),"-","")," ","")),"　","")`)}&"|")`;
  const S = (c) => `${quote(nm.find)}!$${c}$3`;
  const score = (r) => `IF(AND(${S("W")}<>"",ISNUMBER(SEARCH("|"&${S("W")}&"|",$P${r}))),4,IF(AND(${S("W")}<>"",ISNUMBER(SEARCH("|"&${S("W")},$P${r}))),3,2))`;
  fillData(wsD, nm, { inHeads: ["PITWORK品番", "種類", ...PIT_MAKERS], inWidths: [15, 14, ...PIT_MAKERS.map(() => 22)], hayCol: "P", flagCol: "Q", scoreCol: "R", rankCol: "S", finder: nm.find, hay, score, sample: false }, opt.rows || []);
  fillGuide(wsG, { font, protect, K: {}, nameP: "", blocks: guideBlocks(protect, nm) });
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];
  return { wb, list, names: nm, kinds: ["pit"] };
}
