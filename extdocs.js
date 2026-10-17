// 外装検索：部位（フロントバンパーなど）の解説と、「車のどこにあるか」を示す模式図
// 部品の名前・図番（データ）は、Firebaseに入れる。ここには、部位の説明・別名・絵だけを置く（公開されても困らない内容）

const esc = (s) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

// ---- 車の模式図（実物の車の形ではなく、位置を示すための絵） ----
const BODY = 'fill="#8a8f98" fill-opacity=".16" stroke="#8a8f98" stroke-width="2" stroke-linejoin="round"';
const GLS = 'fill="#cfe8ff" fill-opacity=".22" stroke="#8a8f98" stroke-width="1.6" stroke-linejoin="round"';
const DARK = 'fill="#17191c" fill-opacity=".55" stroke="#8a8f98" stroke-width="1.4" stroke-linejoin="round"';
const TIRE = 'fill="#17191c" fill-opacity=".8" stroke="#8a8f98" stroke-width="2"';
const HL = 'fill="#f2c230" fill-opacity=".9" stroke="#b08400" stroke-width="2.2" stroke-linejoin="round"';
const HLD = 'fill="#f2c230" fill-opacity=".28" stroke="#d9a400" stroke-width="2.4" stroke-dasharray="6 4" stroke-linejoin="round"';
const part = (hl, name, d, normal, high = HL) => `<path data-r="${name}" d="${d}" ${hl === name ? high : normal}/>`;

// 正面から。hl は、色をつける部位の名前
export function carFront(hl = "", click = false) {
  const on = (n) => (hl === n || (hl === "fenders" && /^fender/.test(n)) || (hl === "hoodLedges" && /^hoodLedge/.test(n)) ? n : "");
  const R = (n, d, normal, high) => part(on(n) || (hl === n ? n : ""), n, d, normal, high);
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 196" class="ex-svg" role="img" aria-label="車を正面から見た模式図">
    <rect x="32" y="146" width="26" height="42" rx="7" ${TIRE}/><rect x="262" y="146" width="26" height="42" rx="7" ${TIRE}/>
    <path d="M44 150 L40 120 Q42 102 60 92 L78 54 Q82 40 98 38 L222 38 Q238 40 242 54 L260 92 Q278 102 280 120 L276 150 Z" ${BODY}/>
    <path d="M92 52 L228 52 L246 88 L74 88 Z" ${GLS}/>
    ${R("hood", "M74 88 L246 88 L262 114 L58 114 Z", `fill="#8a8f98" fill-opacity=".1" stroke="#8a8f98" stroke-width="1.6" stroke-linejoin="round"`)}
    ${R("coreSupport", "M68 112 L252 112 L252 156 L68 156 Z", `fill="none" stroke="none"`, HLD)}
    ${R("hoodLedgeL", "M64 90 L86 88 L92 114 L66 116 Z", `fill="none" stroke="none"`, HLD)}
    ${R("hoodLedgeR", "M256 90 L234 88 L228 114 L254 116 Z", `fill="none" stroke="none"`, HLD)}
    ${R("hoodLock", "M146 104 L174 104 L174 118 L146 118 Z", `fill="none" stroke="none"`, HLD)}
    ${R("dash", "M88 90 L232 90 L238 108 L82 108 Z", `fill="none" stroke="none"`, HLD)}
    ${R("headL", "M60 116 L114 116 L110 134 L56 130 Z", DARK)}
    ${R("headR", "M260 116 L206 116 L210 134 L264 130 Z", DARK)}
    ${R("grille", "M120 116 L200 116 L196 140 L124 140 Z", DARK, HL)}
    ${R("bumper", "M50 140 L270 140 L276 172 Q276 180 266 180 L54 180 Q44 180 44 172 Z", `fill="#8a8f98" fill-opacity=".22" stroke="#8a8f98" stroke-width="2" stroke-linejoin="round"`)}
    <path d="M118 152 L202 152 L196 168 L124 168 Z" ${DARK}/>
    <path d="M62 152 L96 152 L92 168 L66 168 Z" ${DARK}/><path d="M258 152 L224 152 L228 168 L254 168 Z" ${DARK}/>
    ${R("fenderL", "M44 128 Q44 106 62 96 L76 88 L86 92 L76 130 L58 132 Z", `fill="#8a8f98" fill-opacity=".1" stroke="#8a8f98" stroke-width="1.6" stroke-linejoin="round"`)}
    ${R("fenderR", "M276 128 Q276 106 258 96 L244 88 L234 92 L244 130 L262 132 Z", `fill="#8a8f98" fill-opacity=".1" stroke="#8a8f98" stroke-width="1.6" stroke-linejoin="round"`)}
    ${hl === "fenderL" || hl === "fenderR" ? "" : ""}
  </svg>`;
}

// 横から（前が左）
export function carSide(hl = "") {
  const R = (n, d, normal, high) => part(hl, n, d, normal, high);
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 360 176" class="ex-svg" role="img" aria-label="車を横から見た模式図">
    <path d="M12 124 L14 108 Q16 100 28 98 L100 86 L132 52 Q136 46 146 46 L238 46 Q248 46 254 54 L286 84 L326 96 Q340 100 342 112 L344 124 Q344 134 334 134 L22 134 Q12 134 12 124 Z" ${BODY}/>
    <path d="M136 56 L232 56 L262 84 L114 84 Z" ${GLS}/>
    <path d="M168 56 L168 134 M226 58 L226 134" fill="none" stroke="#8a8f98" stroke-width="1.4"/>
    ${R("fender", "M26 100 L114 86 L116 134 L102 134 A32 32 0 0 0 38 134 L26 134 Z", `fill="#8a8f98" fill-opacity=".1" stroke="#8a8f98" stroke-width="1.4" stroke-linejoin="round"`)}
    ${R("hood", "M24 98 L98 83 L116 83 L114 88 L26 104 Z", `fill="none" stroke="none"`, HL)}
    ${R("hoodLedgeS", "M44 98 L108 88 L108 100 L46 110 Z", `fill="none" stroke="none"`, HLD)}
    ${R("dashS", "M104 82 L126 82 L126 134 L104 134 Z", `fill="none" stroke="none"`, HLD)}
    ${R("bumperF", "M12 112 L28 108 L28 134 L22 134 Q12 134 12 124 Z", `fill="#8a8f98" fill-opacity=".25" stroke="#8a8f98" stroke-width="1.6" stroke-linejoin="round"`)}
    ${R("coreSupportS", "M30 104 L52 102 L52 130 L30 130 Z", `fill="none" stroke="none"`, HLD)}
    <circle cx="70" cy="134" r="33" fill="none" stroke="#8a8f98" stroke-width="1.6"/>
    <circle cx="282" cy="134" r="33" fill="none" stroke="#8a8f98" stroke-width="1.6"/>
    <circle cx="70" cy="134" r="25" ${TIRE}/><circle cx="70" cy="134" r="11" fill="#cfd3d9" fill-opacity=".7"/>
    <circle cx="282" cy="134" r="25" ${TIRE}/><circle cx="282" cy="134" r="11" fill="#cfd3d9" fill-opacity=".7"/>
    <text x="22" y="164" font-size="12" fill="#8a8f98">← 前</text><text x="318" y="164" font-size="12" fill="#8a8f98">後ろ →</text>
  </svg>`;
}

// ---- 部位の設定（名前・別名・説明・どの絵のどこにあるか） ----
// name は、データの「部位」の名前とそろえる
export const EXT_GROUPS = [
  {
    name: "フロントバンパー",
    en: "FRONT BUMPER",
    area: "フロント",
    figs: [["front", "bumper", "正面から"], ["side", "bumperF", "横から"]],
    hot: ["front:bumper"],
    one: "車の一番前の下にある部分です。外側のカバー（バンパー本体）の内側に、芯材と衝撃吸収材が入っています。",
    alias: ["バンパー", "ばんぱー", "フロント バンパー", "前バンパー", "ファシア", "芯材", "リインフォース", "アブソーバー", "衝撃吸収", "ステー", "ブラケット", "ナンバー", "ナンバープレート", "モール", "フィニッシャー"],
  },
  {
    name: "フロントグリル",
    en: "FRONT GRILLE",
    area: "フロント",
    figs: [["front", "grille", "正面から"]],
    hot: ["front:grille"],
    one: "車の正面の、真ん中にある格子状の部分です。メーカーのエンブレムが付きます。",
    alias: ["グリル", "ぐりる", "ラジエーターグリル", "ラジエターグリル", "フロントマスク", "エンブレム"],
  },
  {
    name: "ラジエーターコアサポート",
    en: "FRONT APRON & RADIATOR CORE SUPPORT",
    area: "フロント",
    figs: [["front", "coreSupport", "正面から（奥にある骨組み）"], ["side", "coreSupportS", "横から（奥にある骨組み）"]],
    hot: ["front:coreSupport"],
    one: "ボンネットの下、バンパーの奥にある骨組みです。ラジエーターやライトを支えます。外からは見えません。",
    alias: ["コアサポート", "こあさぽーと", "ラジエターサポート", "ラジエータサポート", "ラジエーターサポート", "フロントパネル", "ボンネットロック", "フードロック", "骨組み"],
  },
  {
    name: "フロントフェンダー",
    en: "FRONT FENDER & FITTING",
    area: "フロント",
    figs: [["side", "fender", "横から"], ["front", "fenders", "正面から（左右の端）"]],
    hot: ["front:fenderL", "front:fenderR"],
    one: "前のタイヤの上にある、ボディの外側の板です。内側に、樹脂のプロテクター（ライナー）が付きます。",
    alias: ["フェンダー", "ふぇんだー", "前フェンダー", "ウイング", "プロテクター", "ライナー", "インナーフェンダー", "フェンダーライナー", "ボンネットシール"],
  },
];
EXT_GROUPS.push(
  {
    name: "ボンネット",
    en: "HOOD PANEL, HINGE & FITTING",
    area: "フロント",
    figs: [["front", "hood", "正面から"], ["side", "hood", "横から"]],
    hot: ["front:hood"],
    one: "エンジンルームを上から覆うフタです。ヒンジ（開閉の軸）と、支え棒（ステー）で支え、ゴムのシールやバンパー（ゴムの当たり）が付きます。",
    alias: ["フード", "ふーど", "ボンネットヒンジ", "ヒンジ", "ステー", "支え棒", "ロッド", "インシュレーター", "防音", "シール", "ゴム", "バンパーゴム"],
  },
  {
    name: "ボンネットロック",
    en: "HOOD LOCK CONTROL",
    area: "フロント",
    figs: [["front", "hoodLock", "正面から（奥にあるロック）"]],
    hot: ["front:hoodLock"],
    one: "ボンネットを閉めたまま固定する部分です。車内のハンドルを引くと、ケーブルでロックが外れます。",
    alias: ["フードロック", "ロック", "ボンネットオープナー", "オープナー", "ケーブル", "ハンドル", "ベルクランク", "開閉"],
  },
  {
    name: "フードレッジ",
    en: "HOOD LEDGE & FITTING; MAIN PARTS",
    area: "フロント",
    figs: [["front", "hoodLedges", "正面から（左右の奥の骨組み）"], ["side", "hoodLedgeS", "横から（奥の骨組み）"]],
    hot: ["front:hoodLedgeL", "front:hoodLedgeR"],
    one: "ボンネットの下、エンジンルームの左右の壁にあたる骨組みです。ボンネットとフェンダーなどを支えます。外からは見えません。",
    alias: ["フードリッジ", "ふーどれっじ", "エプロン", "フロントエプロン", "ストラット", "ストラットハウジング", "サスペンションハウジング", "補強", "レインフォース", "骨組み", "スプラッシュカバー"],
  },
  {
    name: "ダッシュパネル",
    en: "DASH PANEL & FITTING",
    area: "フロント",
    figs: [["front", "dash", "正面から（奥にある壁）"], ["side", "dashS", "横から（奥にある壁）"]],
    hot: [],
    one: "エンジンルームと車内の間にある、仕切りの壁です。外からは見えません。",
    alias: ["ダッシュ", "だっしゅ", "隔壁", "防火壁", "カウル", "カウルトップ", "ダッシュサイド", "インシュレーター", "防音"],
  },
  {
    name: "ダッシュトリム",
    en: "DASH TRIMMING & FITTING",
    area: "フロント",
    figs: [["side", "dashS", "横から（車内側の壁）"]],
    hot: [],
    one: "ダッシュパネル（仕切りの壁）の、車内側につく防音材と、足元まわりの飾り板です。",
    alias: ["ダッシュ", "だっしゅ", "インシュレーター", "防音", "フィニッシャー", "サイドフィニッシャー", "足元"],
  },
);
export const AREAS = ["フロント", "サイド", "リア"];

// 一覧に載せない小物（探すときの言葉）。入力されたら、部位のページへ案内する
export const SMALL_RE = /クリップ|くりっぷ|ボルト|ぼると|ナット|なっと|ネジ|ねじ|ビス|ぴす|スクリュー|グロメット|プラグ|とめる|止める|留める|止め|留め|固定/g;

const figOf = (view, hl) => (view === "side" ? carSide(hl) : carFront(hl));
export function groupFigs(meta) {
  return (meta?.figs || []).map(([view, hl, cap]) => ({ svg: figOf(view, hl), cap }));
}
// 「絵から探す」：正面図の、押せる部位
export function pickerSvg() {
  const svg = carFront("");
  return svg;
}
export const metaOf = (name) => EXT_GROUPS.find((g) => g.name === name);
export { esc };
