// 部品検索（内装・機関）：部位の解説・別名・絵（エンジンなどの模式図）
// 部品の名前・図番（データ）は Firebase に入れる。ここには、部位の説明・別名・絵だけを置く

const BODY = 'fill="#8a8f98" fill-opacity=".16" stroke="#8a8f98" stroke-width="2" stroke-linejoin="round"';
const HL = 'fill="#f2c230" fill-opacity=".9" stroke="#b08400" stroke-width="2.2" stroke-linejoin="round"';

// エンジンを横から見た、簡単な模式図。hl の部位に色をつける
export function engineSide(hl = []) {
  const on = (n) => (hl.includes(n) ? HL : BODY);
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 196" class="ex-svg" role="img" aria-label="エンジンを横から見た模式図">
    <path data-r="cover" d="M82 36 L238 36 L244 52 L76 52 Z" ${on("cover")}/>
    <path data-r="head" d="M72 52 L248 52 L252 90 L68 90 Z" ${on("head")}/>
    <path data-r="block" d="M68 90 L252 90 L256 150 L64 150 Z" ${on("block")}/>
    <path data-r="pan" d="M76 150 L244 150 L228 182 L92 182 Z" ${on("pan")}/>
    <path d="M110 52 L110 90 M160 52 L160 90 M210 52 L210 90 M110 90 L110 150 M160 90 L160 150 M210 90 L210 150" fill="none" stroke="#8a8f98" stroke-width="1" stroke-opacity=".5"/>
    <circle cx="160" cy="118" r="20" fill="none" stroke="#8a8f98" stroke-width="1.6" stroke-opacity=".6"/>
    <text x="14" y="190" font-size="11" fill="#8a8f98">←前</text>
  </svg>`;
}

// name は、データの「部位」の名前とそろえる
export const INN_GROUPS = [
  {
    name: "シリンダーヘッド・ロッカーカバー",
    en: "CYLINDER HEAD & ROCKER COVER; ASSEMBLY",
    area: "エンジン",
    figs: [["engine", ["head", "cover"], "横から（エンジンの上の部分）"]],
    hot: [],
    one: "エンジンの上の部分です。シリンダーヘッドは、燃焼室のフタで、バルブやカムシャフトが入ります。いちばん上のロッカーカバーが、それを覆い、間にガスケット（パッキン）が入ります。",
    alias: ["ヘッドガスケット", "ヘッドパッキン", "ヘッドパッキン", "ガスケット", "ロッカーカバー", "タペットカバー", "ヘッドカバー", "シリンダーヘッド", "ヘッド", "オイルフィラーキャップ", "オイル給油口", "給油口キャップ", "PCVバルブ", "ブローバイ", "ブローバイバルブ", "エンジン吊り", "オイル漏れ", "にじみ"],
  },
];
INN_GROUPS.push({
  name: "シリンダーブロック・オイルパン",
  en: "CYLINDER BLOCK & OIL PAN; ASSEMBLY / COMPONENT PARTS",
  area: "エンジン",
  figs: [["engine", ["block", "pan"], "横から（エンジンの下の部分）"]],
  hot: [],
  one: "エンジンの本体（シリンダーブロック）と、いちばん下のオイルパンです。ブロックの中でピストンが動き、オイルパンにエンジンオイルがたまります。オイルを抜くドレンプラグは、オイルパンの底にあります。",
  alias: ["オイルパン", "ドレンプラグ", "ドレンボルト", "ドレンパッキン", "ドレンワッシャー", "オイル交換", "オイルレベルゲージ", "オイルゲージ", "レベルゲージ", "ディップスティック", "バランサー", "ブロック", "エンジンブロック", "シリンダーブロック", "リアオイルシール", "クランクシール", "クランクシャフトシール", "オイルジェット", "液体ガスケット", "バッフルプレート", "リアプレートカバー", "ブロックガスケット", "オイル漏れ", "オイルにじみ"],
});
INN_GROUPS[0].alias.push("バルブガイド", "バルブシート", "プラグチューブ", "スパークプラグチューブ", "プラグホール", "点火プラグ", "ヘッドの部品", "バルブ");
export const INN_AREAS = ["エンジン", "ミッション", "足回り", "ブレーキ", "電装", "空調", "内装"];

export const metaOfInn = (name) => INN_GROUPS.find((g) => g.name.normalize("NFKC") === String(name).normalize("NFKC"));
export function groupFigsInn(meta) {
  return (meta?.figs || []).map(([view, hl, cap]) => ({ svg: engineSide(Array.isArray(hl) ? hl : [hl]), cap }));
}
