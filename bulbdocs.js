// バルブの解説書：形状ごとに、使う場所・見分け方・接客のポイント・豆知識をまとめる
// 初心者が見ながら接客できる／学習用。店のバルブデータ（品番・規格・用途）を、形状ごとにつなげて表示する
// 書いてあることは、メーカーの資料や規格の説明で確かめられたことと、店のデータにあることだけ。分からないことは書かない

const esc = (s) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

// ---- イメージ図（模式図。形のイメージで、実物とは細部が違います） ----
const GLASS = 'fill="#cfe8ff" fill-opacity=".38" stroke="#8fb4d9" stroke-width="2"';
const METAL = 'fill="#c9ccd1" stroke="#6b7078" stroke-width="2"';
const FIL = 'fill="none" stroke="#f2c230" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round"';
const svg = (inner, w = 140, h = 150) =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-hidden="true" class="bg-svg">${inner}</svg>`;
const coil = (x, y, w, n = 5) => {
  let d = `M${x} ${y}`;
  for (let i = 0; i < n; i++) d += ` l${w / (2 * n)} -7 l${w / (2 * n)} 7`;
  return `<path d="${d}" ${FIL}/>`;
};

export const SVGS = {
  // ウェッジ球（ガラスの足で差し込む）。size は太さ
  wedge: (size = 28) => {
    const x = 70 - size / 2;
    return svg(
      `<rect x="${x}" y="14" width="${size}" height="76" rx="${size / 2}" ${GLASS}/>` +
        `<path d="M${x + 4} 88 H${x + size - 4} V118 H${x + 4} Z" ${GLASS}/>` +
        `<path d="M${70 - 4} 40 V88" ${FIL}/>${coil(70 - size / 4, 52, size / 2, 3)}` +
        `<rect x="${x + 5}" y="94" width="5" height="40" rx="1.5" ${METAL}/>` +
        `<rect x="${x + size - 10}" y="94" width="5" height="40" rx="1.5" ${METAL}/>`,
    );
  },
  // バヨネット球（金属の筒を差し込んで回す）。double は、ピンの高さが左右でちがう
  bayonet: (double = false, r = 30) =>
    svg(
      `<circle cx="70" cy="${22 + r}" r="${r}" ${GLASS}/>` +
        (double
          ? `<path d="M${70 - 14} ${22 + r + 2} l6 -12 l6 12 M${70 + 2} ${22 + r + 2} l6 -12 l6 12" ${FIL}/>`
          : `<path d="M${70 - 12} ${22 + r + 2} l6 -14 l6 14 l6 -14 l6 14" ${FIL}/>`) +
        `<rect x="48" y="${22 + 2 * r - 6}" width="44" height="${double ? 44 : 40}" rx="3" ${METAL}/>` +
        (double
          ? `<rect x="40" y="${22 + 2 * r + 6}" width="10" height="6" rx="2" ${METAL}/><rect x="90" y="${22 + 2 * r + 20}" width="10" height="6" rx="2" ${METAL}/>`
          : `<rect x="40" y="${22 + 2 * r + 12}" width="10" height="6" rx="2" ${METAL}/><rect x="90" y="${22 + 2 * r + 12}" width="10" height="6" rx="2" ${METAL}/>`) +
        (double
          ? `<circle cx="62" cy="${22 + 2 * r + 38}" r="3.5" fill="#6b7078"/><circle cx="78" cy="${22 + 2 * r + 38}" r="3.5" fill="#6b7078"/>`
          : `<circle cx="70" cy="${22 + 2 * r + 36}" r="3.5" fill="#6b7078"/>`),
    ),
  // 両口金（細長いガラス管の、両端に金具）
  festoon: () =>
    svg(
      `<rect x="30" y="56" width="80" height="26" rx="13" ${GLASS}/>` +
        `<path d="M44 69 l7 -8 l7 8 l7 -8 l7 8 l7 -8 l7 8" ${FIL}/>` +
        `<rect x="14" y="58" width="20" height="22" rx="3" ${METAL}/><rect x="106" y="58" width="20" height="22" rx="3" ${METAL}/>`,
    ),
  // ハロゲン（小さなガラスの筒＋金属のつば＋端子）。tabs は、つばの爪の数
  halogen: (tabs = 2) => {
    const t = [];
    for (let i = 0; i < tabs; i++) {
      const a = (-90 + (360 / tabs) * i + (tabs === 3 ? 0 : 90)) * (Math.PI / 180);
      t.push(`<rect x="${70 + Math.cos(a) * 34 - 5}" y="${96 + Math.sin(a) * 9 - 3}" width="10" height="7" rx="2" ${METAL}/>`);
    }
    return svg(
      `<rect x="52" y="16" width="36" height="66" rx="16" ${GLASS}/>` +
        `<path d="M62 62 l5 -10 l5 10 l5 -10 l5 10" ${FIL}/><path d="M60 40 l5 -10 l5 10 l5 -10 l5 10" ${FIL}/>` +
        `<ellipse cx="70" cy="96" rx="34" ry="9" ${METAL}/>${t.join("")}` +
        `<rect x="56" y="102" width="28" height="18" rx="3" ${METAL}/>` +
        `<rect x="58" y="120" width="5" height="14" ${METAL}/><rect x="67.5" y="120" width="5" height="14" ${METAL}/><rect x="77" y="120" width="5" height="14" ${METAL}/>`,
    );
  },
  // L字の口金の仲間（H8・H11・H16など）：樹脂の台に小さなガラス球
  halogenL: () =>
    svg(
      `<rect x="52" y="14" width="36" height="58" rx="16" ${GLASS}/>` +
        `<path d="M62 54 l5 -10 l5 10 l5 -10 l5 10" ${FIL}/>` +
        `<ellipse cx="70" cy="82" rx="30" ry="8" ${METAL}/>` +
        `<path d="M46 86 H94 V104 H80 V126 H60 V104 H46 Z" fill="#8a8f98" stroke="#555a62" stroke-width="2"/>` +
        `<rect x="48" y="90" width="9" height="6" rx="1.5" fill="#c9ccd1"/><rect x="83" y="90" width="9" height="6" rx="1.5" fill="#c9ccd1"/>`,
    ),
  // コード（リード線）つき（H3）
  halogenLead: () =>
    svg(
      `<rect x="40" y="18" width="30" height="62" rx="14" ${GLASS}/>` +
        `<path d="M48 62 l4 -9 l4 9 l4 -9 l4 9" ${FIL}/>` +
        `<ellipse cx="55" cy="90" rx="28" ry="8" ${METAL}/><rect x="42" y="94" width="26" height="14" rx="3" ${METAL}/>` +
        `<path d="M68 100 C 100 100, 96 120, 108 128" fill="none" stroke="#6b7078" stroke-width="3"/>` +
        `<rect x="104" y="124" width="16" height="10" rx="2" ${METAL}/>`,
    ),
  // HID：細長い放電管。shade=true は、ガラスの上側に黒いシェード（D2R・D4R）
  hid: (shade = false) =>
    svg(
      `<rect x="52" y="12" width="36" height="86" rx="17" ${GLASS}/>` +
        `<circle cx="70" cy="48" r="9" fill="#e9f3ff" stroke="#8fb4d9" stroke-width="2"/>` +
        `<path d="M64 48 h12" stroke="#f2c230" stroke-width="3.2" stroke-linecap="round"/>` +
        (shade ? `<path d="M52 30 a17 17 0 0 1 36 0 V43 H52 Z" fill="#25272b" stroke="#aab0b8" stroke-width="1.6"/>` : "") +
        `<rect x="46" y="98" width="48" height="24" rx="4" ${METAL}/>` +
        `<rect x="60" y="122" width="20" height="12" rx="2" ${METAL}/>`,
    ),
};

// ---- 解説の中身 ----
// keys：データの「形状」にあたるもの（大文字・空白なしにそろえた形）。
const HALO_CAUTION = "ガラスの部分は、素手で触らないでください（手の油がつくと、そこだけ熱で傷みやすくなります）。触ったら、きれいな布で拭いてから使います。";
const FIT_TIP = "車に合うかどうかは、「車種・年式」から適合表で確認するのがいちばん確実です。元のバルブに書いてある数字（例：12V 55W）を見せてもらうのも有効です。";

export const GROUPS = [
  ["head", "ヘッドライト（ハロゲン）"],
  ["fog", "フォグランプ（ハロゲン）"],
  ["hid", "HID（キセノン）"],
  ["wedge", "ウェッジ球（ガラスの足で差し込む）"],
  ["bay", "バヨネット球（金属の筒を回して付ける）"],
  ["fest", "両口金"],
];

export const GUIDE = [
  {
    id: "h4",
    group: "head",
    title: "H4（HB2・H4uを含む）",
    keys: ["H4", "HB2(H4)", "H4U"],
    img: () => SVGS.halogen(3),
    cap: "ヘッドライト用ハロゲン。爪が3つある口金の形のイメージ",
    one: "1つのバルブで、ハイビームとロービームの両方を点けるヘッドライト用バルブです。",
    where: ["ヘッドライト（ハイ・ロー兼用）", "12Vは60/55W、24V（トラックなど）は75/70W（店のデータより）"],
    spot: [
      "バルブの中にフィラメントが2本あり、ハイビーム用（60W）とロービーム用（55W）に分かれています。",
      "口金のつば（金属の輪）に、爪（切り欠き）が3つあるのが特徴です。",
      "HB2は、H4と同じバルブの別の呼び名です（輸入車などの表記）。",
    ],
    tips: [
      "お客さんには「ハイとローが1個のバルブになっているタイプですか？」と聞くと、H1・H7などの1役のバルブと区別できます。",
      "12Vか24Vかを必ず確認します（トラック・バスは24Vが多い）。",
      FIT_TIP,
    ],
    caution: [
      "H4uは、H4と同じ60/55Wで、口金の位置決めの突起のうち1つだけが細い派生品です。H4uの車にH4が付かない場合があるので、元のバルブの表記（H4かH4uか）を確認してください。",
      HALO_CAUTION,
    ],
    trivia: [
      "LEDに替える場合は、車検（ロービームの光の境目＝カットオフライン）に対応した製品を選びます。",
    ],
  },
  {
    id: "h1",
    group: "head",
    title: "H1",
    keys: ["H1"],
    img: () => SVGS.halogen(2),
    cap: "ヘッドライト用ハロゲン（口金の形は種類で違います）",
    one: "ハイビームかロービームのどちらか1つだけを受け持つ、ヘッドライト用バルブです。12V 55W。",
    where: ["ヘッドライト（ハイ、またはロー）"],
    spot: ["フィラメントは1本です。ハイ・ロー兼用のH4と違い、1つのバルブで1つの役割です。"],
    tips: ["ハイ用・ロー用で、バルブが別々の車に使われます。", FIT_TIP],
    caution: [HALO_CAUTION],
    trivia: [],
  },
  {
    id: "h7",
    group: "head",
    title: "H7",
    keys: ["H7"],
    img: () => SVGS.halogen(2),
    cap: "ヘッドライト用ハロゲン（口金の形は種類で違います）",
    one: "ヘッドライト用の、1つの役割を受け持つバルブです。12V 55W（口金の規格名はPX26d）。",
    where: ["ヘッドライト"],
    spot: ["フィラメントは1本です。H11と見た目が似ていますが、口金の形がちがうので入れ替えられません。"],
    tips: [FIT_TIP],
    caution: [HALO_CAUTION, "H7とH11は、どちらも55Wですが、口金が違います。名前を確認して選びます。"],
    trivia: [],
  },
  {
    id: "hb3",
    group: "head",
    title: "HB3",
    keys: ["HB3"],
    img: () => SVGS.halogenL(),
    cap: "ヘッドライト用ハロゲン（口金の形は種類で違います）",
    one: "ヘッドライト用バルブです。12V 65W。ハイビームに使われることが多いバルブです。",
    where: ["ヘッドライト（ハイビーム用が多い）"],
    spot: ["アメリカの規格名では9005といいます。口金の規格名はP20dです。"],
    tips: [FIT_TIP, "HB4と形が近いので、お客さんが持ってきた古いバルブの刻印（HB3かHB4か）を確認します。"],
    caution: [HALO_CAUTION],
    trivia: [],
  },
  {
    id: "hb4",
    group: "head",
    title: "HB4",
    keys: ["HB4"],
    img: () => SVGS.halogenL(),
    cap: "ヘッドライト用ハロゲン（口金の形は種類で違います）",
    one: "ヘッドライトやフォグランプに使われるバルブです。12V 55W。ロービーム用として使われることが多いバルブです。",
    where: ["ヘッドライト（ロービーム）", "フォグランプ"],
    spot: ["アメリカの規格名では9006といいます。口金の規格名はP22dです。HB3とは、口金の形がちがいます。"],
    tips: [FIT_TIP, "HB3と間違えやすいので、刻印を確認します。"],
    caution: [HALO_CAUTION],
    trivia: [],
  },
  {
    id: "h3",
    group: "fog",
    title: "H3",
    keys: ["H3"],
    img: () => SVGS.halogenLead(),
    cap: "リード線つきの形のイメージ",
    one: "フォグランプ用のバルブです。バルブからコード（リード線）が出ていて、端子で電源につなぎます。",
    where: ["フォグランプ", "12Vは55W、24Vは55W・70W（店のデータより）"],
    spot: [
      "バルブの根元から、配線（コード）が出ています。",
      "店のデータでは、コードの先の端子（平端子）を、ギボシ端子に変換するコネクターが入っている商品があります。",
    ],
    tips: ["24Vには55Wと70Wがあります。元のバルブのW数を確認します。", FIT_TIP],
    caution: [HALO_CAUTION, "コネクターの形が合わない場合があります。商品に入っている変換コネクターを確認します。"],
    trivia: [],
  },
  {
    id: "h8",
    group: "fog",
    title: "H8",
    keys: ["H8"],
    img: () => SVGS.halogenL(),
    cap: "H8・H11・H16の口金のイメージ（形は似ています）",
    one: "フォグランプ用のバルブです。12V 35W。H8・H11・H16は形がとても似ています。",
    where: ["フォグランプ"],
    spot: ["H8・H11・H16は、L字の口金が似ていますが、ワット数が違います（H8=35W、H11=55W、H16=19W）。"],
    tips: ["見た目で判断せず、必ず名前（H8かH11かH16か）で選びます。", FIT_TIP],
    caution: [
      HALO_CAUTION,
      "形が合っても、ワット数が違うバルブを付けると、車の配線やソケットに合わない負荷がかかります。元と同じ規格のものを選んでください。",
    ],
    trivia: [],
  },
  {
    id: "h11",
    group: "fog",
    title: "H11",
    keys: ["H11"],
    img: () => SVGS.halogenL(),
    cap: "H8・H11・H16の口金のイメージ（形は似ています）",
    one: "フォグランプやヘッドライトに使われるバルブです。12V 55W。",
    where: ["フォグランプ", "ヘッドライト（車による）"],
    spot: ["H8・H16と形が似ていますが、55Wでいちばん大きなW数です。"],
    tips: ["H8・H16と間違えないよう、元のバルブの表記を確認します。", FIT_TIP],
    caution: [HALO_CAUTION, "ワット数が違うH8・H16の車に、H11を付けないでください。"],
    trivia: [],
  },
  {
    id: "h16",
    group: "fog",
    title: "H16",
    keys: ["H16"],
    img: () => SVGS.halogenL(),
    cap: "H8・H11・H16の口金のイメージ（形は似ています）",
    one: "フォグランプ用のバルブです。12V 19W。",
    where: ["フォグランプ"],
    spot: ["H8・H11と形が似ていますが、19Wでいちばん小さなW数です。"],
    tips: ["お客さんが「H16」と言ったときは、元のバルブの表記も確認してください（海外の「5202」をH16と呼ぶ商品もあります）。", FIT_TIP],
    caution: [HALO_CAUTION, "ワット数が違うバルブ（H8・H11）を付けないでください。"],
    trivia: [],
  },
  {
    id: "hid-s",
    group: "hid",
    title: "D2S・D4S（シェードなし）",
    keys: ["D2S", "D4S"],
    img: () => SVGS.hid(false),
    cap: "シェードのない放電管のイメージ",
    one: "HID（キセノン）ヘッドライトのバルブです。35W。プロジェクター式のヘッドライトで使われます。",
    where: ["ヘッドライト（HID車）"],
    spot: [
      "ガラスに黒い塗装（シェード）がないのがSタイプ（D2S・D4S）です。",
      "D4はD2の水銀フリー版です。D2とD4は、点灯させる装置（バラスト）が別のものなので、入れ替えできません。",
    ],
    tips: ["「D2S」か「D4S」か、バルブに刻印されている名前を確認します。", FIT_TIP],
    caution: [
      "点灯させるときに高い電圧が出ます。交換するときは、ライトを消して、エンジンを止めて、時間をおいてから作業してください。感電に注意します。",
      "D2S・D4Sの車に、シェード付きのD2R・D4Rを付けると、光の形が変わることがあります。",
    ],
    trivia: ["HID用のLED（スフィアライトの純正HID用LEDなど）に替えることもできますが、車によって付けられない場合があります。"],
  },
  {
    id: "hid-r",
    group: "hid",
    title: "D2R・D4R（シェードあり）",
    keys: ["D2R", "D4R"],
    img: () => SVGS.hid(true),
    cap: "ガラスの上側に黒いシェードがあるイメージ",
    one: "HID（キセノン）ヘッドライトのバルブです。35W。リフレクター式のヘッドライトで使われます。",
    where: ["ヘッドライト（HID車）"],
    spot: [
      "ガラスの一部に、黒い塗装（シェード）があるのがRタイプ（D2R・D4R）です。シェードは、光を遮って、まぶしさを防ぎます。",
      "D4はD2の水銀フリー版です。D2とD4は、点灯させる装置（バラスト）が別のものなので、入れ替えできません。",
    ],
    tips: ["シェードの有無（SかRか）は、見て分かる大きな違いです。お客さんの古いバルブを見せてもらいます。", FIT_TIP],
    caution: [
      "点灯させるときに高い電圧が出ます。交換するときは、ライトを消して、エンジンを止めて、時間をおいてから作業してください。感電に注意します。",
      "D2R・D4Rの車に、シェードのないD2S・D4Sを付けると、光の形が変わることがあります。",
    ],
    trivia: [],
  },
  {
    id: "t5",
    group: "wedge",
    title: "T5",
    keys: ["T5"],
    img: () => SVGS.wedge(20),
    cap: "T5のイメージ（ウェッジ球の中で一番細い）",
    one: "とても小さなウェッジ球です。主にメーター（計器）の照明に使われます。",
    where: ["メーター", "24Vは1.2W（店のデータより）"],
    spot: ["ガラスの足（ウェッジ）を、そのままソケットに差し込みます。T10より細いです。"],
    tips: ["メーターまわりのバルブは、ほかと見た目が似ています。元のバルブの形（T5かT10か）を確認します。"],
    caution: [],
    trivia: ["「T」はチューブ（筒）の形のこと、数字はガラスの太さの目安です（数字が大きいほど太い）。"],
  },
  {
    id: "t10",
    group: "wedge",
    title: "T10",
    keys: ["T10"],
    img: () => SVGS.wedge(28),
    cap: "T10のイメージ",
    one: "いちばんよく出る小さなウェッジ球です。ポジション・ナンバー・ルーム・メーターなど、車のいろいろな場所で使われます。",
    where: ["ポジションランプ（車幅灯）", "ナンバー灯", "ルームランプ", "メーター・表示灯", "12Vは5W、24Vも5W（店のデータより）"],
    spot: [
      "ガラスの足（ウェッジ）をソケットに差し込みます。",
      "アンバー色（オレンジ）のT10は、方向指示灯などに使われます（店のデータに、欧州の認証を取ったアンバー色の商品があります）。",
    ],
    tips: ["いちばん出るので、在庫をいつも確認します。", "12Vか24Vかを確認します。", FIT_TIP],
    caution: ["同じT10でも、W数・色（白とアンバー）が違う商品があります。元のバルブと同じものを選びます。"],
    trivia: ["T10のガラスの太さは、約10mmです（小糸製作所のT10の寸法より）。"],
  },
  {
    id: "t16",
    group: "wedge",
    title: "T16",
    keys: ["T16"],
    img: () => SVGS.wedge(40),
    cap: "T16のイメージ（T10より太い）",
    one: "T10より太いウェッジ球です。バックランプなどに使われます。12V 16W。",
    where: ["バックランプ", "表示灯"],
    spot: ["T10より太く、T20より細いウェッジ球です。ガラスの足をソケットに差し込みます。"],
    tips: [FIT_TIP],
    caution: [],
    trivia: [],
  },
  {
    id: "t20",
    group: "wedge",
    title: "T20（シングル・ダブル）",
    keys: ["T20"],
    img: () => SVGS.wedge(50),
    cap: "T20のイメージ（ウェッジ球で太いほう）",
    one: "大きめのウェッジ球です。ウインカー（方向指示灯）、バックランプ、テール＆ストップなどに使われます。",
    where: ["ウインカー（方向指示灯）", "バックランプ", "後ろのテール＆ストップ", "シングル（21W）とダブル（21/5W）がある"],
    spot: [
      "シングル球は、フィラメントが1本（21W）。ウインカーやバックランプなど、1つの役割に使います。",
      "ダブル球は、フィラメントが2本（21/5W）。ブレーキ（ストップ）とテール（尾灯）の2つの役割に使います。",
      "アンバー色のT20は、ウインカー用です（店のデータにあります）。",
    ],
    tips: [
      "シングルとダブルは、差し込み口の形が違うので、基本的に入れ替えできません。W数の表記（21Wか21/5Wか）で必ず確認します。",
      FIT_TIP,
    ],
    caution: ["ウインカー（アンバー）と、バックランプ（白）を間違えないようにします。"],
    trivia: [],
  },
  {
    id: "s25",
    group: "bay",
    title: "S25（シングル・ダブル）",
    keys: ["S25"],
    img: () => SVGS.bayonet(false),
    imgAlt: [["ダブル球（ピンの高さが左右で違う）", () => SVGS.bayonet(true)]],
    cap: "シングル球のイメージ（左右のピンが同じ高さ）",
    one: "丸いガラス球に、金属の筒の口金がついたバルブです。差し込んで、回して固定します（バヨネット式）。ウインカー・ブレーキ・テールなどに使われます。",
    where: ["ウインカー（方向指示灯）", "ブレーキ・テール", "バックランプなど", "12Vと24V（ディーゼル車・トラック）がある"],
    spot: [
      "シングル球は、フィラメントが1本。口金の左右のピンが、同じ高さにあります。",
      "ダブル球は、フィラメントが2本（例：21/5W）。口金の左右のピンの高さが違うので、一方向にしか入りません。",
      "店のデータでは、12Vに21W・23W・21/5W・23/8Wが、24Vに25W・25/10Wがあります。",
    ],
    tips: [
      "シングルかダブルかは、W数の表記（21Wか21/5Wか）と、ピンの高さで見分けます。",
      "ダブル球は、ピンの位置が違うので、シングルのソケットに入れ替えできません。",
      "12Vか24Vかを必ず確認します。",
    ],
    caution: [
      "店のデータの一部に、S25を「ウェッジ球」と書いたものがあります。S25は、バヨネット式（ピンを回して固定）で、ウェッジ球ではありません。",
    ],
    trivia: ["口金の規格名は、シングル球がBA15s、ダブル球がBAY15dです。"],
  },
  {
    id: "g18",
    group: "bay",
    title: "G18",
    keys: ["G18"],
    img: () => SVGS.bayonet(false, 22),
    cap: "G18のイメージ（丸く小さなガラス球・バヨネット口金）",
    one: "丸くて小さなガラス球のバルブです。口金は、S25と同じ形（バヨネット式のBA15s）です。ナンバー灯やパーキングランプなどに使われます。",
    where: ["ライセンスランプ（ナンバー灯）", "車幅灯・表示灯", "バックランプ・ウインカー（24V車）", "12Vは10W、24Vは5W・12W（店のデータより）"],
    spot: ["球がS25より小さい、丸いバルブです。口金はS25と同じ形で、左右のピンが同じ高さです。"],
    tips: ["トラック・商用車（24V）に多く使われます。12Vか24Vかを確認します。", FIT_TIP],
    caution: [],
    trivia: ["「G」は、球の形が丸い（グローブ）ことを表します。"],
  },
  {
    id: "t10x31",
    group: "fest",
    title: "T10×31（両口金）",
    keys: ["T10×31", "T10X31"],
    img: () => SVGS.festoon(),
    cap: "両口金（フェストゥーン）のイメージ",
    one: "細長いガラス管の両端に、金具がついたバルブです。ルームランプ・ナンバー灯などに使われます。",
    where: ["ルームランプ（車内灯）", "方向指示灯・ナンバー灯（車による）"],
    spot: ["両端の金具を、ソケットのバネで挟んで固定します。"],
    tips: [
      "長さ（31mm、36mmなど）とW数（5W、10Wなど）を、元のバルブと同じにします。",
      "店のデータでは、12V 10Wの商品があります。",
      FIT_TIP,
    ],
    caution: ["長さが違うと、ソケットに入らなかったり、外れたりします。必ず元のバルブと見比べます。"],
    trivia: ["W数は、5W（C5W）や10Wなどがあります。元のバルブと同じものを選びます。"],
  },
];

export const BASICS = [
  ["① どこのバルブですか？", "ヘッドライト・フォグ・ウインカー・バック・ポジション・ナンバー・ルーム・メーターなど。使う場所で、バルブの種類が決まります。"],
  ["② 12Vですか、24Vですか？", "乗用車は12V、トラック・バスなどは24Vが多いです。ここを間違えると、すぐ切れたり、点かなかったりします。"],
  ["③ ワット（W）はいくつですか？", "元のバルブに、「12V 55W」のように書いてあります。元と同じW数のバルブを選びます。"],
  ["④ 形（口金）は？", "H4・H7・T10・S25など、形で名前が決まっています。同じ場所でも、車によって形が違います。"],
  ["⑤ ハロゲン・HID・LED？", "ハロゲンは普通の電球タイプ、HIDは放電で光るタイプ、LEDは省電力のタイプです。元のタイプに合うものを選びます。"],
];

export const WORDS = [
  ["ハロゲンバルブ", "ガラスの中のフィラメント（細い線）が熱で光る、普通の電球のタイプ。H4・H7などのヘッドライト用や、フォグ用があります。"],
  ["HID（キセノン）", "ガラスの中の放電で光る、明るいタイプ。点灯させる装置（バラスト）が必要です。D2S・D4Sなど。"],
  ["フィラメント", "バルブの中で光る、細い線。シングル球は1本、ダブル球は2本あります。"],
  ["口金（くちがね）", "バルブの根元の、ソケットに差し込む部分の形。H4・T10・S25などの名前は、主に口金の形で決まります。"],
  ["シングル球・ダブル球", "1つの役割（1フィラメント）がシングル球、2つの役割（2フィラメント。ブレーキとテールなど）がダブル球。"],
  ["ウェッジ球", "ガラスの足を、そのままソケットに差し込むタイプ（T5・T10・T16・T20）。"],
  ["バヨネット式", "金属の筒を差し込んで、回して固定するタイプ（S25・G18）。"],
  ["12V・24V", "車の電気の電圧。乗用車は12V、トラック・バスは24Vが多い。"],
  ["W（ワット）", "消費電力。数字が大きいほど、明るく、熱くなります。"],
];

// 場所 → 形状の早見表（店のバルブデータにある用途から）
export const WHERE = [
  ["ヘッドライト", "H4（ハイ・ロー兼用）、H1・H7・HB3・HB4、HID（D2S・D2R・D4S・D4R）"],
  ["フォグランプ", "H3、H8、H11、H16、HB4"],
  ["ウインカー・ブレーキ・テール", "S25（シングル・ダブル）、T20（シングル・ダブル）"],
  ["バックランプ", "T16、T20、（24V車は）G18"],
  ["ポジション・ナンバー・ルーム", "T10、G18、T10×31（両口金）"],
  ["メーター・表示灯", "T5、T10"],
];

// ---- 画面のHTML ----
const li = (a) => a.map((x) => `<li>${esc(x)}</li>`).join("");
const volt = (spec, cat) => (/24\s*v/i.test(`${spec} ${cat}`) ? "24V" : /12\s*v/i.test(spec) ? "12V" : "");
const specText = (s) =>
  String(s || "")
    .replace(/^(バルブ|BULB|ハロゲン|バルフ)\s*/i, "")
    .replace(/(\d+)\s*v\s*-?\s*/i, "$1V ")
    .replace(/([\d./]+)\s*w\b/i, "$1W")
    .trim();
const shapeKey = (s) =>
  String(s || "")
    .toUpperCase()
    .replace(/\s/g, "");

export function bulbGuideHtml(list) {
  const used = new Set();
  const prods = (g) => {
    const keys = g.keys.map(shapeKey);
    const rows = list.filter((b) => keys.includes(shapeKey(b.shape)));
    rows.forEach((b) => used.add(b.id));
    return rows;
  };
  const section = (g) => {
    const rows = prods(g);
    const photos = rows.filter((b) => b.photo).slice(0, 3);
    return `<section class="bg-sec" id="bg-${g.id}"><h3>${esc(g.title)}</h3>
      <p class="bg-one">${esc(g.one)}</p>
      <div class="bg-top">
        <figure class="bg-fig">${g.img()}<figcaption>${esc(g.cap)}<br><small>模式図（実物とは細部が違います）</small></figcaption></figure>
        ${g.imgAlt ? g.imgAlt.map(([c, f]) => `<figure class="bg-fig">${f()}<figcaption>${esc(c)}<br><small>模式図</small></figcaption></figure>`).join("") : ""}
        ${photos.map((b) => `<figure class="bg-fig"><span class="oph"><img data-bphoto="${esc(b.id)}" alt=""></span><figcaption>店の写真<br><small>${esc(b.code)}</small></figcaption></figure>`).join("")}
      </div>
      <div class="bg-cols">
        <div><h4>どこに使う？</h4><ul>${li(g.where)}</ul></div>
        <div><h4>見分けるポイント</h4><ul>${li(g.spot)}</ul></div>
      </div>
      <h4>接客のポイント</h4><ul>${li(g.tips)}</ul>
      ${g.caution.length ? `<div class="bg-warn"><b>注意</b><ul>${li(g.caution)}</ul></div>` : ""}
      ${g.trivia.length ? `<div class="bg-learn"><b>豆知識（学習用）</b><ul>${li(g.trivia)}</ul></div>` : ""}
      <h4>店のバルブ（${rows.length}件）</h4>
      ${
        rows.length
          ? `<div class="bg-tw"><table class="bg-t"><thead><tr><th>品番</th><th>規格</th><th>用途</th><th>入数</th></tr></thead><tbody>${rows
              .map(
                (b) =>
                  `<tr><td class="mono">${esc(b.code)}</td><td>${esc(specText(b.spec))}${volt(b.spec, b.cat) ? "" : ""}</td><td>${esc(b.use || "")}${b.note ? ` <i>【${esc(b.note)}】</i>` : ""}</td><td>${esc(b.unit ?? "")}</td></tr>`,
              )
              .join("")}</tbody></table></div>`
          : '<p class="bg-none">この形状の商品は、まだ店のデータに入っていません。</p>'
      }
    </section>`;
  };
  const groups = GROUPS.map(
    ([gid, gname]) => `<h2 class="bg-gh" id="bg-g-${gid}">${esc(gname)}</h2>${GUIDE.filter((g) => g.group === gid).map(section).join("")}`,
  ).join("");
  // データにあるのに、解説のない形状
  const rest = list.filter((b) => !used.has(b.id));
  const restHtml = rest.length
    ? `<section class="bg-sec"><h3>そのほかの形状</h3><p class="bg-one">解説をまだ書いていない形状です。</p><div class="bg-tw"><table class="bg-t"><thead><tr><th>形状</th><th>品番</th><th>規格</th><th>用途</th></tr></thead><tbody>${rest
        .map((b) => `<tr><td>${esc(b.shape || "")}</td><td class="mono">${esc(b.code)}</td><td>${esc(specText(b.spec))}</td><td>${esc(b.use || "")}</td></tr>`)
        .join("")}</tbody></table></div></section>`
    : "";
  return `<div id="bguide" class="bguide">
    <div class="bg-head"><div><h2>バルブ解説書</h2><div class="bg-sub">初心者でも、見ながらお客さんに説明できる／学習用（店のバルブデータとつながっています）</div></div><button class="btn" data-bprint="1">印刷する</button></div>
    <nav class="bg-toc">${GROUPS.map(([id, n]) => `<button class="chip" data-bgo="g-${id}">${esc(n.replace(/（.*$/, ""))}</button>`).join("")}<button class="chip" data-bgo="basic">選び方の基本</button><button class="chip" data-bgo="words">用語</button></nav>
    <section class="bg-sec" id="bg-basic"><h3>バルブ選びの基本（お客さんに聞くこと）</h3>
      ${BASICS.map(([q, a]) => `<div class="bg-qa"><b>${esc(q)}</b><span>${esc(a)}</span></div>`).join("")}
      <h4>場所から探す</h4>
      <div class="bg-tw"><table class="bg-t"><tbody>${WHERE.map(([a, b]) => `<tr><th>${esc(a)}</th><td>${esc(b)}</td></tr>`).join("")}</tbody></table></div>
    </section>
    ${groups}${restHtml}
    <section class="bg-sec" id="bg-words"><h3>用語集</h3>${WORDS.map(([w, d]) => `<div class="bg-qa"><b>${esc(w)}</b><span>${esc(d)}</span></div>`).join("")}</section>
    <p class="bg-foot">※ 規格名や数字は一般的なものです。車ごとの適合は、必ず適合表と元のバルブの表記で確認してください。</p>
  </div>`;
}

// ---- Excel用：解説書を、文章・図・表の並び（モデル）にする ----
const R = (t, b = false) => [{ t, b }];
// 図（模式図）の一覧。Excelに絵として入れるため、キーとSVGと説明を返す
export function guideFigs() {
  const out = [];
  for (const g of GUIDE) {
    out.push({ key: `${g.id}-0`, svg: g.img(), cap: g.cap });
    (g.imgAlt || []).forEach(([cap, f], i) => out.push({ key: `${g.id}-${i + 1}`, svg: f(), cap }));
  }
  return out;
}
export function bulbGuideModel(list) {
  const sections = [];
  sections.push({
    title: "バルブ選びの基本（お客さんに聞くこと）",
    blocks: [
      { t: "qalist", items: BASICS.map(([q, a]) => ({ q: R(q), a: R(a) })) },
      { t: "h4", runs: R("場所から探す") },
      { t: "table", head: [R("場所"), R("形状")], rows: WHERE.map(([a, b]) => [R(a, true), R(b)]), pairRows: WHERE.map(([a, b]) => [R(a, true), R(b)]) },
    ],
  });
  for (const [gid, gname] of GROUPS) {
    for (const g of GUIDE.filter((x) => x.group === gid)) {
      const keys = g.keys.map(shapeKey);
      const rows = list.filter((b) => keys.includes(shapeKey(b.shape)));
      const figs = [{ key: `${g.id}-0`, cap: g.cap }, ...(g.imgAlt || []).map(([cap], i) => ({ key: `${g.id}-${i + 1}`, cap }))];
      const blocks = [
        { t: "lead", runs: R(g.one, true) },
        { t: "figs", figs },
        { t: "h4", runs: R("どこに使う？") },
        { t: "list", items: g.where.map((x) => R(x)) },
        { t: "h4", runs: R("見分けるポイント") },
        { t: "list", items: g.spot.map((x) => R(x)) },
        { t: "h4", runs: R("接客のポイント") },
        { t: "list", items: g.tips.map((x) => R(x)) },
      ];
      if (g.caution.length) blocks.push({ t: "box", warn: true, title: R("注意"), items: g.caution.map((x) => R(x)) });
      if (g.trivia.length) blocks.push({ t: "box", warn: false, title: R("豆知識（学習用）"), items: g.trivia.map((x) => R(x)) });
      blocks.push({ t: "h4", runs: R(`店のバルブ（${rows.length}件）`) });
      if (rows.length) {
        blocks.push({
          t: "table",
          head: ["品番", "規格", "用途", "入数"].map((x) => R(x)),
          rows: rows.map((b) => [R(b.code, true), R(specText(b.spec)), R(`${b.use || ""}${b.note ? `【${b.note}】` : ""}`), R(String(b.unit ?? ""))]),
          pairRows: rows.map((b) => [R(`${b.code}　${specText(b.spec)}`, true), R(`${b.use || ""}${b.note ? `【${b.note}】` : ""}　（入数 ${b.unit ?? ""}）`)]),
        });
      } else blocks.push({ t: "note", runs: R("この形状の商品は、まだ店のデータに入っていません。") });
      sections.push({ title: g.title, blocks });
    }
  }
  sections.push({
    title: "用語集",
    blocks: [{ t: "qalist", items: WORDS.map(([w, d]) => ({ q: R(w), a: R(d) })) }, { t: "note", runs: R("※ 規格名や数字は一般的なものです。車ごとの適合は、必ず適合表と元のバルブの表記で確認してください。") }],
  });
  return { title: "バルブ解説書", sub: "初心者でも、見ながらお客さんに説明できる／学習用（店のバルブデータとつながっています）", sections };
}
