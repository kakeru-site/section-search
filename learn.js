// 学び：発見したこと・学んだことを、1タップで書き留め → 間隔をあけて復習カードに出す → 確実に習得するしくみ
// データ：learn/{id}（発見ノート＋復習の進み具合）、learnLog/{日付}（その日に復習した枚数）
// 復習のしくみ：「思い出せた」が続くと、次に出るまでの間が、3日 → 7日 → 14日 → 30日 → 60日 と延びる。
//   5回続けて（30日あけても）思い出せたら「習得済み」。「忘れた」は、最初からやり直し。

let C = null;
export function setup(ctx) {
  C = ctx;
}

const DAY = 86400000;
const INT = [3, 7, 14, 30, 60, 90]; // 思い出せたとき、次に出るまでの日数（何回目の成功か）
const MASTER = 5; // この回数、思い出せたら習得済み
export const TAGS = ["バルブ", "オイル・ケミカル", "部品", "型式・車種", "接客", "お客さんの言葉", "ミス防止", "その他"];

const esc = (s) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const $ = (s, r = document) => r.querySelector(s);
const sod = (t = Date.now()) => {
  const d = new Date(t);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
};
const ymd = (t = Date.now()) => {
  const d = new Date(t);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const mdLabel = (t) => {
  const d = new Date(t);
  return `${d.getMonth() + 1}/${d.getDate()}`;
};

// ---- データ ----
export const allNotes = () => Object.entries(C?.st?.learn || {}).map(([id, v]) => ({ id, ...v }));
export const dueNotes = (now = Date.now()) =>
  allNotes()
    .filter((n) => (n.due || 0) <= now)
    .sort((a, b) => (a.due || 0) - (b.due || 0) || (a.at || 0) - (b.at || 0));
export const dueCount = () => (C?.st?.learn ? dueNotes().length : 0);
const stageOf = (n) => n.stage || 0;
const mastered = (n) => stageOf(n) >= MASTER;
const weak = (n) => (n.lapses || 0) >= 2 && !mastered(n);

// 復習の結果から、次の出題日などを決める（r: ok 思い出せた／shaky あやしい／ng 忘れた）
export function nextState(n, r, now = Date.now()) {
  let stage = stageOf(n);
  let streak = n.streak || 0;
  let lapses = n.lapses || 0;
  let due;
  if (r === "ok") {
    stage = Math.min(stage + 1, 6);
    streak += 1;
    due = sod(now) + INT[Math.min(stage - 1, INT.length - 1)] * DAY;
  } else if (r === "shaky") {
    streak = 0;
    due = sod(now) + DAY; // あやしいときは、段階は進めずに、明日もう一度
  } else {
    stage = 0;
    streak = 0;
    lapses += 1;
    due = now; // 忘れたときは、最初から。今日のうちに、もう一度出す
  }
  return {
    stage,
    streak,
    lapses,
    due,
    last: now,
    reps: (n.reps || 0) + 1,
    masteredAt: stage >= MASTER ? n.masteredAt || now : null,
    updatedAt: now,
  };
}

// 連続で復習した日数（今日まだなら、昨日までの連続）
function streakDays() {
  const log = C?.st?.learnLog || {};
  let d = sod();
  if (!log[ymd(d)]) d -= DAY;
  let n = 0;
  while (log[ymd(d)] > 0) {
    n++;
    d -= DAY;
  }
  return n;
}
const weekCount = () => {
  const log = C?.st?.learnLog || {};
  let n = 0;
  for (let i = 0; i < 7; i++) n += log[ymd(sod() - i * DAY)] || 0;
  return n;
};
const logReview = async () => {
  const k = ymd();
  await C.store.set(`learnLog/${k}`, (C.st.learnLog?.[k] || 0) + 1).catch(() => {});
};

// ---- 問題と答え ----
const promptOf = (n) => {
  if (n.q) return n.q;
  const hint = [n.tag, n.ref ? `「${n.ref}」` : ""].filter(Boolean).join(" ");
  return hint ? `${hint}で、覚えたことは？` : "これについて、覚えたことは？";
};
const dots = (n) => {
  const s = Math.min(stageOf(n), MASTER);
  return `<span class="ln-dots" title="思い出せた回数 ${stageOf(n)}／${MASTER}（${MASTER}で習得済み）">${"●".repeat(s)}${"○".repeat(MASTER - s)}</span>`;
};
const dueLabel = (n, now = Date.now()) => {
  if ((n.due || 0) <= now) return "今日 復習";
  const dd = Math.round((sod(n.due) - sod(now)) / DAY);
  return dd === 1 ? "明日" : `${dd}日後（${mdLabel(n.due)}）`;
};

// ---- 画面の状態 ----
const ui = { tab: "review", filter: "all", tag: "", q: "" };
let sess = null; // 復習の進行：{ queue:[id], i, shown, res:{ok,shaky,ng}, total, retried:Set }

// ---- 書く／直す（どの画面からでも呼べる）----
// ctx: { tag, ref } 今見ている画面から、自動で入る
export async function openAdd(ctx = {}, id = null) {
  const old = id ? C.st.learn?.[id] : null;
  const b = old || { text: "", q: "", tag: ctx.tag || "", ref: ctx.ref || "" };
  const r = await C.modal(
    `<h3>${old ? "発見ノートを直す" : "📝 発見・学んだこと"}</h3><form>
    <div class="field"><label>発見したこと・学んだこと（短くてOK）</label><textarea class="inp" name="text" rows="3" maxlength="500" required autofocus placeholder="例：H4uの車に、H4は付かないことがある（口金の突起が違う）">${esc(b.text)}</textarea></div>
    <div class="field"><label>問い（任意）— 書くと、復習カードの問題になります</label><input class="inp" name="q" maxlength="120" value="${esc(b.q)}" placeholder="例：H4とH4uの違いは？"></div>
    <div class="field"><label>分野</label><input type="hidden" name="tag" value="${esc(b.tag)}"><div class="chips ln-tags" style="margin:2px 0 0">${TAGS.map((t) => `<button type="button" class="chip ${b.tag === t ? "on" : ""}" data-lntag="${esc(t)}">${esc(t)}</button>`).join("")}</div></div>
    <div class="field"><label>関連（品番・型式・お客さんの言葉 など。任意）</label><input class="inp" name="ref" maxlength="60" value="${esc(b.ref)}" placeholder="例：AY080-10037／ひとことメモ"></div>
    </form>
    ${old ? `<p class="ln-info">${dots(old)}　次の復習：${esc(dueLabel(old))}　／　忘れた ${old.lapses || 0}回　／　復習 ${old.reps || 0}回${mastered(old) ? "　✔ 習得済み" : ""}</p>` : '<p class="ln-info">書いたら、<b>明日</b>から復習カードに出ます（思い出せるまで、間隔をあけて何度か出ます）。</p>'}
    <div class="modal-foot">${old ? '<button class="btn danger left" data-m="delete">削除</button><button class="btn left" data-m="reset" title="復習の進みをゼロにする">最初からやり直す</button>' : ""}<button class="btn" data-m="cancel">キャンセル</button>${old ? "" : '<button class="btn" data-m="next">保存して次へ</button>'}<button class="btn pri" data-m="ok">保存</button></div>`,
    {
      onOpen(bg) {
        bg.addEventListener("click", (e) => {
          const c = e.target.closest("[data-lntag]");
          if (!c) return;
          const inp = bg.querySelector("input[name=tag]");
          inp.value = inp.value === c.dataset.lntag ? "" : c.dataset.lntag;
          bg.querySelectorAll("[data-lntag]").forEach((x) => x.classList.toggle("on", x.dataset.lntag === inp.value));
        });
        // Ctrl/⌘+Enter で保存
        bg.addEventListener("keydown", (e) => {
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) bg.querySelector('[data-m="ok"]')?.click();
        });
      },
    },
  );
  if (!r) return;
  if (r.act === "delete") {
    if (!(await C.confirmBox("この発見ノートを削除しますか？", "削除", true))) return;
    await C.store.remove(`learn/${id}`);
    return C.toast("削除しました");
  }
  if (r.act === "reset") {
    await C.store.update(`learn/${id}`, { stage: 0, streak: 0, lapses: 0, reps: 0, due: sod() + DAY, masteredAt: null, updatedAt: Date.now() });
    return C.toast("最初からやり直します（明日、復習に出ます）");
  }
  const now = Date.now();
  const rec = {
    text: String(r.data.text || "").trim(),
    q: String(r.data.q || "").trim(),
    tag: String(r.data.tag || "").trim(),
    ref: String(r.data.ref || "").trim(),
  };
  if (!rec.text) return;
  if (old) {
    await C.store.update(`learn/${id}`, { ...rec, updatedAt: now });
    C.toast("保存しました");
  } else {
    await C.store.set(`learn/${C.localKey()}`, { ...rec, at: now, updatedAt: now, stage: 0, streak: 0, lapses: 0, reps: 0, due: sod(now) + DAY, masteredAt: null });
    C.toast("書き留めました。明日から復習に出ます");
    if (r.act === "next") return openAdd(ctx);
  }
}

// ---- 画面 ----
export function render(main) {
  const due = dueCount();
  const notes = allNotes();
  main.innerHTML = `<div class="view-head"><div><h1 class="h1">学び</h1><div class="tip">学んだこと・発見したことを書き留める → 間隔をあけて復習カードに出る → 確実に身につく</div></div>
    <div class="head-acts"><button class="btn pri" data-lnadd="1">📝 発見を書く</button></div></div>
    <div class="chips ln-tabs">${[
      ["review", `今日の復習${due ? ` <span class="ln-n">${due}</span>` : ""}`],
      ["notes", `発見ノート <span class="ln-n2">${notes.length}</span>`],
      ["stats", "習得状況"],
    ]
      .map(([k, l]) => `<button class="chip ${ui.tab === k ? "on" : ""}" data-lntab="${k}">${l}</button>`)
      .join("")}</div>
    <div id="ln-body"></div>`;
  main.onclick = onClick;
  body();
}
function body() {
  const el = $("#ln-body");
  if (!el) return;
  if (ui.tab === "review") el.innerHTML = reviewHtml();
  else if (ui.tab === "notes") el.innerHTML = notesHtml();
  else el.innerHTML = statsHtml();
}
export function onData(key) {
  if (!["learn", "learnLog"].includes(key)) return;
  // 復習の途中は、画面を作り直さない（カードが切り替わってしまわないように）
  if (ui.tab === "review" && sess && !sess.done) return;
  const m = $("#main");
  if (m && C.st.view === "learn") render(m);
}
const refreshTabs = () => {
  const m = $("#main");
  if (m && C.st.view === "learn") render(m);
};

// ---- 今日の復習 ----
function startSession() {
  const q = dueNotes();
  sess = { queue: q.map((n) => n.id), i: 0, shown: false, res: { ok: 0, shaky: 0, ng: 0 }, total: q.length, retried: new Set(), done: !q.length };
}
function reviewHtml() {
  if (!sess || sess.done) {
    const due = dueNotes();
    if (sess?.done && sess.total) {
      const r = sess.res;
      const left = dueNotes().length;
      return `<div class="ln-card ln-end"><div class="ln-big">🎉 今日の復習 おわり</div>
        <p>${sess.total}枚　思い出せた <b>${r.ok}</b>　あやしい <b>${r.shaky}</b>　忘れた <b>${r.ng}</b></p>
        <p class="ln-sub">連続 ${streakDays()}日</p>
        ${left ? `<button class="btn pri" data-lnstart="1">続けて ${left}枚 復習する</button>` : ""}</div>`;
    }
    if (!due.length) {
      const next = allNotes()
        .filter((n) => (n.due || 0) > Date.now())
        .sort((a, b) => a.due - b.due)[0];
      return `<div class="ln-card ln-end"><div class="ln-big">${allNotes().length ? "✔ 今日の復習は ありません" : "まだ発見ノートがありません"}</div>
        <p class="ln-sub">${allNotes().length ? `次の復習：${esc(dueLabel(next || {}))}　／　連続 ${streakDays()}日` : "仕事中に「あれ？」「なるほど」と思ったら、すぐ書いておきましょう。"}</p>
        <button class="btn pri" data-lnadd="1">📝 発見を書く</button></div>`;
    }
    return `<div class="ln-card ln-end"><div class="ln-big">復習 <b>${due.length}</b>枚</div>
      <p class="ln-sub">1枚ずつ、答えを思い出してから「答えを見る」。連続 ${streakDays()}日</p>
      <button class="btn pri ln-go" data-lnstart="1">復習をはじめる</button></div>`;
  }
  const n = { id: sess.queue[sess.i], ...C.st.learn?.[sess.queue[sess.i]] };
  if (!C.st.learn?.[n.id]) {
    sess.i += 1;
    if (sess.i >= sess.queue.length) sess.done = true;
    return reviewHtml();
  }
  const prog = `${Math.min(sess.i + 1, sess.queue.length)} / ${sess.queue.length}`;
  return `<div class="ln-prog"><span>${prog}</span><i style="width:${Math.round((sess.i / sess.queue.length) * 100)}%"></i></div>
    <div class="ln-card"><div class="ln-meta">${n.tag ? `<span class="ln-tag">${esc(n.tag)}</span>` : ""}${n.ref ? `<span class="ln-ref">${esc(n.ref)}</span>` : ""}${dots(n)}</div>
      <div class="ln-q">${esc(promptOf(n))}</div>
      ${
        sess.shown
          ? `<div class="ln-a"><div class="ln-alab">答え・覚えたこと</div>${esc(n.text).replace(/\n/g, "<br>")}</div>
             <div class="ln-ask">思い出せましたか？</div>
             <div class="ln-btns"><button class="btn ln-ng" data-lng="ng"><b>忘れた</b><small>1</small></button><button class="btn ln-sh" data-lng="shaky"><b>あやしい</b><small>2</small></button><button class="btn pri ln-ok" data-lng="ok"><b>思い出せた</b><small>3</small></button></div>`
          : `<div class="ln-hint">まず、頭の中で答えを思い出してみましょう</div><button class="btn pri ln-show" data-lnshow="1">答えを見る<small>Space</small></button>`
      }
    </div>`;
}
async function grade(r) {
  if (!sess || !sess.shown) return;
  const id = sess.queue[sess.i];
  const n = C.st.learn?.[id];
  if (n) {
    const patch = nextState({ id, ...n }, r);
    sess.res[r] += 1;
    C.store.update(`learn/${id}`, patch).catch((e) => C.toast("保存できませんでした：" + e.message, true));
    logReview();
    if (r === "ng" && !sess.retried.has(id)) {
      sess.retried.add(id);
      sess.queue.push(id); // 忘れたものは、最後にもう一度
    }
    if (r === "ok" && patch.stage === MASTER) C.toast("🎉 習得済みになりました！");
  }
  sess.i += 1;
  sess.shown = false;
  if (sess.i >= sess.queue.length) sess.done = true;
  refreshTabs();
}

// ---- 発見ノート（一覧）----
function notesHtml() {
  const all = allNotes();
  const now = Date.now();
  const f = {
    all: () => true,
    due: (n) => (n.due || 0) <= now,
    prac: (n) => !mastered(n),
    done: mastered,
    weak,
  }[ui.filter];
  const nq = ui.q.trim().toLowerCase();
  const list = all
    .filter(f)
    .filter((n) => !ui.tag || n.tag === ui.tag)
    .filter((n) => !nq || `${n.text} ${n.q || ""} ${n.ref || ""} ${n.tag || ""}`.toLowerCase().includes(nq))
    .sort((a, b) => (b.at || 0) - (a.at || 0));
  const cnt = { all: all.length, due: all.filter((n) => (n.due || 0) <= now).length, prac: all.filter((n) => !mastered(n)).length, done: all.filter(mastered).length, weak: all.filter(weak).length };
  const tags = [...new Set(all.map((n) => n.tag).filter(Boolean))];
  return `<div class="toolbar"><label class="search"><input class="inp" id="ln-q" type="search" autocomplete="off" placeholder="発見ノートを検索" value="${esc(ui.q)}"></label></div>
    <div class="chips">${[
      ["all", "すべて"],
      ["due", "復習待ち"],
      ["prac", "練習中"],
      ["done", "習得済み"],
      ["weak", "苦手"],
    ]
      .map(([k, l]) => `<button class="chip ${ui.filter === k ? "on" : ""}" data-lnf="${k}">${l} ${cnt[k]}</button>`)
      .join("")}</div>
    ${tags.length ? `<div class="chips">${tags.map((t) => `<button class="chip ${ui.tag === t ? "on" : ""}" data-lntf="${esc(t)}">${esc(t)}</button>`).join("")}</div>` : ""}
    <div class="ln-list">${
      list
        .map(
          (n) => `<button class="ln-row ${mastered(n) ? "done" : ""}" data-lnedit="${esc(n.id)}">
        <div class="ln-rt">${n.q ? `<b>${esc(n.q)}</b><br>` : ""}${esc(n.text).replace(/\n/g, "<br>")}</div>
        <div class="ln-rm">${n.tag ? `<span class="ln-tag">${esc(n.tag)}</span>` : ""}${n.ref ? `<span class="ln-ref">${esc(n.ref)}</span>` : ""}${dots(n)}<span>${mastered(n) ? "✔ 習得済み" : esc(dueLabel(n))}</span>${weak(n) ? '<span class="ln-weak">苦手</span>' : ""}<span class="ln-date">${mdLabel(n.at || now)}</span></div></button>`,
        )
        .join("") || '<p class="empty">該当する発見ノートはありません</p>'
    }</div>`;
}

// ---- 習得状況 ----
function statsHtml() {
  const all = allNotes();
  if (!all.length) return '<div class="ln-card ln-end"><div class="ln-big">まだ発見ノートがありません</div><p class="ln-sub">書いて、復習すると、ここに習得の様子が出ます。</p><button class="btn pri" data-lnadd="1">📝 発見を書く</button></div>';
  const md = all.filter(mastered).length;
  const prac = all.filter((n) => !mastered(n) && (n.reps || 0) > 0).length;
  const fresh = all.length - md - prac;
  const pct = Math.round((md / all.length) * 100);
  const by = {};
  for (const n of all) {
    const k = n.tag || "（分野なし）";
    (by[k] ||= { n: 0, md: 0, pr: 0 });
    by[k].n++;
    if (mastered(n)) by[k].md++;
    else if ((n.reps || 0) > 0) by[k].pr++;
  }
  const rows = Object.entries(by).sort((a, b) => b[1].n - a[1].n);
  const weaks = all.filter(weak).sort((a, b) => (b.lapses || 0) - (a.lapses || 0)).slice(0, 5);
  return `<div class="ln-kpis">
      <div class="ln-kpi"><b>${md}</b><span>習得済み</span></div>
      <div class="ln-kpi"><b>${prac}</b><span>練習中</span></div>
      <div class="ln-kpi"><b>${fresh}</b><span>まだ復習してない</span></div>
      <div class="ln-kpi"><b>${streakDays()}</b><span>連続 日</span></div>
      <div class="ln-kpi"><b>${weekCount()}</b><span>今週の復習 回</span></div></div>
    <div class="ln-card"><div class="ln-alab">全体の習得率　${pct}%（${md} / ${all.length}）</div>
      <div class="ln-bar"><i class="m" style="width:${(md / all.length) * 100}%"></i><i class="p" style="width:${(prac / all.length) * 100}%"></i></div>
      <div class="ln-leg"><span><i class="m"></i>習得済み</span><span><i class="p"></i>練習中</span><span><i class="f"></i>まだ</span></div></div>
    <h3 class="ln-h">分野ごと</h3>
    <div class="ln-card">${rows
      .map(
        ([k, v]) => `<div class="ln-trow"><span class="ln-tn">${esc(k)}</span><div class="ln-bar sm"><i class="m" style="width:${(v.md / v.n) * 100}%"></i><i class="p" style="width:${(v.pr / v.n) * 100}%"></i></div><span class="ln-tc">${v.md}/${v.n}</span></div>`,
      )
      .join("")}</div>
    ${weaks.length ? `<h3 class="ln-h">苦手（2回以上 忘れた）</h3><div class="ln-list">${weaks.map((n) => `<button class="ln-row" data-lnedit="${esc(n.id)}"><div class="ln-rt">${n.q ? `<b>${esc(n.q)}</b><br>` : ""}${esc(n.text)}</div><div class="ln-rm"><span class="ln-weak">忘れた ${n.lapses}回</span>${n.tag ? `<span class="ln-tag">${esc(n.tag)}</span>` : ""}</div></button>`).join("")}</div>` : ""}
    <p class="ln-sub" style="margin-top:14px">「習得済み」＝ ${MASTER}回続けて思い出せた（間隔：1日 → 3日 → 7日 → 14日 → 30日）。習得済みも、60日・90日あけて確認に出ます。</p>`;
}

// ---- 操作 ----
function onClick(e) {
  const b = e.target.closest("button");
  if (!b) return;
  const d = b.dataset;
  if (d.lnadd) return openAdd(C.context?.() || {});
  if (d.lntab) {
    ui.tab = d.lntab;
    return refreshTabs();
  }
  if (d.lnstart) {
    startSession();
    return refreshTabs();
  }
  if (d.lnshow) {
    if (sess) sess.shown = true;
    return body();
  }
  if (d.lng) return grade(d.lng);
  if (d.lnf) {
    ui.filter = d.lnf;
    return body();
  }
  if (d.lntf !== undefined) {
    ui.tag = ui.tag === d.lntf ? "" : d.lntf;
    return body();
  }
  if (d.lnedit) return openAdd({}, d.lnedit);
}
document.addEventListener("input", (e) => {
  if (e.target?.id === "ln-q") {
    ui.q = e.target.value;
    const pos = e.target.selectionStart;
    body();
    const i = $("#ln-q");
    if (i) {
      i.focus();
      i.setSelectionRange(pos, pos);
    }
  }
});
// キー操作（復習中）：Space/Enter＝答えを見る　1＝忘れた　2＝あやしい　3＝思い出せた
document.addEventListener("keydown", (e) => {
  if (!C || C.st.view !== "learn" || ui.tab !== "review" || !sess || sess.done) return;
  if (e.isComposing || e.ctrlKey || e.metaKey || e.altKey) return;
  if (/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName || "") || document.querySelector(".modal-bg")) return;
  if (!sess.shown && (e.key === " " || e.key === "Enter")) {
    e.preventDefault();
    sess.shown = true;
    body();
  } else if (sess.shown && ["1", "2", "3"].includes(e.key)) {
    e.preventDefault();
    grade({ 1: "ng", 2: "shaky", 3: "ok" }[e.key]);
  }
});
// 画面を開き直したら、復習のやり直し（最新の「今日の分」から）
export function reset() {
  sess = null;
}
