// データの読み書きとログインをまとめた層。
// Firebase（本番）と デモモード（ブラウザ内保存）を同じ形で扱えるようにしています。
import { FIREBASE, ROOT_PATH, SDK_VERSION } from "./config.js";

export const DEMO = !FIREBASE.apiKey;

/* ---------------- 共通ユーティリティ ---------------- */
const PUSH_CHARS = "-0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz";
let lastPush = 0, lastRand = [];
export function localKey() {           // Firebaseのpush IDと同じ「時刻順に並ぶID」
  let now = Date.now();
  const dup = now === lastPush; lastPush = now;
  const ts = new Array(8);
  for (let i = 7; i >= 0; i--) { ts[i] = PUSH_CHARS.charAt(now % 64); now = Math.floor(now / 64); }
  let id = ts.join("");
  if (!dup) { for (let i = 0; i < 12; i++) lastRand[i] = Math.floor(Math.random() * 64); }
  else { let i = 11; for (; i >= 0 && lastRand[i] === 63; i--) lastRand[i] = 0; if (i >= 0) lastRand[i]++; }
  for (let i = 0; i < 12; i++) id += PUSH_CHARS.charAt(lastRand[i]);
  return id;
}
const split = (p) => p.split("/").filter(Boolean);

/* ---------------- デモ（localStorage） ---------------- */
function demoStore() {
  const KEY = "secsearch.demo.v1";
  let tree = {};
  try { tree = JSON.parse(localStorage.getItem(KEY) || "{}") || {}; } catch { tree = {}; }
  const subs = new Map(); // path -> Set(cb)
  let authCb = null, user = null;
  try { if (sessionStorage.getItem("secsearch.demo.login") || localStorage.getItem("secsearch.demo.login")) user = { uid: "demo", email: "demo@example.com" }; } catch {}

  const get = (path) => split(path).reduce((o, k) => (o == null ? undefined : o[k]), tree);
  const clean = (v) => (v === undefined ? null : JSON.parse(JSON.stringify(v)));
  function put(path, val) {
    const ks = split(path);
    if (!ks.length) { tree = val || {}; return; }
    let o = tree;
    for (let i = 0; i < ks.length - 1; i++) { if (typeof o[ks[i]] !== "object" || o[ks[i]] === null) o[ks[i]] = {}; o = o[ks[i]]; }
    const last = ks[ks.length - 1];
    if (val === null || val === undefined) delete o[last]; else o[last] = clean(val);
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(tree)); } catch (e) { console.warn(e); }
    for (const [p, set] of subs) { const v = get(p); for (const cb of set) cb(v === undefined ? null : clean(v)); }
  }
  return {
    mode: "demo",
    onAuth(cb) { authCb = cb; setTimeout(() => cb(user), 0); },
    async login(email) {
      user = { uid: "demo", email: email || "demo@example.com" };
      try { localStorage.setItem("secsearch.demo.login", "1"); } catch {}
      authCb && authCb(user);
    },
    async logout() { user = null; try { localStorage.removeItem("secsearch.demo.login"); } catch {} authCb && authCb(null); },
    async resetPassword() { throw new Error("デモモードではパスワード再設定は使えません"); },
    get user() { return user; },
    sub(path, cb) {
      if (!subs.has(path)) subs.set(path, new Set());
      subs.get(path).add(cb);
      setTimeout(() => { const v = get(path); cb(v === undefined ? null : clean(v)); }, 0);
      return () => subs.get(path).delete(cb);
    },
    async set(path, v) { put(path, v); save(); },
    async update(path, obj) { for (const [k, v] of Object.entries(obj)) put(path + "/" + k, v); save(); },
    async remove(path) { put(path, null); save(); },
    async push(path, v) { const k = localKey(); put(path + "/" + k, v); save(); return k; },
    async inc(path, by = 1) { put(path, (get(path) || 0) + by); save(); },
  };
}

/* ---------------- Firebase ---------------- */
async function firebaseStore() {
  const base = `https://www.gstatic.com/firebasejs/${SDK_VERSION}/`;
  const [{ initializeApp }, A, D] = await Promise.all([
    import(base + "firebase-app.js"),
    import(base + "firebase-auth.js"),
    import(base + "firebase-database.js"),
  ]);
  const app = initializeApp(FIREBASE);
  const auth = A.getAuth(app);
  try { await A.setPersistence(auth, A.browserLocalPersistence); } catch {}
  const db = D.getDatabase(app);
  const R = (p) => D.ref(db, p ? `${ROOT_PATH}/${p}` : ROOT_PATH);
  const errMsg = (e) => {
    const c = (e && e.code) || "";
    if (c.includes("invalid-credential") || c.includes("wrong-password") || c.includes("user-not-found") || c.includes("invalid-email")) return "メールアドレスかパスワードが違います";
    if (c.includes("too-many-requests")) return "試行回数が多すぎます。少し時間をおいてください";
    if (c.includes("network")) return "ネットワークに接続できません";
    if (c.includes("PERMISSION_DENIED") || String(e).includes("permission_denied")) return "データベースのルールで拒否されました（UIDの設定を確認）";
    return (e && e.message) || String(e);
  };
  const wrap = (fn) => async (...a) => { try { return await fn(...a); } catch (e) { throw new Error(errMsg(e)); } };
  return {
    mode: "firebase",
    onAuth(cb) { A.onAuthStateChanged(auth, (u) => cb(u ? { uid: u.uid, email: u.email } : null)); },
    login: wrap((email, pw) => A.signInWithEmailAndPassword(auth, email, pw)),
    logout: wrap(() => A.signOut(auth)),
    resetPassword: wrap((email) => A.sendPasswordResetEmail(auth, email)),
    get user() { const u = auth.currentUser; return u ? { uid: u.uid, email: u.email } : null; },
    sub(path, cb, onErr) {
      return D.onValue(R(path), (s) => cb(s.val()), (e) => { console.error(e); onErr && onErr(new Error(errMsg(e))); });
    },
    set: wrap((p, v) => D.set(R(p), v)),
    update: wrap((p, obj) => D.update(R(p), obj)),
    remove: wrap((p) => D.remove(R(p))),
    push: wrap(async (p, v) => { const r = D.push(R(p)); await D.set(r, v); return r.key; }),
    inc: wrap((p, by = 1) => D.set(R(p), D.increment(by))),
  };
}

export async function createStore() {
  return DEMO ? demoStore() : await firebaseStore();
}
