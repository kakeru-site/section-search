// 生体認証ロック（WebAuthn）
// 端末ごとのロック解除用。ログイン自体はFirebase側
const KEY = "secsearch.webauthn.v1";
const enc = (buf) =>
  btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
const dec = (s) =>
  Uint8Array.from(atob(s.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((s.length + 3) % 4)), (c) =>
    c.charCodeAt(0),
  );
const rand = (n) => crypto.getRandomValues(new Uint8Array(n));

export function supported() {
  return !!(window.PublicKeyCredential && navigator.credentials && window.isSecureContext);
}
export async function platformAvailable() {
  if (!supported()) return false;
  try {
    return await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable();
  } catch {
    return false;
  }
}
export function registered() {
  try {
    return JSON.parse(localStorage.getItem(KEY) || "null");
  } catch {
    return null;
  }
}
export function settings() {
  const r = registered();
  return { lockMin: r?.lockMin ?? 30 };
}
export function setLockMin(min) {
  const r = registered();
  if (!r) return;
  r.lockMin = min;
  localStorage.setItem(KEY, JSON.stringify(r));
}

export async function register(label = "セクション検索") {
  if (!supported()) throw new Error("このブラウザは生体認証に対応していません（https で開いているか確認）");
  const cred = await navigator.credentials.create({
    publicKey: {
      challenge: rand(32),
      rp: { name: "セクション検索" },
      user: { id: rand(16), name: label, displayName: label },
      pubKeyCredParams: [
        { type: "public-key", alg: -7 },
        { type: "public-key", alg: -257 },
      ],
      authenticatorSelection: {
        authenticatorAttachment: "platform",
        userVerification: "required",
        residentKey: "discouraged",
      },
      timeout: 60000,
      attestation: "none",
    },
  });
  localStorage.setItem(KEY, JSON.stringify({ id: enc(cred.rawId), at: Date.now(), lockMin: 30 }));
  return true;
}

export async function verify() {
  const r = registered();
  if (!r) return true;
  await navigator.credentials.get({
    publicKey: {
      challenge: rand(32),
      allowCredentials: [{ type: "public-key", id: dec(r.id), transports: ["internal", "hybrid"] }],
      userVerification: "required",
      timeout: 60000,
    },
  });
  return true;
}

export function unregister() {
  localStorage.removeItem(KEY);
}
