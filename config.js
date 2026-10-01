// ============================================================
//  セクション検索 — 設定ファイル
//  Firebaseコンソール → プロジェクトの設定 → マイアプリ → 「SDK の設定と構成」
//  に表示される firebaseConfig の中身を、下の FIREBASE にそのまま貼ってください。
//
//  ※ apiKey が空のままだと「デモモード」（このブラウザにだけ保存）で動きます。
//  ※ この値はGitHubに公開されても問題ありません（守りはデータベースのルールが担当）。
// ============================================================
export const FIREBASE = {
  apiKey: "AIzaSyDF6k9f4q2Mg502d_6WNr6jJRzm2xnLxpM",
  authDomain: "mame-chishiki-15b2c.firebaseapp.com",
  databaseURL: "https://mame-chishiki-15b2c-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "mame-chishiki-15b2c",
  storageBucket: "mame-chishiki-15b2c.firebasestorage.app",
  messagingSenderId: "595151039388",
  appId: "1:595151039388:web:785747b000406dd756b659"
};

// データベースの中で、このアプリが使う置き場所（豆知識帳などと混ざらないように分けます）
export const ROOT_PATH = "sectionApp";

// Firebase SDK のバージョン
export const SDK_VERSION = "10.12.2";
