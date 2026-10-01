// Firebaseの設定
// コンソール > プロジェクトの設定 > マイアプリ の firebaseConfig をそのまま貼る
// apiKey が空だとデモモード（ブラウザ内保存）で動く
// ※この値は公開されてもOK。守りはDBのルール側
export const FIREBASE = {
  apiKey: "AIzaSyDF6k9f4q2Mg502d_6WNr6jJRzm2xnLxpM",
  authDomain: "mame-chishiki-15b2c.firebaseapp.com",
  databaseURL: "https://mame-chishiki-15b2c-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "mame-chishiki-15b2c",
  storageBucket: "mame-chishiki-15b2c.firebasestorage.app",
  messagingSenderId: "595151039388",
  appId: "1:595151039388:web:785747b000406dd756b659",
};

// データベースの中で、このアプリが使う置き場所（豆知識帳などと混ざらないように分けます）
export const ROOT_PATH = "sectionApp";

// Firebase SDK のバージョン
export const SDK_VERSION = "10.12.2";
