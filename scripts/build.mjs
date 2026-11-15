// 公開用（GitHub Pages）の中身を dist/ に作る。
// ・HTML/CSS/画像/データなどはそのままコピー
// ・.js だけ、難読化してから書き出す（中身は同じ、読みにくくするだけ）
// 元のファイル（リポジトリの中身）には一切手を触れない
import { readdirSync, statSync, mkdirSync, copyFileSync, readFileSync, writeFileSync, rmSync, existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JavaScriptObfuscator from "javascript-obfuscator";

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const OUT = path.join(ROOT, "dist");
const EXCLUDE = new Set([".git", "node_modules", "dist", "scripts", ".github", "package.json", "package-lock.json", "README.md", "docs"]);

function copyDir(src, dst) {
  mkdirSync(dst, { recursive: true });
  for (const name of readdirSync(src)) {
    if (EXCLUDE.has(name)) continue;
    const s = path.join(src, name);
    const d = path.join(dst, name);
    const st = statSync(s);
    if (st.isDirectory()) copyDir(s, d);
    else copyFileSync(s, d);
  }
}

function obfuscateJsFiles(dir) {
  for (const name of readdirSync(dir)) {
    const p = path.join(dir, name);
    const st = statSync(p);
    if (st.isDirectory()) {
      obfuscateJsFiles(p);
      continue;
    }
    if (!name.endsWith(".js")) continue;
    const src = readFileSync(p, "utf8");
    const result = JavaScriptObfuscator.obfuscate(src, {
      compact: true,
      controlFlowFlattening: false,
      deadCodeInjection: false,
      stringArray: true,
      stringArrayThreshold: 0.75,
      rotateStringArray: true,
      identifierNamesGenerator: "hexadecimal",
      renameGlobals: false,
      selfDefending: false,
      target: "browser",
    });
    writeFileSync(p, result.getObfuscatedCode());
  }
}

if (existsSync(OUT)) rmSync(OUT, { recursive: true, force: true });
copyDir(ROOT, OUT);
obfuscateJsFiles(OUT);
console.log("build done:", OUT);
