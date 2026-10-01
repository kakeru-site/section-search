// 画像の読み込み・縮小まわり
// DBに直接入れるので、保存前にサイズをかなり落としている

const FULL_MAX = 1600; // 長辺px
const THUMB_MAX = 220;

let pdfjs = null;

function loadPdfJs() {
  if (pdfjs) return pdfjs;
  pdfjs = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => {
      const lib = window.pdfjsLib;
      lib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(lib);
    };
    s.onerror = () => {
      pdfjs = null;
      reject(new Error("PDFの読み込み部品を取得できませんでした"));
    };
    document.head.appendChild(s);
  });
  return pdfjs;
}

function fileToImage(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("画像として開けませんでした"));
    };
    img.src = url;
  });
}

// webpが使えればwebp、だめならjpeg
function toDataUrl(canvas, q) {
  const webp = canvas.toDataURL("image/webp", q);
  if (webp.startsWith("data:image/webp")) return webp;
  return canvas.toDataURL("image/jpeg", q);
}

function shrink(src, max, q) {
  const w0 = src.naturalWidth || src.width;
  const h0 = src.naturalHeight || src.height;
  const r = Math.min(1, max / Math.max(w0, h0));
  const w = Math.max(1, Math.round(w0 * r));
  const h = Math.max(1, Math.round(h0 * r));
  const c = document.createElement("canvas");
  c.width = w;
  c.height = h;
  const ctx = c.getContext("2d");
  ctx.fillStyle = "#fff"; // 透過PNGが黒くならないように
  ctx.fillRect(0, 0, w, h);
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(src, 0, 0, w, h);
  return { url: toDataUrl(c, q), w, h };
}

export function makeImages(src) {
  const full = shrink(src, FULL_MAX, 0.82);
  const thumb = shrink(src, THUMB_MAX, 0.7);
  return { full: full.url, thumb: thumb.url, w: full.w, h: full.h };
}

// PDFを開いて、ページを描画できるようにする
export async function openPdf(file) {
  const lib = await loadPdfJs();
  const doc = await lib.getDocument({ data: await file.arrayBuffer() }).promise;
  return {
    pages: doc.numPages,
    async render(n) {
      const page = await doc.getPage(n);
      const v1 = page.getViewport({ scale: 1 });
      const scale = Math.min(3, FULL_MAX / Math.max(v1.width, v1.height));
      const vp = page.getViewport({ scale });
      const c = document.createElement("canvas");
      c.width = Math.round(vp.width);
      c.height = Math.round(vp.height);
      const ctx = c.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      return c;
    },
  };
}

// ドロップ／貼り付けされたものを判定
export function kindOf(file) {
  const name = (file.name || "").toLowerCase();
  if (file.type.startsWith("image/")) return "image";
  if (file.type === "application/pdf" || name.endsWith(".pdf")) return "pdf";
  if (name.endsWith(".xdw") || name.endsWith(".xbd")) return "docuworks";
  return "other";
}

export { fileToImage };

export function sizeKB(dataUrl) {
  return Math.round((dataUrl.length * 3) / 4 / 1024);
}
