// PDFs, with PDF.js (anydoc is not used for them: it refuses a whole PDF when a single page is a scan,
// and its text has no page boundaries). Used by docs.js:
//   convertPdf(bytes)  -> { ok, format: "pdf", text, images: [{index, mediaType, data}], pages }
//   renderPages(bytes, pages, maxSide) -> [{ page, data (PNG, base64) }]   for docread page
//
// The text is per page: a "[page 3]" line, then the page's text lines as PDF.js reads them, then a
// [[img:N]] line for each image drawn on that page with where it sits ("upper left" ...). A page with
// (almost) no text gets a note instead: it is a scan or a picture, and the model looks at it with
// docread page. Text from PDFs can be incomplete or garbled (charts, fonts without a character map),
// which is what the page renderings are for.
import * as pdfjs from "./vendor/pdfjs/pdf.mjs";

pdfjs.GlobalWorkerOptions.workerSrc = new URL("vendor/pdfjs/pdf.worker.mjs", import.meta.url).href;
const assets = new URL("vendor/pdfjs/", import.meta.url).href;

// Images smaller than this (icons, rules, pieces of a layout) are left out, and at most MaxImages are kept.
const MinSide = 64, MinArea = 120 * 120, MaxImages = 60;

function open(bytes) {
  // PDF.js takes ownership of the buffer it is given: pass a copy.
  return pdfjs.getDocument({
    data: bytes.slice(),
    cMapUrl: assets + "cmaps/", cMapPacked: true,
    wasmUrl: assets + "wasm/", iccUrl: assets + "iccs/",
  }).promise;
}

export async function convertPdf(bytes) {
  let doc;
  try { doc = await open(bytes); }
  catch (e) {
    const reason = e && e.name === "PasswordException" ? "encrypted" : "malformed";
    return { ok: false, reason, detail: String((e && e.message) || e) };
  }
  try {
    const out = [], images = [], seen = new Set();
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      out.push(`[page ${p}]`);
      const lines = textLines(await page.getTextContent());
      const chars = lines.reduce((n, l) => n + l.replace(/\s/g, "").length, 0);
      if (chars < 20) out.push(`(page ${p} has ${chars ? "almost " : ""}no text: a scan or a picture - look at it with docread page)`);
      for (const l of lines) out.push(l);
      if (images.length < MaxImages) {
        for (const img of await pageImages(page)) {
          if (images.length >= MaxImages) break;
          const key = await digest(img.data);
          if (seen.has(key)) continue;
          seen.add(key);
          out.push(`[[img:${images.length}]] (page ${p}, ${img.where})`);
          images.push({ index: images.length, mediaType: "image/png", data: img.data });
        }
      }
      page.cleanup();
    }
    return { ok: true, format: "pdf", text: out.join("\n"), images, pages: doc.numPages };
  } finally {
    await doc.loadingTask.destroy();   // ends the worker's copy of the document too
  }
}

// Lines in reading order as the PDF stores them: an item ends a line when PDF.js says so (hasEOL) or
// the next one sits on another baseline; items on one line are joined, with a space where there is a gap.
function textLines(content) {
  const lines = [];
  let line = "", lastY = null, lastEnd = null, lastH = 0;
  const flush = () => { const t = line.replace(/\s+/g, " ").trim(); if (t) lines.push(t); line = ""; lastEnd = null; };
  for (const it of content.items) {
    if (typeof it.str !== "string") continue;
    const x = it.transform[4], y = it.transform[5], h = Math.abs(it.transform[3]) || it.height || 10;
    if (lastY !== null && Math.abs(y - lastY) > Math.max(h, lastH) * 0.5) flush();
    if (lastEnd !== null && x - lastEnd > h * 0.3 && !line.endsWith(" ")) line += " ";
    line += it.str;
    lastY = y; lastH = h; lastEnd = x + (it.width || 0);
    if (it.hasEOL) { flush(); lastY = null; }
  }
  flush();
  return lines;
}

// The images painted on the page, as PNG bytes, with where they sit (from the transform in effect when
// each is painted). The same image painted twice on a page counts once.
async function pageImages(page) {
  const ops = await page.getOperatorList();
  const O = pdfjs.OPS;
  const view = page.view;   // [x0, y0, x1, y1] in PDF units
  const pw = view[2] - view[0], ph = view[3] - view[1];
  const found = [], ids = new Set();
  const stack = [];
  let ctm = [1, 0, 0, 1, 0, 0];
  const mul = (a, b) => [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3],
                         a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];
  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i], args = ops.argsArray[i];
    if (fn === O.save) stack.push(ctm);
    else if (fn === O.restore) ctm = stack.pop() || ctm;
    else if (fn === O.transform) ctm = mul(ctm, args);
    else if (fn === O.paintImageXObject && typeof args[0] === "string" && !ids.has(args[0])) {
      ids.add(args[0]);
      const obj = await objectOf(page, args[0]);
      if (!obj || obj.width < MinSide || obj.height < MinSide || obj.width * obj.height < MinArea) continue;
      const data = await toPng(obj);
      if (!data) continue;
      // the unit square mapped by the transform: its center, relative to the page
      const cx = (ctm[4] + (ctm[0] + ctm[2]) / 2 - view[0]) / pw;
      const cy = 1 - (ctm[5] + (ctm[1] + ctm[3]) / 2 - view[1]) / ph;
      found.push({ data, where: describe(cx, cy, Math.abs(ctm[0] * ctm[3]) / (pw * ph)) });
    }
  }
  return found;
}

function objectOf(page, id) {
  const store = id.startsWith("g_") ? page.commonObjs : page.objs;
  return new Promise(resolve => {
    try { store.get(id, resolve); } catch { resolve(null); }
    setTimeout(() => resolve(null), 5000);
  });
}

function describe(cx, cy, area) {
  const v = cy < 0.34 ? "upper" : cy > 0.66 ? "lower" : "middle";
  const h = cx < 0.34 ? " left" : cx > 0.66 ? " right" : "";
  const size = area > 0.4 ? "most of the page" : area > 0.12 ? "large" : area > 0.03 ? "medium" : "small";
  return `${v === "middle" && !h ? "center" : v + h}, ${size}`;
}

// PDF.js hands decoded images either as an ImageBitmap or as raw pixels (kind 1: 1 bit gray, 2: RGB, 3: RGBA).
async function toPng(img) {
  const canvas = new OffscreenCanvas(img.width, img.height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, img.width, img.height);
  if (img.bitmap) ctx.drawImage(img.bitmap, 0, 0);
  else if (img.data) {
    const rgba = new Uint8ClampedArray(img.width * img.height * 4);
    const src = img.data, n = img.width * img.height;
    if (img.kind === 3) rgba.set(src.subarray(0, rgba.length));
    else if (img.kind === 2) for (let i = 0, j = 0; i < n; i++, j += 3) { rgba[i * 4] = src[j]; rgba[i * 4 + 1] = src[j + 1]; rgba[i * 4 + 2] = src[j + 2]; rgba[i * 4 + 3] = 255; }
    else if (img.kind === 1) {
      const rowBytes = (img.width + 7) >> 3;
      for (let y = 0; y < img.height; y++) for (let x = 0; x < img.width; x++) {
        const v = (src[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0, k = (y * img.width + x) * 4;
        rgba[k] = rgba[k + 1] = rgba[k + 2] = v; rgba[k + 3] = 255;
      }
    }
    else return null;
    ctx.putImageData(new ImageData(rgba, img.width, img.height), 0, 0);
  }
  else return null;
  const blob = await canvas.convertToBlob({ type: "image/png" });
  return new Uint8Array(await blob.arrayBuffer());
}

async function digest(bytes) {
  const h = new Uint8Array(await crypto.subtle.digest("SHA-256", bytes));
  return Array.from(h.subarray(0, 12), b => b.toString(16).padStart(2, "0")).join("");
}

// Pages drawn as the reader sees them, the long side at most maxSide pixels, on white.
export async function renderPages(bytes, pages, maxSide) {
  const doc = await open(bytes);
  try {
    const out = [];
    for (const p of pages) {
      if (p < 1 || p > doc.numPages) { out.push({ page: p, error: `the PDF has ${doc.numPages} pages` }); continue; }
      const page = await doc.getPage(p);
      const base = page.getViewport({ scale: 1 });
      const viewport = page.getViewport({ scale: maxSide / Math.max(base.width, base.height) });
      const canvas = new OffscreenCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
      const blob = await canvas.convertToBlob({ type: "image/png" });
      out.push({ page: p, blob });
      page.cleanup();
    }
    return { pages: doc.numPages, out };
  } finally {
    await doc.loadingTask.destroy();   // ends the worker's copy of the document too
  }
}
