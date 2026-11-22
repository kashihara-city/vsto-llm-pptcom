// PDFs, with PDF.js (anydoc is not used for them: it refuses a whole PDF when a single page is a scan,
// and its text has no page boundaries). Used by docs.js:
//   convertPdf(bytes)  -> { ok, format: "pdf", text, images: [{index, mediaType, data, textlessPage}], pages, textlessPages }
//   renderPages(bytes, pages, maxSide) -> [{ page, data (PNG, base64) }]   for docread page
//
// The text is per page: a "[page 3]" line, then the page's text lines as PDF.js reads them, then a
// [[img:N]] line for each image drawn on that page with where it sits ("upper left" ...). A page with
// (almost) no text gets a note instead: it is a scan or a picture, and the model looks at it with
// docread page. Text from PDFs can be incomplete or garbled (charts, fonts without a character map),
// which is what the page renderings are for. The add-in is told how many pages have no text and which
// images sit on them (textlessPages, textlessPage): reading those pages means looking at each one, which
// it counts when it sizes what is attached.
import * as pdfjs from "./vendor/pdfjs/pdf.mjs";

pdfjs.GlobalWorkerOptions.workerSrc = new URL("vendor/pdfjs/pdf.worker.mjs", import.meta.url).href;
const assets = new URL("vendor/pdfjs/", import.meta.url).href;

// Images smaller than this (icons, rules, pieces of a layout) are left out, and at most MaxImages are kept. One whose
// long side is over MaxSide is taken smaller: a scan at a printer's resolution has tens of megapixels, and as a PNG
// of its own size it took the page's memory and the user's time, for a picture that is shown far smaller.
const MinSide = 64, MinArea = 120 * 120, MaxImages = 60, MaxSide = 4000;

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
    let textlessPages = 0;
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      out.push(`[page ${p}]`);
      const lines = textLines(await page.getTextContent());
      const chars = lines.reduce((n, l) => n + l.replace(/\s/g, "").length, 0);
      const textless = chars < 20;
      if (textless) {
        textlessPages++;
        out.push(`(page ${p} has ${chars ? "almost " : ""}no text: a scan or a picture - look at it with docread page)`);
      }
      for (const l of lines) out.push(l);
      if (images.length < MaxImages) {
        // one at a time: once the document has as many images as it keeps, no further one is made into a PNG (a
        // page's images were all made first, and those over the number thrown away)
        for await (const img of pageImages(page)) {
          const key = await digest(img.data);
          if (seen.has(key)) continue;
          seen.add(key);
          out.push(`[[img:${images.length}]] (page ${p}, ${img.where})`);
          images.push({ index: images.length, mediaType: "image/png", data: img.data, textlessPage: textless });
          if (images.length >= MaxImages) break;
        }
      }
      page.cleanup();
    }
    return { ok: true, format: "pdf", text: out.join("\n"), images, pages: doc.numPages, textlessPages };
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
// each is painted), one after the other as they are asked for. The same image painted twice on a page counts once.
async function* pageImages(page) {
  const ops = await page.getOperatorList();
  const O = pdfjs.OPS;
  const view = page.view;   // [x0, y0, x1, y1] in PDF units
  const pw = view[2] - view[0], ph = view[3] - view[1];
  const ids = new Set();
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
      yield { data, where: describe(cx, cy, Math.abs(ctm[0] * ctm[3]) / (pw * ph)) };
    }
  }
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

// PDF.js hands decoded images either as an ImageBitmap or as raw pixels (kind 1: 1 bit gray, 2: RGB, 3: RGBA). The
// PNG's long side is at most MaxSide: a larger image is made smaller.
async function toPng(img) {
  const scale = Math.min(1, MaxSide / Math.max(img.width, img.height));
  const w = Math.max(1, Math.round(img.width * scale)), h = Math.max(1, Math.round(img.height * scale));
  const canvas = new OffscreenCanvas(w, h);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, w, h);
  if (img.bitmap) ctx.drawImage(img.bitmap, 0, 0, w, h);
  else if (img.data) {
    const rgba = pixels(img, w, h);
    if (!rgba) return null;
    ctx.putImageData(new ImageData(rgba, w, h), 0, 0);
  }
  else return null;
  const blob = await canvas.convertToBlob({ type: "image/png" });
  return new Uint8Array(await blob.arrayBuffer());
}

// An image's raw pixels as RGBA, w x h: its own size, or smaller - each pixel is then the average of the pixels of
// the image it stands for, so that nothing as large as the image is made on the way. A pixel's color counts by how
// much of it shows (its alpha): a transparent pixel has a color too, black as a rule, and averaged in as it is, it
// darkened the edge of what is next to it. Null for a kind of raw data that is none of PDF.js's three.
function pixels(img, w, h) {
  const src = img.data, sw = img.width, sh = img.height, kind = img.kind;
  if (kind !== 1 && kind !== 2 && kind !== 3) return null;
  const out = new Uint8ClampedArray(w * h * 4), rowBytes = (sw + 7) >> 3;
  if (kind === 3 && w === sw && h === sh) { out.set(src.subarray(0, out.length)); return out; }
  for (let y = 0; y < h; y++) {
    const y0 = Math.floor(y * sh / h), y1 = Math.max(y0 + 1, Math.floor((y + 1) * sh / h));
    for (let x = 0; x < w; x++) {
      const x0 = Math.floor(x * sw / w), x1 = Math.max(x0 + 1, Math.floor((x + 1) * sw / w));
      let r = 0, g = 0, b = 0, a = 0;   // the colors' sums, each weighed by its pixel's alpha, and the alphas' sum
      for (let sy = y0; sy < y1; sy++) for (let sx = x0; sx < x1; sx++) {
        if (kind === 3) { const k = (sy * sw + sx) * 4, shows = src[k + 3]; r += src[k] * shows; g += src[k + 1] * shows; b += src[k + 2] * shows; a += shows; }
        else if (kind === 2) { const k = (sy * sw + sx) * 3; r += src[k] * 255; g += src[k + 1] * 255; b += src[k + 2] * 255; a += 255; }
        else { const v = (src[sy * rowBytes + (sx >> 3)] >> (7 - (sx & 7))) & 1 ? 255 : 0; r += v * 255; g += v * 255; b += v * 255; a += 255; }
      }
      const n = (y1 - y0) * (x1 - x0), k = (y * w + x) * 4;
      if (a > 0) { out[k] = r / a; out[k + 1] = g / a; out[k + 2] = b / a; }
      out[k + 3] = a / n;
    }
  }
  return out;
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
