// Documents the user attaches to the conversation: dropped on the pane, picked with the attach button, pasted,
// or sent by the add-in (a file attached from outside, AddinAutomation.Attach). Each one is converted -
// PDFs with PDF.js (pdfdoc.js), everything else in a worker with anydoc (docworker.js) - and handed to the
// add-in, which keeps it for the conversation and lets the model read it (docread). A picture is not converted:
// it goes as format "image", one image and no text, and the model sees it with the request that follows. The
// add-in owns the list; this page shows what it says (the chips above the input) and asks it to remove one. It
// also draws PDF pages for docread page; the add-in sends the PDF along each time, so this page keeps no documents.
//
//   page -> add-in: {type: "attachStarted", ticket, name}                   a file is taken, before it is read (the add-in ties it to the conversation)
//                   {type: "attach", ticket, name, size, format, text, pages, images: [{index, mediaType, data (base64)}],
//                    source (the PDF itself, base64; PDFs only), thumb (a small PNG as a data: URL; pictures only)}
//                   {type: "attachFailed", ticket, name, reason, detail}
//                   {type: "detach", id}
//                   {type: "pagesRendered", request, pages: [{page, data (PNG, base64) | error}]} or {..., error}
//   add-in -> page: {type: "convert", ticket, name, data (base64)}   a file attached from outside
//                   {type: "renderPages", request, data (the PDF, base64), pages: [n], maxSide}
//                   {type: "attachments", items: [{id, name, detail, thumb}]}   (detail: a message, see i18n.js; thumb: "" or a data: URL)
//                   {type: "busy", value}                             (also handled by taskpane.js)
const host = window.chrome && window.chrome.webview;
const $ = id => document.getElementById(id);

// Larger files are refused before reading them (the add-in applies the same limit to Attach).
const MaxBytes = 50 * 1024 * 1024;
const Accept = [".docx", ".doc", ".docm", ".xlsx", ".xlsm", ".pptx", ".ppt", ".pptm", ".pdf", ".odt", ".ods", ".odp",
                ".rtf", ".epub", ".csv", ".txt", ".md", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"];
// Long side of a chip's picture, in pixels (shown at about half that, sharp on a high-DPI screen).
const ThumbSide = 96;

let busy = false;
let converting = 0;
let nextTicket = 1;

const area = $("attach-area"), chips = $("attachments"), picker = $("attach-file"), attachButton = $("attach");
picker.accept = Accept.join(",");

// ---- one conversion at a time, in order --------------------------------------------------------------

// A file is announced (attachStarted) and counted as converting as soon as it is taken, before it is read:
// reading one from a network or cloud drive can take a while, and meanwhile the request must not go without it
// nor the conversation change under it. Its bytes are read when its turn in the queue comes. Every ticket ends
// with attach or attachFailed, whatever fails on the way, and the queue goes on with the next file.
// read: () => a Promise of the file's bytes (a Uint8Array).
let queue = Promise.resolve();
function enqueue(ticket, name, read) {
  host.postMessage({ type: "attachStarted", ticket, name });
  converting++;
  updateState();
  queue = queue
    .then(() => attach(ticket, name, read))
    .finally(() => { converting--; updateState(); })
    .catch(e => console.error(e));
}

async function attach(ticket, name, read) {
  let msg;
  try {
    msg = await convert(ticket, name, await read());
  } catch (e) {
    msg = { type: "attachFailed", ticket, name, reason: "error", detail: String((e && e.message) || e) };
  }
  host.postMessage(msg);
}

const isPdf = bytes => bytes.length > 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46; // %PDF

// The message for the add-in: attach, or attachFailed when the converter says why it could not.
async function convert(ticket, name, bytes) {
  const size = bytes.byteLength;
  const picture = pictureType(bytes);
  if (picture) return pictureMessage(ticket, name, size, picture, bytes);
  const r = isPdf(bytes) ? await (await import("./pdfdoc.js")).convertPdf(bytes) : await inWorker(name, bytes.slice());
  if (!r.ok) return { type: "attachFailed", ticket, name, reason: r.reason, detail: r.detail || "" };
  const images = [];
  for (const img of r.images) images.push({ index: img.index, mediaType: img.mediaType, data: await base64(img.data) });
  const msg = { type: "attach", ticket, name, size, format: r.format, text: r.text, images };
  if (r.format === "pdf") { msg.pages = r.pages; msg.source = await base64(bytes); }
  return msg;
}

// ---- pictures ----------------------------------------------------------------------------------------

// The picture types PowerPoint places as they are (and the add-in can draw for the model), by their first bytes,
// and WebP, which this page turns into a PNG first. null for anything else.
function pictureType(b) {
  const at = (i, ...v) => b.length >= i + v.length && v.every((x, k) => b[i + k] === x);
  if (at(0, 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)) return "image/png";
  if (at(0, 0xFF, 0xD8, 0xFF)) return "image/jpeg";
  if (at(0, 0x47, 0x49, 0x46, 0x38)) return "image/gif";                               // GIF8
  if (at(0, 0x49, 0x49, 0x2A, 0x00) || at(0, 0x4D, 0x4D, 0x00, 0x2A)) return "image/tiff";
  if (at(0, 0x52, 0x49, 0x46, 0x46) && at(8, 0x57, 0x45, 0x42, 0x50)) return "image/webp"; // RIFF....WEBP
  // BM, then a DIB header of a size Windows writes: "BM" alone also starts plain text
  if (at(0, 0x42, 0x4D) && [12, 40, 52, 56, 64, 108, 124].includes(b[14]) && at(15, 0, 0, 0)) return "image/bmp";
  return null;
}

async function pictureMessage(ticket, name, size, mediaType, bytes) {
  let data = bytes;
  if (mediaType === "image/webp") {
    // PowerPoint 2019 does not place WebP
    try { data = await drawn(bytes, Infinity); } catch (e) { return { type: "attachFailed", ticket, name, reason: "unsupported", detail: "" }; }
    mediaType = "image/png";
  }
  let thumb = "";
  try { thumb = "data:image/png;base64," + await base64(await drawn(data, ThumbSide)); } catch (e) { /* TIFF: no chip picture */ }
  return { type: "attach", ticket, name, size, format: "image", text: "", images: [{ index: 0, mediaType, data: await base64(data) }], thumb };
}

// The picture drawn as a PNG Blob, its long side at most maxSide pixels.
async function drawn(data, maxSide) {
  const bitmap = await createImageBitmap(data instanceof Blob ? data : new Blob([data]));
  try {
    const scale = Math.min(1, maxSide / Math.max(bitmap.width, bitmap.height));
    const canvas = new OffscreenCanvas(Math.max(1, Math.round(bitmap.width * scale)), Math.max(1, Math.round(bitmap.height * scale)));
    canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    return await canvas.convertToBlob({ type: "image/png" });
  } finally {
    bitmap.close();
  }
}

function inWorker(name, bytes) {
  return new Promise(resolve => {
    const worker = new Worker("docworker.js", { type: "module" });
    const done = (msg) => { worker.terminate(); resolve(msg); };
    worker.onmessage = e => done(e.data);
    worker.onerror = e => done({ ok: false, reason: "error", detail: e.message || "worker failed" });
    worker.postMessage({ name, bytes }, [bytes.buffer]);
  });
}

async function renderPages(m) {
  try {
    const { out } = await (await import("./pdfdoc.js")).renderPages(fromBase64(m.data), m.pages || [], m.maxSide || 1600);
    const pages = [];
    for (const p of out) pages.push(p.blob ? { page: p.page, data: await base64(p.blob) } : { page: p.page, error: p.error });
    host.postMessage({ type: "pagesRendered", request: m.request, pages });
  } catch (e) {
    host.postMessage({ type: "pagesRendered", request: m.request, error: String((e && e.message) || e) });
  }
}

// bytes: a Uint8Array or a Blob
function base64(bytes) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).replace(/^data:[^,]*,/, ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(bytes instanceof Blob ? bytes : new Blob([bytes]));
  });
}

function fromBase64(s) {
  const bin = atob(s);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function addFiles(files) {
  if (busy) return;
  for (const file of files) {
    const ticket = "p" + nextTicket++;
    if (file.size > MaxBytes) {
      host.postMessage({ type: "attachFailed", ticket, name: file.name, reason: "tooLarge", detail: String(file.size) });
      continue;
    }
    enqueue(ticket, file.name, async () => new Uint8Array(await file.arrayBuffer()));
  }
}

// ---- drop, attach button -----------------------------------------------------------------------------

const hasFiles = e => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
let dragDepth = 0;
document.addEventListener("dragenter", e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  if (dragDepth++ === 0 && !busy) document.body.classList.add("dropping");
});
document.addEventListener("dragleave", e => {
  if (!hasFiles(e)) return;
  if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove("dropping"); }
});
document.addEventListener("dragover", e => {
  if (!hasFiles(e)) return;
  e.preventDefault();          // also keeps WebView2 from opening the file itself
  e.dataTransfer.dropEffect = busy ? "none" : "copy";
});
document.addEventListener("drop", e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove("dropping");
  addFiles(Array.from(e.dataTransfer.files));
});

attachButton.addEventListener("click", () => { if (!busy) picker.click(); });
picker.addEventListener("change", () => { const files = Array.from(picker.files); picker.value = ""; addFiles(files); });

// ---- paste -------------------------------------------------------------------------------------------

// Ctrl+V with text on the clipboard pastes the text, as always: Excel and Word put a picture of what was copied
// next to its text, and whoever copies text means the text. Without text, the clipboard's files are attached like
// dropped ones: a screen snip, a slide or shape copied in PowerPoint (it gives no text for those), a picture copied
// in a browser, files copied in Explorer. The settings' fields keep the browser's own paste.
document.addEventListener("paste", e => {
  const dt = e.clipboardData;
  if (!dt || (e.target && e.target.closest && e.target.closest("#settings"))) return;
  if (dt.getData("text/plain").length > 0) return;
  const files = Array.from(dt.files || []);
  if (files.length === 0) return;
  e.preventDefault();
  addFiles(files);
});

// ---- chips -----------------------------------------------------------------------------------------

function showAttachments(items) {
  chips.textContent = "";
  for (const it of items) {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.title = it.detail ? i18n.t(it.detail) : "";
    if (it.thumb) {
      const img = document.createElement("img");
      img.className = "chip-thumb";
      img.alt = "";
      img.src = it.thumb;
      chip.appendChild(img);
    }
    const label = document.createElement("span");
    label.className = "chip-name";
    label.textContent = it.id + " " + it.name;
    chip.appendChild(label);
    const x = document.createElement("button");
    x.className = "chip-remove";
    x.title = i18n.t("DOC_REMOVE");
    x.setAttribute("aria-label", i18n.fill("DOC_REMOVE_NAMED", it.name));
    x.textContent = "×";
    x.addEventListener("click", () => { if (!busy) host.postMessage({ type: "detach", id: it.id }); });
    chip.appendChild(x);
    chips.appendChild(chip);
  }
}

function updateState() {
  attachButton.disabled = busy;
  for (const b of chips.querySelectorAll(".chip-remove")) b.disabled = busy;
  $("converting").textContent = converting > 0 ? i18n.t("DOC_CONVERTING") : "";
  // taskpane.js keeps the send and new-chat buttons off meanwhile: the request would go without the document.
  document.body.classList.toggle("converting", converting > 0);
  area.classList.toggle("show", chips.children.length > 0 || converting > 0);
}

if (host) {
  host.addEventListener("message", e => {
    const m = e.data || {};
    if (m.type === "attachments") { showAttachments(m.items || []); updateState(); }
    else if (m.type === "busy") { busy = !!m.value; updateState(); }
    else if (m.type === "convert") enqueue(m.ticket, m.name, async () => fromBase64(m.data));
    else if (m.type === "renderPages") renderPages(m);
  });
}
updateState();
