// Documents the user attaches to the conversation: dropped on the pane, picked with the attach button,
// or sent by the add-in (a file attached from outside, AddinAutomation.Attach). Each one is converted -
// PDFs with PDF.js (pdfdoc.js), everything else in a worker with anydoc (docworker.js) - and handed to the
// add-in, which keeps it for the conversation and lets the model read it (docread). The add-in owns the
// list; this page shows what it says (the chips above the input) and asks it to remove one. It also draws
// PDF pages for docread page; the add-in sends the PDF along each time, so this page keeps no documents.
//
//   page -> add-in: {type: "attachStarted", ticket, name}                   converting starts (the add-in ties it to the conversation)
//                   {type: "attach", ticket, name, size, format, text, pages, images: [{index, mediaType, data (base64)}],
//                    source (the PDF itself, base64; PDFs only)}
//                   {type: "attachFailed", ticket, name, reason, detail}
//                   {type: "detach", id}
//                   {type: "pagesRendered", request, pages: [{page, data (PNG, base64) | error}]} or {..., error}
//   add-in -> page: {type: "convert", ticket, name, data (base64)}   a file attached from outside
//                   {type: "renderPages", request, data (the PDF, base64), pages: [n], maxSide}
//                   {type: "attachments", items: [{id, name, detail}]}   (detail: a message, see i18n.js)
//                   {type: "busy", value}                             (also handled by taskpane.js)
const host = window.chrome && window.chrome.webview;
const $ = id => document.getElementById(id);

// Larger files are refused before reading them (the add-in applies the same limit to Attach).
const MaxBytes = 50 * 1024 * 1024;
const Accept = [".docx", ".doc", ".docm", ".xlsx", ".xlsm", ".pptx", ".ppt", ".pptm", ".pdf", ".odt", ".ods", ".odp",
                ".rtf", ".epub", ".csv", ".txt", ".md"];

let busy = false;
let converting = 0;
let nextTicket = 1;

const area = $("attach-area"), chips = $("attachments"), picker = $("attach-file"), attachButton = $("attach");
picker.accept = Accept.join(",");

// ---- one conversion at a time, in order --------------------------------------------------------------

let queue = Promise.resolve();
function enqueue(ticket, name, bytes) {
  host.postMessage({ type: "attachStarted", ticket, name });
  converting++;
  updateState();
  queue = queue.then(() => convert(ticket, name, bytes)).finally(() => { converting--; updateState(); });
}

const isPdf = bytes => bytes.length > 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46; // %PDF

async function convert(ticket, name, bytes) {
  const size = bytes.byteLength;
  let r;
  try {
    if (isPdf(bytes)) r = await (await import("./pdfdoc.js")).convertPdf(bytes);
    else r = await inWorker(name, bytes.slice());
  } catch (e) {
    r = { ok: false, reason: "error", detail: String((e && e.message) || e) };
  }
  if (!r.ok) {
    host.postMessage({ type: "attachFailed", ticket, name, reason: r.reason, detail: r.detail || "" });
    return;
  }
  const images = [];
  for (const img of r.images) images.push({ index: img.index, mediaType: img.mediaType, data: await base64(img.data) });
  const msg = { type: "attach", ticket, name, size, format: r.format, text: r.text, images };
  if (r.format === "pdf") { msg.pages = r.pages; msg.source = await base64(bytes); }
  host.postMessage(msg);
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

async function addFiles(files) {
  if (busy) return;
  for (const file of files) {
    const ticket = "p" + nextTicket++;
    if (file.size > MaxBytes) {
      host.postMessage({ type: "attachFailed", ticket, name: file.name, reason: "tooLarge", detail: String(file.size) });
      continue;
    }
    enqueue(ticket, file.name, new Uint8Array(await file.arrayBuffer()));
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

// ---- chips -----------------------------------------------------------------------------------------

function showAttachments(items) {
  chips.textContent = "";
  for (const it of items) {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.title = it.detail ? i18n.t(it.detail) : "";
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
    else if (m.type === "convert") enqueue(m.ticket, m.name, fromBase64(m.data));
    else if (m.type === "renderPages") renderPages(m);
  });
}
updateState();
