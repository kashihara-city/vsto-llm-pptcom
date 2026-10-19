// How the task pane takes the documents the user attaches (pptcom/UI/web/docs.js): every file is announced to the
// add-in (attachStarted) and counted as converting before it is read, and every announced file ends with attach
// or attachFailed, whatever fails on the way, while the files after it go on. A picture goes as it is (format
// "image"), not through the converter; a paste attaches the clipboard's files only when it holds no text.
// docs.js runs as it is, in a stand-in for the browser and WebView2 made of plain objects (it imports nothing at
// its top level, and the stand-in has no createImageBitmap, so a picture comes without its chip thumbnail); no packages.
//
//   node scripts/test/docs-queue.cjs     exit code 0 = all passed
"use strict";
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "..", "..", "pptcom", "UI", "web", "docs.js"), "utf8");
let passed = 0, failed = 0;

function check(name, ok, detail) {
  if (ok) { passed++; return; }
  failed++;
  console.log("FAIL " + name);
  console.log("     " + detail);
}

// A converter that answers at once: a text document, with one image when the file's name says so.
class TextWorker {
  static started = 0;
  constructor() { TextWorker.started++; }
  postMessage({ name }) {
    const images = name.includes("image") ? [{ index: 0, mediaType: "image/png", data: new Uint8Array([1]) }] : [];
    queueMicrotask(() => this.onmessage({ data: { ok: true, format: "text", text: "x", images } }));
  }
  terminate() {}
}

// FileReader.readAsDataURL, as far as docs.js uses it; `fails` makes every read fail.
function fileReader(fails) {
  return class {
    readAsDataURL(blob) {
      if (fails) { this.error = new Error("read failed"); queueMicrotask(() => this.onerror()); return; }
      blob.arrayBuffer().then(b => { this.result = "data:x;base64," + Buffer.from(b).toString("base64"); this.onload(); });
    }
  };
}

// docs.js in a fresh page. `messages` is what it posted to the add-in; `fromAddin` delivers a message to it.
function page({ worker = TextWorker, readerFails = false } = {}) {
  const messages = [], classes = new Set(), elements = new Map(), onDocument = {};
  let onMessage = null;
  const element = () => ({
    children: [], classList: { toggle() {} }, addEventListener() {}, querySelectorAll: () => [], appendChild() {},
  });
  const context = vm.createContext({
    console, Uint8Array, Blob, atob, queueMicrotask,
    window: { chrome: { webview: {
      postMessage: m => messages.push(m),
      addEventListener: (type, fn) => { if (type === "message") onMessage = fn; },
    } } },
    document: {
      body: { classList: { toggle: (k, on) => (on ? classes.add(k) : classes.delete(k)), add() {}, remove() {} } },
      getElementById: id => { if (!elements.has(id)) elements.set(id, element()); return elements.get(id); },
      addEventListener: (type, fn) => { onDocument[type] = fn; },
      createElement: element,
    },
    i18n: { t: s => s, fill: s => s },
    Worker: worker,
    FileReader: fileReader(readerFails),
  });
  vm.runInContext(source, context);
  return {
    messages,
    converting: () => classes.has("converting"),
    // a failure that escapes docs.js leaves its ticket without an end, which the checks report
    addFiles: files => {
      context.files = files;
      Promise.resolve(vm.runInContext("addFiles(files)", context)).catch(() => {});
    },
    fromAddin: data => onMessage({ data }),
    // Ctrl+V: the clipboard's text and files, into the request box or a settings field; returns whether the page
    // took it over (preventDefault)
    paste: ({ text = "", files = [], inSettings = false }) => {
      let prevented = false;
      onDocument.paste({
        clipboardData: { getData: type => (type === "text/plain" ? text : ""), files },
        target: { closest: selector => (inSettings && selector === "#settings" ? {} : null) },
        preventDefault: () => { prevented = true; },
      });
      return prevented;
    },
    settled: async () => { for (let i = 0; i < 5; i++) await vm.runInContext("queue", context).catch(() => {}); },
  };
}

const PngHeader = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0x0D];

// A dropped file whose bytes come when the test says (or fail to); "A" unless given.
function file(name, size = 1, content = [65]) {
  let resolve, reject, reads = 0;
  const bytes = new Promise((res, rej) => { resolve = res; reject = rej; });
  return {
    name, size,
    arrayBuffer: () => { reads++; return bytes; },
    arrive: () => resolve(new Uint8Array(content).buffer),
    fail: () => reject(new Error("the file could not be read")),
    reads: () => reads,
  };
}

const show = messages => messages.map(m => m.type + ":" + m.ticket).join(", ");

// Each ticket: one attachStarted, then exactly one attach or attachFailed.
function eachEndsOnce(messages) {
  const tickets = new Map();
  for (const m of messages) {
    if (!tickets.has(m.ticket)) tickets.set(m.ticket, []);
    tickets.get(m.ticket).push(m.type);
  }
  for (const types of tickets.values()) {
    const ends = types.filter(t => t === "attach" || t === "attachFailed").length;
    if (types[0] !== "attachStarted" || ends !== 1) return false;
  }
  return true;
}

async function announcedBeforeRead() {
  const p = page(), a = file("a.txt"), b = file("b.txt");
  p.addFiles([a, b]);
  check("every file is announced before any is read",
        show(p.messages) === "attachStarted:p1, attachStarted:p2" && p.converting() && a.reads() === 0 && b.reads() === 0,
        show(p.messages) + "; converting=" + p.converting() + "; reads=" + a.reads() + "," + b.reads());
  a.arrive();
  b.arrive();
  await p.settled();
  check("then each is attached, in order",
        show(p.messages) === "attachStarted:p1, attachStarted:p2, attach:p1, attach:p2" && !p.converting(),
        show(p.messages) + "; converting=" + p.converting());
}

async function unreadable() {
  const p = page(), a = file("a.txt"), b = file("b.txt");
  p.addFiles([a, b]);
  a.fail();
  b.arrive();
  await p.settled();
  const f = p.messages.find(m => m.type === "attachFailed");
  check("a file that cannot be read: attachFailed, and the next one goes on",
        eachEndsOnce(p.messages) && show(p.messages).endsWith("attachFailed:p1, attach:p2") && f.reason === "error" &&
        f.detail.includes("could not be read") && !p.converting(),
        show(p.messages) + "; converting=" + p.converting());
}

async function encodingFails() {
  const p = page({ readerFails: true }), a = file("image.docx"), b = file("b.txt");
  p.addFiles([a, b]);
  a.arrive();
  b.arrive();
  await p.settled();
  check("a failure while the result is encoded: attachFailed, and the next one goes on",
        eachEndsOnce(p.messages) && show(p.messages).endsWith("attachFailed:p1, attach:p2") && !p.converting(),
        show(p.messages) + "; converting=" + p.converting());
}

async function converterThrows() {
  let first = true;
  class Flaky extends TextWorker {
    constructor() { super(); if (first) { first = false; throw new Error("no worker"); } }
  }
  const p = page({ worker: Flaky }), a = file("a.txt"), b = file("b.txt");
  p.addFiles([a, b]);
  a.arrive();
  b.arrive();
  await p.settled();
  check("a converter that fails to start: attachFailed, and the next one goes on",
        eachEndsOnce(p.messages) && show(p.messages).endsWith("attachFailed:p1, attach:p2") && !p.converting(),
        show(p.messages) + "; converting=" + p.converting());
}

async function tooLarge() {
  const p = page(), big = file("big.pdf", 51 * 1024 * 1024);
  p.addFiles([big]);
  await p.settled();
  check("a file over the limit is refused without being read",
        show(p.messages) === "attachFailed:p1" && p.messages[0].reason === "tooLarge" && big.reads() === 0 && !p.converting(),
        show(p.messages) + "; reads=" + big.reads());
}

async function fromAddin() {
  const p = page();
  p.fromAddin({ type: "convert", ticket: "a1", name: "outside.txt", data: Buffer.from("A").toString("base64") });
  check("a file sent by the add-in is announced at once", show(p.messages) === "attachStarted:a1" && p.converting(),
        show(p.messages));
  await p.settled();
  check("and attached", show(p.messages) === "attachStarted:a1, attach:a1" && !p.converting(), show(p.messages));
}

async function picture() {
  const before = TextWorker.started;
  const p = page(), png = file("image.png", PngHeader.length, PngHeader);
  p.addFiles([png]);
  png.arrive();
  await p.settled();
  const m = p.messages.find(x => x.type === "attach");
  check("a picture is attached as it is, as one image and no text, without the converter",
        eachEndsOnce(p.messages) && m && m.format === "image" && m.text === "" && m.images.length === 1 &&
        m.images[0].mediaType === "image/png" && m.images[0].data === Buffer.from(PngHeader).toString("base64") &&
        m.thumb === "" && TextWorker.started === before && !p.converting(),
        JSON.stringify(m) + "; converters started=" + (TextWorker.started - before));
}

async function textStartingLikeBmp() {
  const p = page(), txt = file("notes.txt", 3, [0x42, 0x4D, 0x41]); // "BMA"
  p.addFiles([txt]);
  txt.arrive();
  await p.settled();
  const m = p.messages.find(x => x.type === "attach");
  check("a text file that starts with BM is a document, not a picture", m && m.format === "text", JSON.stringify(m));
}

async function paste() {
  let p = page();
  let prevented = p.paste({ text: "1\r\n2\r\n", files: [file("image.png")] });
  check("a paste with text (Excel puts a picture next to it) is left to the text box",
        !prevented && p.messages.length === 0, "prevented=" + prevented + "; " + show(p.messages));

  p = page();
  const snip = file("image.png", PngHeader.length, PngHeader);
  prevented = p.paste({ files: [snip] });
  snip.arrive();
  await p.settled();
  check("a paste of a picture alone attaches it",
        prevented && show(p.messages) === "attachStarted:p1, attach:p1" && p.messages[1].format === "image",
        "prevented=" + prevented + "; " + show(p.messages));

  p = page();
  prevented = p.paste({ files: [file("image.png")], inSettings: true });
  check("a paste into the settings is left alone", !prevented && p.messages.length === 0,
        "prevented=" + prevented + "; " + show(p.messages));

  p = page();
  prevented = p.paste({});
  check("a paste with nothing on the clipboard is left alone", !prevented && p.messages.length === 0,
        "prevented=" + prevented + "; " + show(p.messages));
}

(async () => {
  await announcedBeforeRead();
  await unreadable();
  await encodingFails();
  await converterThrows();
  await tooLarge();
  await fromAddin();
  await picture();
  await textStartingLikeBmp();
  await paste();
  console.log();
  console.log(failed === 0 ? `ALL PASSED (${passed})` : `${failed} FAILED, ${passed} passed`);
  process.exitCode = failed === 0 ? 0 : 1;
})().catch(e => { console.error(e); process.exitCode = 1; });
