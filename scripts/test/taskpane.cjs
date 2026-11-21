// How the task pane's page (pptcom/UI/web/taskpane.js) sends a request and shows how it ends: one request at a
// time, also before the add-in's busy state arrives and while another window's pane runs one; a refused send does
// not leave the page waiting; a request that stops (the deck was closed, 中止, an error) closes its block as stopped,
// one that finishes as done; that what ran of a stopped script stays on the slides is told below the block, where it
// is seen; the conversation follows its end until the user scrolls up, and again once they are back at the end;
// 続ける shows last after a request that stopped at the turn limit, and sends like
// a request; the line above the input shows what is selected in PowerPoint, and the settings the version, whether
// the browser uses the add-in's refusing proxy, the context limit, the models (none chosen yet, 使わない), the image
// prompt and the API key (typed in, deleted).
// taskpane.js runs as it is, in a stand-in for the browser and WebView2 made of plain objects; messages
// from the add-in arrive as WebView2 delivers them. No packages.
//
//   node scripts/test/taskpane.cjs     exit code 0 = all passed
"use strict";
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "..", "..", "pptcom", "UI", "web", "taskpane.js"), "utf8");
let passed = 0, failed = 0;

function check(name, ok, detail) {
  if (ok) { passed++; return; }
  failed++;
  console.log("FAIL " + name);
  console.log("     " + detail);
}

function element() {
  const classes = new Set();
  let text = "";
  return {
    children: [], parentNode: null, value: "", scrollHeight: 0, scrollTop: 0, clientHeight: 0, disabled: false, listeners: {},
    // as in a browser, setting the text replaces what the element held
    get textContent() { return text; },
    set textContent(s) { text = s; this.children.length = 0; },
    set className(s) { classes.clear(); for (const c of s.split(/\s+/)) if (c) classes.add(c); },
    get className() { return [...classes].join(" "); },
    classList: {
      add: c => classes.add(c), remove: c => classes.delete(c), contains: c => classes.has(c),
      toggle: (c, on) => (on ? classes.add(c) : classes.delete(c)),
    },
    appendChild(e) { e.parentNode = this; this.children.push(e); },
    insertBefore(e, ref) {
      e.parentNode = this;
      const i = ref ? this.children.indexOf(ref) : -1;
      if (i < 0) this.children.push(e); else this.children.splice(i, 0, e);
    },
    remove() {
      if (!this.parentNode) return;
      const siblings = this.parentNode.children;
      siblings.splice(siblings.indexOf(this), 1);
      this.parentNode = null;
    },
    get nextSibling() {
      if (!this.parentNode) return null;
      const siblings = this.parentNode.children;
      return siblings[siblings.indexOf(this) + 1] || null;
    },
    addEventListener(type, fn) { this.listeners[type] = fn; }, setAttribute() {}, insertAdjacentHTML() {}, focus() {},
  };
}

// taskpane.js in a fresh page. `sent` is what it posted to the add-in; `host(message)` delivers one to it.
// `texts`: how i18n.load ends (by default never: the codes show, and "ready" is not sent); `readyState`: the page's
// document.readyState as taskpane.js starts ("loading" until the test fires DOMContentLoaded).
function page({ texts = () => new Promise(() => {}), readyState = "complete" } = {}) {
  const sent = [], fetched = [], elements = new Map(), onDocument = {}, onResize = [];
  let onMessage = null;
  const context = vm.createContext({
    console, URLSearchParams, location: { search: "" },
    // nothing is laid out here: a test says when what the page observes changed its size (resize)
    ResizeObserver: class { constructor(fn) { onResize.push(fn); } observe() {} },
    // nothing answers: the add-in's proxy refuses whatever the page asks the network for
    fetch: url => { fetched.push(url); return Promise.reject(new TypeError("Failed to fetch")); },
    window: { chrome: { webview: {
      postMessage: m => sent.push(m),
      addEventListener: (type, fn) => { if (type === "message") onMessage = fn; },
    } } },
    document: {
      body: element(), createElement: element, readyState,
      getElementById: id => { if (!elements.has(id)) elements.set(id, element()); return elements.get(id); },
      addEventListener: (type, fn) => { onDocument[type] = fn; },
    },
    i18n: { t: m => (typeof m === "string" ? m : m.code), fill: code => code, load: texts, applyText() {} },
  });
  vm.runInContext(source, context);
  const run = code => vm.runInContext(code, context);
  return {
    sends: () => sent.filter(m => m.type === "send").map(m => m.text),
    sent: type => sent.filter(m => m.type === type),
    posted: () => sent.map(m => m.type),
    fetched: () => fetched,
    fire: type => { if (onDocument[type]) onDocument[type](); },
    resize: () => { for (const fn of onResize) fn(); },
    submit: text => { context.text = text; run("input.value = text; submit();"); },
    host: data => onMessage({ data }),
    log: (code, ...args) => onMessage({ data: { type: "log", code, args } }),
    work: () => run("work"),
    element: id => run("document.getElementById(" + JSON.stringify(id) + ")"),
    bodyHas: c => run("document.body.classList.contains(" + JSON.stringify(c) + ")"),
    click: id => run("document.getElementById(" + JSON.stringify(id) + ")").listeners.click(),
    lastBlock: () => { const children = run("log.children"); return children.filter(e => e.className.startsWith("work")).pop(); },
  };
}

function oneAtATime() {
  const p = page();
  p.submit("first");
  const onItsWay = p.bodyHas("sending");
  p.submit("second");
  check("a second send before the busy state arrives is not sent; the page shows the first is on its way (body.sending, for docs.js)",
        p.sends().join(",") === "first" && onItsWay, p.sends().join(",") + " sending=" + onItsWay);
  p.host({ type: "busy", value: true });
  p.submit("third");
  check("nor while the request runs (then busy, not sending)", p.sends().join(",") === "first" && !p.bodyHas("sending"),
        p.sends().join(",") + " sending=" + p.bodyHas("sending"));
  p.host({ type: "busy", value: false });
  p.submit("fourth");
  check("the next one goes once it has ended", p.sends().join(",") === "first,fourth", p.sends().join(","));
}

function elsewhere() {
  const p = page();
  p.host({ type: "busy", value: true, elsewhere: true });
  p.submit("first");
  const action = p.element("action"), status = p.element("status");
  check("while another window's pane runs a request, nothing is sent and there is nothing to stop",
        p.sends().length === 0 && action.disabled === true && status.textContent === "STATUS_BUSY_ELSEWHERE",
        "sends=" + p.sends().length + " disabled=" + action.disabled + " status=" + status.textContent);
  p.host({ type: "busy", value: false });
  p.submit("second");
  check("once it has ended, this pane sends again", p.sends().join(",") === "second" && action.disabled === false && status.textContent === "",
        p.sends().join(",") + " disabled=" + action.disabled + " status=" + status.textContent);
}

function refused() {
  const p = page();
  p.submit("first");
  // the add-in refuses it (no endpoint): it logs why, and its busy state goes back to false without having been true
  p.log("USER_REQUEST", "first");
  p.log("NO_ENDPOINT_OR_MODEL", "HKCU");
  p.host({ type: "busy", value: false });
  p.submit("second");
  check("a refused send does not leave the page waiting", p.sends().join(",") === "first,second", p.sends().join(","));
}

// How a request's block ends after `steps`: "" when as `state` (its class and title), or what it showed instead.
function endedOtherwise(steps, state) {
  const p = page();
  p.submit("request");
  p.log("USER_REQUEST", "request");
  p.host({ type: "busy", value: true });
  steps(p);
  p.host({ type: "busy", value: false });
  const block = p.lastBlock();
  const cls = block ? block.className : "(no block)";
  const title = block ? block.children[0].children[1].textContent : "";
  const expected = state === "done" ? "WORK_DONE" : "WORK_STOPPED";
  return cls.split(" ").includes(state) && title.startsWith(expected) && p.work() === null ? "" : cls + " / " + title;
}

// A request ends as stopped whatever stopped it (each kind of log line has its own entry in logKinds), and as done when
// it finished.
function ends() {
  const stoppedBy = { AI_DECK_CLOSED: [], AI_CANCELLED: [], AI_ERROR: ["x"] };
  const wrong = Object.entries(stoppedBy).map(([code, args]) => [code, endedOtherwise(p => p.log(code, ...args), "stopped")]).filter(([, got]) => got !== "");
  check("a request whose deck was closed, stopped by 中止, or that failed ends as stopped", wrong.length === 0, wrong.map(w => w.join(": ")).join(", "));
  const done = endedOtherwise(p => { p.log("AI_REPLY", "ok"); p.log("REQUEST_DONE", "0:01", "1", 0); }, "done");
  check("a request that finished ends as done", done === "", done);
}

function selection() {
  const p = page();
  const line = p.element("selection"), text = p.element("selection-text");
  p.host({ type: "selection", text: { code: "SELECTION_SHAPES", args: [3, "x"] } });
  check("what is selected in PowerPoint shows above the input",
        line.classList.contains("show") && text.textContent === "SELECTION_SHAPES", line.className + " / " + text.textContent);
  p.host({ type: "selection", text: null });
  check("and goes when there is nothing to show", !line.classList.contains("show") && text.textContent === "",
        line.className + " / " + text.textContent);
}

function about() {
  const p = page();
  p.host({ type: "about", text: { code: "ABOUT_VERSION", args: ["0.0.4", "5efe257"] } });
  check("the version and commit show at the bottom of the settings", p.element("about").textContent === "ABOUT_VERSION",
        p.element("about").textContent);
}

// The line under the version: whether the pane's browser sends what it asks for to the add-in's refusing proxy. The
// add-in gives an address to ask for once; the page asks, says when that request has ended (it fails: the proxy
// refuses it), and shows what the add-in then found.
async function proxyState() {
  const tick = () => new Promise(resolve => setTimeout(resolve, 0));
  const p = page();
  const line = p.element("proxy");
  p.host({ type: "proxy", probe: "https://probe.example.invalid/" });
  check("the page asks for the address the add-in gives, and shows nothing yet",
        p.fetched().join(",") === "https://probe.example.invalid/" && line.textContent === "", p.fetched().join(",") + " / " + line.textContent);
  await tick();
  check("it tells the add-in once that request has ended, failed as it is", p.sent("probed").length === 1, p.posted().join(","));
  p.host({ type: "proxy", on: true });
  check("the settings show that the refusing proxy is in use", line.textContent === "PROXY_ON" && p.fetched().length === 1, line.textContent);
  p.host({ type: "proxy", on: false });
  check("or that it is not", line.textContent === "PROXY_OFF" && p.fetched().length === 1, line.textContent);
}

// The settings show the add-in's values (the switches, the context limit, the image prompt); everything in them is
// locked while a request runs, and free again once it has ended (削除 of the API key has rules of its own: apiKeyDelete).
function options() {
  const p = page();
  p.host({ type: "options", openPane: false, runAsArrives: true, saveProgress: true, externalControl: false, contextLimit: 196608, imagePrompt: "zimage" });
  const open = p.element("open-pane"), arrives = p.element("run-as-arrives"), save = p.element("save-progress"), ext = p.element("external-control");
  const limit = p.element("context-limit"), prompt = p.element("image-prompt");
  check("the settings show the add-in's values",
        open.checked === false && arrives.checked === true && save.checked === true && ext.checked === false && limit.value === "196608" && prompt.value === "zimage",
        `openPane=${open.checked} runAsArrives=${arrives.checked} saveProgress=${save.checked} externalControl=${ext.checked} contextLimit=${limit.value} imagePrompt=${prompt.value}`);
  const ids = ["endpoint", "fetch", "chat-model", "image-model", "apikey-input", "apikey", "new-chat", "open-pane", "run-as-arrives", "save-progress",
               "external-control", "context-limit", "image-prompt"];
  p.host({ type: "busy", value: true });
  const unlocked = ids.filter(id => p.element(id).disabled !== true);
  check("everything in the settings is locked while a request runs", unlocked.length === 0, "not locked: " + unlocked.join(","));
  p.host({ type: "busy", value: false });
  const locked = ids.filter(id => p.element(id).disabled !== false);
  check("and free again once it has ended", locked.length === 0, "still locked: " + locked.join(","));
}

// A changed context limit goes to the add-in as typed (it checks the number).
function contextLimit() {
  const p = page();
  const input = p.element("context-limit");
  input.value = " 50000 ";
  input.listeners.change({ target: input });
  const posted = p.sent("contextLimit").map(m => m.value).join(",");
  check("a changed limit goes to the add-in", posted === "50000", posted);
}

// The model drop-downs: with no chat model chosen, a first choice that asks for one and cannot be chosen back; the image
// model's first choice is （使わない）, the empty model, which goes to the add-in like any other; a saved model the endpoint
// does not list stays.
function modelChoices() {
  const p = page();
  const chat = p.element("chat-model"), image = p.element("image-model");
  const choices = select => select.children.map(o => (o.disabled ? "!" : "") + (o.value || "''") + "=" + o.textContent).join(" ");
  p.host({ type: "models", chat: ["c1", "c2"], chatSelected: "", image: ["i1"], imageSelected: "" });
  check("no chat model chosen: the first choice asks for one", choices(chat) === "!''=CHAT_MODEL_NONE c1=c1 c2=c2" && chat.value === "", choices(chat) + " / " + chat.value);
  check("no image model: （使わない）", choices(image) === "''=IMAGE_MODEL_NONE i1=i1" && image.value === "", choices(image) + " / " + image.value);
  p.host({ type: "models", chat: ["c1", "c2"], chatSelected: "old", image: ["i1"], imageSelected: "i1" });
  check("a chosen chat model: no first choice, and a saved one the endpoint does not list stays", choices(chat) === "old=old c1=c1 c2=c2" && chat.value === "old",
        choices(chat) + " / " + chat.value);
  check("a chosen image model: （使わない） can still be chosen", choices(image) === "''=IMAGE_MODEL_NONE i1=i1" && image.value === "i1", choices(image) + " / " + image.value);
  image.value = "";
  image.listeners.change();
  const posted = p.sent("imageModel").map(m => JSON.stringify(m.value)).join(",");
  check("（使わない） goes to the add-in as the empty model", posted === '""', posted);
}

// A changed image prompt goes to the add-in.
function imagePrompt() {
  const p = page();
  const select = p.element("image-prompt");
  select.value = "other";
  select.listeners.change({ target: select });
  const posted = p.sent("imagePrompt").map(m => m.value).join(",");
  check("a changed image prompt goes to the add-in", posted === "other", posted);
}

// The API key: typed into the settings and sent with 設定 or Enter, once, and gone from the field at once; the field
// says only whether a key is set; an empty field sends nothing.
function apiKey() {
  const p = page();
  const input = p.element("apikey-input"), button = p.element("apikey");
  const sent = () => JSON.stringify(p.sent("apiKey").map(m => m.value));
  p.host({ type: "apiKeyState", value: true });
  check("the field says a key is set", input.placeholder === "API_KEY_SET", input.placeholder);
  p.click("apikey");
  check("an empty field sends nothing", sent() === "[]", sent());
  input.value = " test-key-1 ";
  p.click("apikey");
  check("設定 sends the key once and empties the field", sent() === '["test-key-1"]' && input.value === "", sent() + " / " + input.value);
  input.value = "test-key-2";
  input.listeners.keydown({ key: "Enter", preventDefault() {} });
  check("Enter sends it too", sent() === '["test-key-1","test-key-2"]' && input.value === "", sent() + " / " + input.value);
}

// 削除 of the API key: only while a key is set and nothing runs; it asks under the key's row, やめる sends nothing, and
// 削除する sends it once; a request starting takes the question away.
function apiKeyDelete() {
  const p = page();
  const button = p.element("apikey-delete"), asking = () => p.element("apikey-confirm").classList.contains("show");
  const sent = () => p.sent("apiKeyDelete").length;
  p.host({ type: "apiKeyState", value: false });
  p.click("apikey-delete");
  check("with no key, 削除 cannot be pressed and asks nothing", button.disabled === true && !asking() && sent() === 0,
        `disabled=${button.disabled} asking=${asking()} sent=${sent()}`);
  p.host({ type: "apiKeyState", value: true });
  p.click("apikey-delete");
  check("with a key it asks first, and sends nothing yet", button.disabled === false && asking() && sent() === 0, `asking=${asking()} sent=${sent()}`);
  p.click("apikey-delete-no");
  check("やめる: the question goes, nothing is sent", !asking() && sent() === 0, `asking=${asking()} sent=${sent()}`);
  p.click("apikey-delete");
  p.click("apikey-delete-yes");
  check("削除する: it is sent once, and the question goes", !asking() && sent() === 1, `asking=${asking()} sent=${sent()}`);
  p.click("apikey-delete");
  p.host({ type: "busy", value: true });
  check("a request starting takes the question away and locks 削除", !asking() && button.disabled === true, `asking=${asking()} disabled=${button.disabled}`);
  p.host({ type: "busy", value: false });
  check("and 削除 is back once it has ended, while the key is set", button.disabled === false, `disabled=${button.disabled}`);
}

// The settings open when there is no endpoint, and never fold by themselves: the add-in sends the endpoint again after
// 取得 and after the key changes, while the user is still in them.
function settingsStayOpen() {
  const p = page();
  const settings = p.element("settings");
  p.host({ type: "endpoint", value: "" });
  check("with no endpoint the settings open", settings.open === true, `open=${settings.open}`);
  p.host({ type: "endpoint", value: "https://llm.example.com" });
  check("an endpoint sent again does not fold them", settings.open === true, `open=${settings.open}`);
}

// 新しい会話: not while a request is on its way (sent, its busy state not back yet) nor while it runs - it would end the
// conversation that request runs in; once it has ended, it goes.
function newChat() {
  const p = page();
  p.submit("request");
  p.click("new-chat");
  check("新しい会話 right after a send is not sent", p.sent("newChat").length === 0, p.posted().join(","));
  p.host({ type: "busy", value: true });
  p.click("new-chat");
  check("nor while the request runs", p.sent("newChat").length === 0, p.posted().join(","));
  p.host({ type: "busy", value: false });
  p.click("new-chat");
  check("once it has ended, it goes", p.sent("newChat").length === 1, p.posted().join(","));
}

// 続ける: last in the conversation, after all that the stopped request showed (what comes later goes above it);
// sent like a request - once, until the add-in answers with its busy state - and gone once pressed or when the add-in
// says that request can no longer be carried on. While a request runs elsewhere it stays, disabled.
function continueButton() {
  const p = page();
  const log = p.element("log");
  p.submit("request");
  p.log("USER_REQUEST", "request");
  p.host({ type: "busy", value: true });
  p.log("AI_TURN_LIMIT", 20);
  p.log("REQUEST_DONE", "0:30", "1", 10);
  p.host({ type: "busy", value: false });
  p.log("RUN_SAVED", "C:\\run");
  p.host({ type: "continue", value: true });
  const classes = () => log.children.map(e => e.className).join(" | ");
  const row = log.children[log.children.length - 1];
  check("続ける shows last, after all that the stopped request showed",
        row.className === "continue-row" && row.children[0].textContent === "CONTINUE" && row.children[0].disabled === false, classes());
  p.log("MODELS_LISTED", "x", 1, 1);
  check("what comes while it shows goes above it", log.children[log.children.length - 1] === row && log.children.length >= 2 &&
        log.children[log.children.length - 2].className === "note", classes());
  row.children[0].listeners.click();
  row.children[0].listeners.click();
  check("pressed, it is sent once and goes", p.posted().filter(t => t === "continue").length === 1 && !log.children.includes(row),
        p.posted().join(",") + " / " + classes());
  p.log("CONTINUE_REQUEST");
  const last = log.children.filter(e => !e.className.startsWith("work")).pop();
  check("its 続けて comes after all of the stopped request, as the user's message, and starts the request's block",
        !!last && last.className === "msg user" && p.work() !== null, classes());
  p.host({ type: "continue", value: false });
  p.host({ type: "busy", value: true });
  p.log("AI_TURN_LIMIT", 20);
  p.host({ type: "busy", value: false });
  p.host({ type: "continue", value: true });
  p.host({ type: "busy", value: true, elsewhere: true });
  const again = log.children.find(e => e.className === "continue-row");
  check("while another window's pane runs a request it stays, disabled", !!again && again.children[0].disabled === true,
        again ? `disabled=${again.children[0].disabled}` : "(no 続ける)");
  p.host({ type: "busy", value: false });
  check("and can be pressed once that has ended", again.children[0].disabled === false, `disabled=${again.children[0].disabled}`);
  p.host({ type: "continue", value: false });
  check("another request (or 新しい会話) takes it away", !log.children.some(e => e.className === "continue-row"),
        log.children.map(e => e.className).join(" | "));
}

// While a request runs, the add-in keeps the slides of its window from being edited by hand; the page says so
// (body.locked shows the line above the input) for as long as the add-in says it does, and not in a pane that only
// waits for another window's request.
function slidesLocked() {
  const p = page();
  p.host({ type: "busy", value: true, locked: true });
  const during = p.bodyHas("locked");
  p.host({ type: "busy", value: false });
  check("the page says the slides are held while its request runs, and no longer once it has ended",
        during && !p.bodyHas("locked"), "during=" + during + " after=" + p.bodyHas("locked"));
  p.host({ type: "busy", value: true, elsewhere: true });
  const waiting = p.bodyHas("locked");
  p.host({ type: "busy", value: true, elsewhere: true, locked: true });
  check("a pane that waits for another window's request says so only when its own window's slides are held too",
        !waiting && p.bodyHas("locked"), "waiting=" + waiting + " same deck=" + p.bodyHas("locked"));
}

// A script whose reply was stopped after part of it ran: that the part stays on the slides is said below the request's
// block, not as one of its rows (the block may be folded), before how the request ended.
function partKept() {
  const p = page();
  const log = p.element("log");
  p.submit("request");
  p.log("USER_REQUEST", "request");
  p.host({ type: "busy", value: true });
  p.log("SCRIPT_PART_KEPT");
  p.log("AI_CANCELLED");
  p.host({ type: "busy", value: false });
  const shown = log.children.map(e => (e.className.startsWith("work") ? "work" : e.className.split(" ")[0] + ":" + e.textContent)).join(", ");
  check("what stays of a stopped script is told below the request's block, before how the request ended",
        shown === "msg:USER_REQUEST, work, note:SCRIPT_PART_KEPT, msg:AI_CANCELLED", shown);
}

// A file that was not attached shows in red (alert), with the reason: as a note between requests, as a row of the
// request's block while one runs. One that was attached is a side note like any other.
function notAttached() {
  const p = page();
  const log = p.element("log");
  p.log("DOC_ATTACHED", "d1", "a.pdf", "x");
  p.log("DOC_NOT_ATTACHED", "b.pdf", { code: "DOC_FAILED_TOO_LONG", args: ["1", "2"] });
  p.log("DOC_DISCARDED", "c.pdf");
  const classes = log.children.map(e => e.className).join(" | ");
  check("a file that was not attached shows in red, with the reason",
        classes === "note | note alert | note alert" && log.children[1].textContent === "DOC_NOT_ATTACHED" && log.children[2].textContent === "DOC_DISCARDED",
        classes + " / " + log.children.map(e => e.textContent).join(" | "));

  p.submit("request");
  p.log("USER_REQUEST", "request");
  p.host({ type: "busy", value: true });
  p.log("DOC_NOT_ATTACHED", "d.pdf", { code: "DOC_FAILED_BUSY", args: [] });
  const rows = p.lastBlock().children[1].children;
  const row = rows[rows.length - 1];
  check("and while a request runs, as a row of its block, in red too",
        !!row && row.className === "row extra alert" && row.children[0].textContent === "DOC_NOT_ATTACHED", row ? row.className : "(no row)");
}

// The conversation follows its end until the user scrolls up to look back, and again once they scroll down to the
// end, press the arrow, or a request is sent. Nothing is laid out here: the test says how tall the conversation and
// the room for it are, and fires the events a browser would; the scroll position is kept within what there is to
// scroll, as a browser keeps it.
function followsItsEnd() {
  const p = page();
  const log = p.element("log"), arrow = p.element("to-end");
  let top = 0;
  Object.defineProperty(log, "scrollTop", { get: () => top, set: v => { top = Math.max(0, Math.min(v, log.scrollHeight - log.clientHeight)); } });
  const tall = (all, room) => { log.scrollHeight = all; if (room != null) log.clientHeight = room; };
  const below = () => log.scrollHeight - log.scrollTop - log.clientHeight;   // how far the end is under what is seen
  const shown = () => arrow.classList.contains("show");
  const scrollTo = to => { log.scrollTop = to; log.listeners.scroll(); };
  const state = () => "end " + below() + " px below, arrow " + (shown() ? "shown" : "hidden");

  tall(1000, 700);
  p.log("USER_REQUEST", "one");
  check("the conversation follows its end as things are added", below() === 0 && !shown(), state());
  // A request begins: the status line and the note that the slides are locked take room from the conversation. Its
  // end is then out of sight though nobody scrolled, and the scroll event of the page's own move comes after that.
  tall(1000, 632);
  log.listeners.scroll();
  tall(1060);
  p.log("AI_STEP", "a notice");
  check("it still follows when the room for it got smaller: the end left the sight, but nobody scrolled up", below() === 0 && !shown(), state());
  tall(1060, 600);
  const left = below();
  p.resize();
  check("and comes back to its end when the room changes", left === 32 && below() === 0, "left " + left + " px below; then " + state());
  tall(1100);
  p.log("TOOL_RESULT", 1, "pptcom", "OK");
  check("a row of the request's block is followed too", below() === 0, state());

  scrollTo(200);
  const up = shown();
  tall(1200);
  p.log("AI_STEP", "another");
  p.resize();
  check("once the user scrolls up it stays where they put it, and the arrow shows", up && log.scrollTop === 200 && shown(), "arrow shown=" + up + "; top " + log.scrollTop + ", " + state());
  scrollTo(400);
  tall(1300);
  p.log("AI_STEP", "more");
  check("scrolling down short of the end does not follow yet", log.scrollTop === 400 && shown(), "top " + log.scrollTop + ", " + state());
  scrollTo(1e9);
  const back = !shown();
  tall(1400);
  p.log("AI_STEP", "yet more");
  check("scrolled down to the end, it follows again and the arrow goes", back && below() === 0 && !shown(), "arrow gone=" + back + "; " + state());

  scrollTo(100);
  p.click("to-end");
  check("the arrow takes it to the end and follows again", below() === 0 && !shown(), state());
  scrollTo(100);
  tall(1500);
  p.log("USER_REQUEST", "two");
  check("a request sent is shown wherever the user had scrolled to", below() === 0 && !shown(), state());

  // a block folded or unfolded by the user: the request's own is followed on, another one is being looked at
  tall(1800);
  log.listeners.toggle({ target: p.work().el });
  check("the block of the request in progress, unfolded, is followed to its end", below() === 0 && !shown(), state());
  tall(2200);
  log.listeners.toggle({ target: {} });
  check("another block unfolded stays in sight: the conversation stops following", below() === 400 && shown(), state());
  tall(40, 600);
  p.log("NEW_CONVERSATION");
  check("a new conversation follows from its start", !shown(), state());
}

// "ready" waits for docs.js as well as the texts: the add-in then sends at once what waited, files to convert among it,
// and only docs.js takes those. docs.js is a module script, so it has run by DOMContentLoaded.
async function readyWaitsForDocs() {
  const tick = () => new Promise(resolve => setTimeout(resolve, 0));
  const p = page({ texts: () => Promise.resolve(), readyState: "loading" });
  await tick();
  check("ready waits for the page's module scripts (docs.js) even with the texts in", !p.posted().includes("ready"), p.posted().join(","));
  p.fire("DOMContentLoaded");
  await tick();
  check("and is sent once the page is parsed", p.posted().filter(t => t === "ready").length === 1, p.posted().join(","));
  const q = page({ texts: () => Promise.resolve() });
  await tick();
  check("a page already parsed sends ready with the texts", q.posted().filter(t => t === "ready").length === 1, q.posted().join(","));
}

oneAtATime();
elsewhere();
refused();
selection();
about();
options();
contextLimit();
modelChoices();
imagePrompt();
apiKey();
apiKeyDelete();
settingsStayOpen();
newChat();
continueButton();
ends();
notAttached();
partKept();
slidesLocked();
followsItsEnd();
readyWaitsForDocs().then(proxyState).then(() => {
  console.log();
  console.log(failed === 0 ? `ALL PASSED (${passed})` : `${failed} FAILED, ${passed} passed`);
  process.exitCode = failed === 0 ? 0 : 1;
}).catch(e => { console.error(e); process.exitCode = 1; });
