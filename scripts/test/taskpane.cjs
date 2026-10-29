// How the task pane's page (pptcom/UI/web/taskpane.js) sends a request and shows how it ends: one request at a
// time, also before the add-in's busy state arrives and while another window's pane runs one; a refused send does
// not leave the page waiting; a request that stops (the deck was closed, 中止, an error) closes its block as stopped,
// one that finishes as done; 続ける shows last after a request that stopped at the turn limit, and sends like
// a request; the line above the input shows what is selected in PowerPoint, and the settings the version, the
// context limit, the models (none chosen yet, 使わない), the image prompt and the API key. taskpane.js runs as it is, in a stand-in for the browser and WebView2 made of plain objects; messages
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
  const sent = [], elements = new Map(), onDocument = {};
  let onMessage = null;
  const context = vm.createContext({
    console, URLSearchParams, location: { search: "" },
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
    fire: type => { if (onDocument[type]) onDocument[type](); },
    submit: text => { context.text = text; run("input.value = text; submit();"); },
    host: data => onMessage({ data }),
    log: (code, ...args) => onMessage({ data: { type: "log", code, args } }),
    work: () => run("work"),
    element: id => run("document.getElementById(" + JSON.stringify(id) + ")"),
    click: id => run("document.getElementById(" + JSON.stringify(id) + ")").listeners.click(),
    lastBlock: () => { const children = run("log.children"); return children.filter(e => e.className.startsWith("work")).pop(); },
  };
}

function oneAtATime() {
  const p = page();
  p.submit("first");
  p.submit("second");
  check("a second send before the busy state arrives is not sent", p.sends().join(",") === "first", p.sends().join(","));
  p.host({ type: "busy", value: true });
  p.submit("third");
  check("nor while the request runs", p.sends().join(",") === "first", p.sends().join(","));
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
  // the add-in refuses it (no endpoint): the request's block ends with the reason, and the busy state answers the send
  p.log("USER_REQUEST", "first");
  p.log("NO_ENDPOINT_OR_MODEL", "HKCU");
  p.host({ type: "busy", value: false });
  p.submit("second");
  check("a refused send does not leave the page waiting", p.sends().join(",") === "first,second", p.sends().join(","));
}

function ends(name, steps, state) {
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
  check(name, cls.split(" ").includes(state) && title.startsWith(expected) && p.work() === null, cls + " / " + title);
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

function options() {
  const p = page();
  p.host({ type: "options", openPane: false, saveProgress: true, externalControl: false });
  const open = p.element("open-pane"), save = p.element("save-progress"), ext = p.element("external-control");
  check("the settings' switches show the add-in's values", open.checked === false && save.checked === true && ext.checked === false,
        `openPane=${open.checked} saveProgress=${save.checked} externalControl=${ext.checked}`);
  p.host({ type: "busy", value: true });
  check("and are locked while a request runs", open.disabled === true && save.disabled === true, `open=${open.disabled} save=${save.disabled}`);
}

// The context limit shows the add-in's value, a changed one goes to the add-in as typed (it checks the number), and it is
// locked while a request runs like the other settings.
function contextLimit() {
  const p = page();
  const input = p.element("context-limit");
  p.host({ type: "options", openPane: true, saveProgress: false, externalControl: false, contextLimit: 196608 });
  check("the context limit shows the add-in's value", input.value === "196608", input.value);
  input.value = " 50000 ";
  input.listeners.change({ target: input });
  const posted = p.sent("contextLimit").map(m => m.value).join(",");
  check("a changed limit goes to the add-in", posted === "50000", posted);
  p.host({ type: "busy", value: true });
  check("and is locked while a request runs", input.disabled === true, `disabled=${input.disabled}`);
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

// The image prompt shows the add-in's value, a changed one goes to the add-in, and it is locked while a request runs.
function imagePrompt() {
  const p = page();
  const select = p.element("image-prompt");
  p.host({ type: "options", openPane: true, saveProgress: false, externalControl: false, contextLimit: 196608, imagePrompt: "zimage" });
  check("the image prompt shows the add-in's value", select.value === "zimage", select.value);
  select.value = "other";
  select.listeners.change({ target: select });
  const posted = p.sent("imagePrompt").map(m => m.value).join(",");
  check("a changed image prompt goes to the add-in", posted === "other", posted);
  p.host({ type: "busy", value: true });
  check("and is locked while a request runs", select.disabled === true, `disabled=${select.disabled}`);
}

// The API key: typed into the settings and sent with 設定 or Enter, once, and gone from the field at once; the field
// says only whether a key is set; an empty field sends nothing; locked while a request runs.
function apiKey() {
  const p = page();
  const input = p.element("apikey-input"), button = p.element("apikey");
  const sent = () => JSON.stringify(p.sent("apiKey").map(m => m.value));
  p.host({ type: "apiKeyState", value: true });
  check("the field says a key is set, and holds nothing", input.placeholder === "API_KEY_SET" && input.value === "", `${input.placeholder} / ${input.value}`);
  p.click("apikey");
  check("an empty field sends nothing", sent() === "[]", sent());
  input.value = " test-key-1 ";
  p.click("apikey");
  check("設定 sends the key once and empties the field", sent() === '["test-key-1"]' && input.value === "", sent() + " / " + input.value);
  input.value = "test-key-2";
  input.listeners.keydown({ key: "Enter", preventDefault() {} });
  check("Enter sends it too", sent() === '["test-key-1","test-key-2"]' && input.value === "", sent() + " / " + input.value);
  p.host({ type: "busy", value: true });
  check("and both are locked while a request runs", input.disabled === true && button.disabled === true, `${input.disabled} ${button.disabled}`);
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
continueButton();
ends("a request whose deck was closed ends as stopped", p => p.log("AI_DECK_CLOSED"), "stopped");
ends("a request stopped by 中止 ends as stopped", p => p.log("AI_CANCELLED"), "stopped");
ends("a request that failed ends as stopped", p => p.log("AI_ERROR", "x"), "stopped");
ends("a request that finished ends as done", p => { p.log("AI_REPLY", "ok"); p.log("REQUEST_DONE", "0:01", "1", 0); }, "done");
readyWaitsForDocs().then(() => {
  console.log();
  console.log(failed === 0 ? `ALL PASSED (${passed})` : `${failed} FAILED, ${passed} passed`);
  process.exitCode = failed === 0 ? 0 : 1;
}).catch(e => { console.error(e); process.exitCode = 1; });
