// How the task pane's page (pptcom/UI/web/taskpane.js) sends a request and shows how it ends: one request at a
// time, also before the add-in's busy state arrives and while another window's pane runs one; a refused send does
// not leave the page waiting; a request that stops (the deck was closed, 中止, an error) closes its block as stopped,
// one that finishes as done; 続ける shows after a request that stopped at the turn limit, and sends like a request;
// the line above the input shows what is selected in PowerPoint, and the settings the version. taskpane.js runs
// as it is, in a stand-in for the browser and WebView2 made of plain objects; messages from the add-in arrive as
// WebView2 delivers them. No packages.
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
  return {
    children: [], value: "", textContent: "", scrollHeight: 0, scrollTop: 0, clientHeight: 0, disabled: false, listeners: {},
    set className(s) { classes.clear(); for (const c of s.split(/\s+/)) if (c) classes.add(c); },
    get className() { return [...classes].join(" "); },
    classList: {
      add: c => classes.add(c), remove: c => classes.delete(c), contains: c => classes.has(c),
      toggle: (c, on) => (on ? classes.add(c) : classes.delete(c)),
    },
    appendChild(e) { this.children.push(e); },
    insertBefore(e) { this.children.unshift(e); },
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

// 続ける: after a request that stopped at the turn limit, hidden while any request runs (here or in another window), and
// sent like a request - once, until the add-in answers with its busy state.
function continueBar() {
  const p = page();
  const bar = p.element("continue-bar");
  p.host({ type: "continue", value: true });
  check("続ける shows after a request that stopped at the turn limit", bar.classList.contains("show"), bar.className);
  p.click("continue");
  p.click("continue");
  check("and is sent once until the busy state arrives", p.posted().filter(t => t === "continue").length === 1, p.posted().join(","));
  p.log("CONTINUE_REQUEST");
  p.host({ type: "busy", value: true });
  check("it shows as the user's message and starts the request's block",
        p.element("log").children.some(e => e.className === "msg user") && p.work() !== null && !bar.classList.contains("show"),
        bar.className);
  p.host({ type: "continue", value: false });
  p.host({ type: "busy", value: false });
  check("a request carried on to the end leaves no 続ける", !bar.classList.contains("show"), bar.className);
  p.host({ type: "continue", value: true });
  p.host({ type: "busy", value: true, elsewhere: true });
  check("hidden while another window's pane runs a request", !bar.classList.contains("show"), bar.className);
  p.host({ type: "busy", value: false });
  check("and back once it has ended", bar.classList.contains("show"), bar.className);
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
continueBar();
ends("a request whose deck was closed ends as stopped", p => p.log("AI_DECK_CLOSED"), "stopped");
ends("a request stopped by 中止 ends as stopped", p => p.log("AI_CANCELLED"), "stopped");
ends("a request that failed ends as stopped", p => p.log("AI_ERROR", "x"), "stopped");
ends("a request that finished ends as done", p => { p.log("AI_REPLY", "ok"); p.log("REQUEST_DONE", "0:01", "1", 0); }, "done");
readyWaitsForDocs().then(() => {
  console.log();
  console.log(failed === 0 ? `ALL PASSED (${passed})` : `${failed} FAILED, ${passed} passed`);
  process.exitCode = failed === 0 ? 0 : 1;
}).catch(e => { console.error(e); process.exitCode = 1; });
