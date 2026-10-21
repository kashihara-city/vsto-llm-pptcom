"use strict";
// The add-in (AiTaskPaneControl.cs) and this page talk only through WebView2 web messages:
//   page -> add-in: {type: "ready" | "send" | "cancel" | "retry" | "newChat" | "endpoint" | "chatModel" | "imageModel" | "apiKey", ...}
//   page -> add-in also: {type: "openPane" | "saveProgress" | "externalControl", value: bool}
//   add-in -> page: {type: "log" | "status" | "busy" | "retry" | "endpoint" | "models" | "apiKeyState" | "options" | "selection" | "about", ...}
//   ("log" and "status" carry a message: {code, args}, shown with the text of locales/<language>.json, see i18n.js)
// Attached documents (drop, attach button, chips, PDF pages) have their own messages: see docs.js.
const host = window.chrome && window.chrome.webview;
const post = msg => { if (host) host.postMessage(msg); };
const $ = id => document.getElementById(id);

const log = $("log"), input = $("input"), action = $("action");
let busy = false, cancelling = false;
let sending = false;      // a request was sent and the add-in has not answered with its busy state yet
let work = null;          // the folded block of the request in progress

// ---- conversation ------------------------------------------------------------------------------

function nearBottom() { return log.scrollHeight - log.scrollTop - log.clientHeight < 40; }
function add(el) {
  const stick = nearBottom();
  log.appendChild(el);
  if (stick) log.scrollTop = log.scrollHeight;
}
function div(cls, text) {
  const el = document.createElement("div");
  el.className = cls;
  if (text != null) el.textContent = text;
  return el;
}

function startWork() {
  const d = document.createElement("details");
  d.className = "work running";
  const s = document.createElement("summary");
  s.appendChild(div("dot"));
  const title = div("work-title", i18n.t("WORK_THINKING"));
  s.appendChild(title);
  s.insertAdjacentHTML("beforeend", '<svg class="chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m6 9 6 6 6-6"/></svg>');
  d.appendChild(s);
  const body = div("work-body");
  d.appendChild(body);
  add(d);
  work = { el: d, title, body, turns: 0, tools: 0 };
}
function workRow(cls, tool, text) {
  if (!work) startWork();
  const stick = nearBottom();
  const r = div("row " + cls);
  if (tool) r.appendChild(div("tool", tool));
  r.appendChild(div("text", text));
  work.body.appendChild(r);
  if (stick) log.scrollTop = log.scrollHeight;
}
function finishWork(state, footer) {
  if (!work) return;
  work.el.classList.remove("running");
  work.el.classList.add(state);
  const parts = [];
  if (work.turns) parts.push(i18n.fill("WORK_TURNS", work.turns));
  if (work.tools) parts.push(i18n.fill("WORK_TOOLS", work.tools));
  if (footer) parts.push(footer);
  const sep = i18n.t("PART_SEPARATOR");
  work.title.textContent = i18n.t(state === "done" ? "WORK_DONE" : "WORK_STOPPED") + (parts.length ? sep + parts.join(sep) : "");
  work = null;
}

// How a log message from the add-in is shown, by its code (ThisAddIn.cs). Any other code is a side note:
// a row of the request in progress, or a note in parentheses between requests.
const logKinds = {
  USER_REQUEST: "user", TOOL_RESULT: "tool", TURN_SIZE: "turn", REQUEST_DONE: "done",
  AI_REPLY: "reply", AI_DONE: "reply",
  AI_ERROR: "stopped", AI_CANCELLED: "stopped", AI_SERVER_ERROR: "error", AI_TURN_LIMIT: "error", AI_REPLY_CUT: "error",
  AI_DECK_CLOSED: "stopped",
  // a request refused before it started: its block ends with the reason
  NO_ENDPOINT_OR_MODEL: "stopped", DOC_STILL_CONVERTING: "stopped",
  HAND_EDITS: "handEdits", NEW_CONVERSATION: "reset",
  READY: "plain", NO_ENDPOINT_HELP: "plain",
};

function onLog(m) {
  const args = m.args || [], text = i18n.t(m);
  switch (logKinds[m.code]) {
    case "user":
      if (work) finishWork("stopped");
      add(div("msg user", text));
      startWork();
      return;
    case "tool": {
      if (!work) startWork();
      work.tools++;
      const turn = args[0], tool = args[1], result = String(args[2] || "");
      const failed = /\bfailed\b|FAILED|rejected|error/i.test(result) && !/^OK/.test(result);
      workRow(failed ? "failed" : "", tool, result);
      work.title.textContent = [i18n.fill("TURN", turn), tool, result.split("\n")[0]].join(i18n.t("PART_SEPARATOR"));
      return;
    }
    case "turn":
      if (!work) startWork();
      work.turns = Math.max(work.turns, +args[0]);
      workRow("turn", null, text);
      return;
    case "done":
      if (work) finishWork("done", text);
      else add(div("note", i18n.fill("NOTE", text)));
      return;
    case "reply":
      add(div("msg assistant", text));
      return;
    case "stopped":   // the request ended with it
      if (work) finishWork("stopped");
      add(div("msg assistant error", text));
      return;
    case "error":
      add(div("msg assistant error", text));
      return;
    case "handEdits": {
      // what the user changed by hand since the last request: above this request's progress block
      const note = div("note", i18n.fill("NOTE", text));
      if (work) log.insertBefore(note, work.el); else add(note);
      return;
    }
    case "reset":     // 新しい会話 pressed: start from an empty pane
      work = null;
      log.textContent = "";
      add(div("divider", i18n.t("NEW_CHAT")));
      return;
    case "plain":
      add(div("note", text));
      return;
    default:
      if (work) workRow("extra", null, text);
      else add(div("note", i18n.fill("NOTE", text)));
  }
}

// ---- input, buttons -----------------------------------------------------------------------------

// A document still being read (body.converting, docs.js) would not go with the request: wait for it.
const converting = () => document.body.classList.contains("converting");

// One request at a time: a second send before the add-in's busy state arrives would start a second request.
// The add-in answers every send with its busy state, a refused one too.
function submit() {
  if (busy || sending || converting()) return;
  const text = input.value.trim();
  if (!text) return;
  input.value = "";
  sending = true;
  post({ type: "send", text });
}
input.addEventListener("keydown", e => {
  if (e.key === "Enter" && e.ctrlKey) { e.preventDefault(); submit(); }
});
action.addEventListener("click", () => {
  if (!busy) { submit(); return; }
  if (cancelling) return;
  cancelling = true;
  action.disabled = true;
  post({ type: "cancel" });
});
$("retry").addEventListener("click", () => post({ type: "retry" }));
$("new-chat").addEventListener("click", () => { if (!busy && !converting()) post({ type: "newChat" }); });

const endpoint = $("endpoint"), chatModel = $("chat-model"), imageModel = $("image-model");
const submitEndpoint = () => post({ type: "endpoint", value: endpoint.value.trim() });
$("fetch").addEventListener("click", submitEndpoint);
endpoint.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); submitEndpoint(); } });
chatModel.addEventListener("change", () => { post({ type: "chatModel", value: chatModel.value }); updateSummary(); });
imageModel.addEventListener("change", () => { post({ type: "imageModel", value: imageModel.value }); updateSummary(); });
$("apikey").addEventListener("click", () => post({ type: "apiKey" }));
$("open-pane").addEventListener("change", e => post({ type: "openPane", value: e.target.checked }));
$("save-progress").addEventListener("change", e => post({ type: "saveProgress", value: e.target.checked }));
$("external-control").addEventListener("change", e => post({ type: "externalControl", value: e.target.checked }));

function updateSummary() {
  const host = endpoint.value.replace(/^https?:\/\//, "");
  $("settings-summary").textContent = host ? (chatModel.value || i18n.t("SETTINGS_NO_MODEL")) + i18n.t("PART_SEPARATOR") + host : i18n.t("SETTINGS_NOT_SET");
}
function fill(select, items, selected) {
  select.textContent = "";
  const list = items.slice();
  if (selected && !list.includes(selected)) list.unshift(selected);   // keep a saved model the endpoint no longer lists
  for (const name of list) {
    const o = document.createElement("option");
    o.value = o.textContent = name;
    select.appendChild(o);
  }
  if (selected) select.value = selected;
}

// ---- messages from the add-in -------------------------------------------------------------------

// `elsewhere`: the request runs in another window's pane (each window has its own). This one waits meanwhile: nothing
// to send or stop here.
function setBusy(on, elsewhere) {
  busy = on;
  sending = false;
  cancelling = false;
  const waiting = on && !!elsewhere;
  document.body.classList.toggle("busy", on);
  action.disabled = waiting;
  action.title = i18n.t(waiting ? "STATUS_BUSY_ELSEWHERE" : on ? "STOP" : "SEND_TITLE");
  action.setAttribute("aria-label", i18n.t(waiting ? "STATUS_BUSY_ELSEWHERE" : on ? "STOP" : "SEND"));
  for (const el of [endpoint, $("fetch"), chatModel, imageModel, $("apikey"), $("new-chat"), $("open-pane"), $("save-progress"), $("external-control")]) el.disabled = on;
  if (waiting) $("status").textContent = i18n.t("STATUS_BUSY_ELSEWHERE");
  if (!on) { $("status").textContent = ""; $("retry-bar").classList.remove("show"); if (work) finishWork("done"); }
}

const handlers = {
  log: m => onLog(m),
  status: m => { $("status").textContent = m.code ? i18n.t(m) : ""; },
  busy: m => setBusy(!!m.value, !!m.elsewhere),
  retry: m => $("retry-bar").classList.toggle("show", !!m.value),
  // The settings stay folded once an endpoint is set; with none, they open so it can be entered.
  endpoint: m => { endpoint.value = m.value || ""; $("settings").open = !endpoint.value; updateSummary(); },
  models: m => { fill(chatModel, m.chat || [], m.chatSelected); fill(imageModel, m.image || [], m.imageSelected); updateSummary(); },
  apiKeyState: m => { $("apikey-state").textContent = i18n.t(m.value ? "API_KEY_SET" : "API_KEY_NOT_SET"); },
  options: m => {
    $("open-pane").checked = !!m.openPane;
    $("save-progress").checked = !!m.saveProgress;
    $("external-control").checked = !!m.externalControl;
  },
  // the add-in's version and the commit it was built from, at the bottom of the settings
  about: m => { $("about").textContent = m.text ? i18n.t(m.text) : ""; },
  // what is selected in PowerPoint (or which slide it shows); a request sends the same to the model
  selection: m => { $("selection-text").textContent = m.text ? i18n.t(m.text) : ""; $("selection").classList.toggle("show", !!m.text); },
};
if (host) host.addEventListener("message", e => { const h = handlers[e.data && e.data.type]; if (h) h(e.data); });

// "ready" once the texts are in and docs.js listens too. The add-in sends nothing before it and then at once all that
// waited, some of it for docs.js (files to convert, the attachment chips, PDF pages to draw): a message that came
// before docs.js listened would be lost, and a file to convert with it for good. docs.js is a module script, which
// runs after the page is parsed and before DOMContentLoaded. Without the texts the codes show.
const parsed = document.readyState === "loading"
  ? new Promise(resolve => document.addEventListener("DOMContentLoaded", resolve, { once: true }))
  : Promise.resolve();
Promise.all([i18n.load(new URLSearchParams(location.search).get("lang")).catch(() => {}), parsed]).then(() => {
  i18n.applyText(document);
  updateSummary();
  input.focus();
  post({ type: "ready" });
});
