"use strict";
// The add-in (AiTaskPaneControl.cs) and this page talk only through WebView2 web messages:
//   page -> add-in: {type: "ready" | "send" | "cancel" | "retry" | "newChat" | "endpoint" | "chatModel" | "imageModel" | "apiKey", ...}
//   page -> add-in also: {type: "saveProgress" | "externalControl", value: bool}
//   add-in -> page: {type: "log" | "status" | "busy" | "retry" | "endpoint" | "models" | "apiKeyState" | "options", ...}
// Attached documents (drop, attach button, chips, PDF pages) have their own messages: see docs.js.
const host = window.chrome && window.chrome.webview;
const post = msg => { if (host) host.postMessage(msg); };
const $ = id => document.getElementById(id);

const log = $("log"), input = $("input"), action = $("action");
let busy = false, cancelling = false;
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
  const title = div("work-title", "考えています…");
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
  if (work.turns) parts.push(work.turns + " ターン");
  if (work.tools) parts.push("ツール " + work.tools + " 回");
  if (footer) parts.push(footer);
  work.title.textContent = (state === "done" ? "処理の経過" : "処理の経過（途中で終了）") + (parts.length ? " · " + parts.join(" · ") : "");
  work = null;
}

// One line from the add-in's log, placed by what it starts with (see ThisAddIn.cs).
function onLog(text) {
  let m;
  if (text.startsWith("あなた: ")) {
    if (work) finishWork("stopped");
    add(div("msg user", text.slice(5)));
    startWork();
    return;
  }
  if ((m = /^AI: \[(\d+)\] (\S+) -> ([\s\S]*)$/.exec(text))) {
    if (!work) startWork();
    work.tools++;
    const result = m[3];
    const failed = /\bfailed\b|FAILED|rejected|error/i.test(result) && !/^OK/.test(result);
    workRow(failed ? "failed" : "", m[2], result);
    work.title.textContent = "ターン" + m[1] + " · " + m[2] + " · " + result.split("\n")[0];
    return;
  }
  if ((m = /^\(ターン(\d+): ([\s\S]*)\)$/.exec(text))) {
    if (!work) startWork();
    work.turns = Math.max(work.turns, +m[1]);
    workRow("turn", null, "ターン" + m[1] + ": " + m[2]);
    return;
  }
  if ((m = /^\((所要 [\s\S]*)\)$/.exec(text))) {
    if (work) finishWork("done", m[1]);
    else add(div("note", text));
    return;
  }
  if (text.startsWith("AI: ")) {
    const body = text.slice(4);
    const bad = /^(エラーが発生しました|サーバ側のエラー|中止しました|ターン数の上限)/.test(body);
    if (bad && /^(エラー|中止)/.test(body) && work) finishWork("stopped");
    add(div("msg assistant" + (bad ? " error" : ""), body));
    return;
  }
  // What the user changed by hand since the last request: shown above this request's progress block.
  if (/^\(前回の依頼のあとに/.test(text)) {
    const note = div("note", text);
    if (work) log.insertBefore(note, work.el); else add(note);
    return;
  }
  // 新しい会話 pressed: start from an empty pane. A conversation that ended because another
  // presentation is active (the request just sent is already shown) gets a divider instead.
  if (text === "(新しい会話を始めました)") {
    work = null;
    log.textContent = "";
    add(div("divider", "新しい会話"));
    return;
  }
  if ((m = /^\((新しい会話[\s\S]*)\)$/.exec(text))) {
    if (work) finishWork("stopped");
    add(div("divider", m[1]));
    return;
  }
  if (work && /^\(/.test(text)) { workRow("extra", null, text.replace(/^\(|\)$/g, "")); return; }
  add(div("note", text));
}

// ---- input, buttons -----------------------------------------------------------------------------

function submit() {
  if (busy) return;
  const text = input.value.trim();
  if (!text) return;
  input.value = "";
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
$("new-chat").addEventListener("click", () => { if (!busy) post({ type: "newChat" }); });

const endpoint = $("endpoint"), chatModel = $("chat-model"), imageModel = $("image-model");
const submitEndpoint = () => post({ type: "endpoint", value: endpoint.value.trim() });
$("fetch").addEventListener("click", submitEndpoint);
endpoint.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); submitEndpoint(); } });
chatModel.addEventListener("change", () => { post({ type: "chatModel", value: chatModel.value }); updateSummary(); });
imageModel.addEventListener("change", () => { post({ type: "imageModel", value: imageModel.value }); updateSummary(); });
$("apikey").addEventListener("click", () => post({ type: "apiKey" }));
$("save-progress").addEventListener("change", e => post({ type: "saveProgress", value: e.target.checked }));
$("external-control").addEventListener("change", e => post({ type: "externalControl", value: e.target.checked }));

function updateSummary() {
  const host = endpoint.value.replace(/^https?:\/\//, "");
  $("settings-summary").textContent = host ? (chatModel.value || "モデル未選択") + " · " + host : "接続設定（未設定）";
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

function setBusy(on) {
  busy = on;
  cancelling = false;
  document.body.classList.toggle("busy", on);
  action.disabled = false;
  action.title = on ? "中止" : "送信 (Ctrl+Enter)";
  action.setAttribute("aria-label", on ? "中止" : "送信");
  for (const el of [endpoint, $("fetch"), chatModel, imageModel, $("apikey"), $("new-chat"), $("save-progress"), $("external-control")]) el.disabled = on;
  if (!on) { $("status").textContent = ""; $("retry-bar").classList.remove("show"); if (work) finishWork("done"); }
}

const handlers = {
  log: m => onLog(m.text),
  status: m => { $("status").textContent = m.text || ""; },
  busy: m => setBusy(!!m.value),
  retry: m => $("retry-bar").classList.toggle("show", !!m.value),
  // The settings stay folded once an endpoint is set; with none, they open so it can be entered.
  endpoint: m => { endpoint.value = m.value || ""; $("settings").open = !endpoint.value; updateSummary(); },
  models: m => { fill(chatModel, m.chat || [], m.chatSelected); fill(imageModel, m.image || [], m.imageSelected); updateSummary(); },
  apiKeyState: m => { $("apikey-state").textContent = m.value ? "設定済み（Windows 資格情報）" : "未設定"; },
  options: m => { $("save-progress").checked = !!m.saveProgress; $("external-control").checked = !!m.externalControl; },
};
if (host) host.addEventListener("message", e => { const h = handlers[e.data && e.data.type]; if (h) h(e.data); });

updateSummary();
input.focus();
post({ type: "ready" });
