"use strict";
// The task pane's text, by message code. Every sentence the user sees is in locales/<language>.json
// ({"CODE": "text with {0} {1}"}); the page and the add-in refer to it only by its code, and the add-in
// sends codes and values, never sentences. ja.json is the base: it has every code, and a language file
// that lacks one shows the Japanese text (scripts/test/i18n.py checks that all files have the same codes
// and placeholders, and that every code used exists). Codes are UPPER_SNAKE_CASE.
//
// A message is a code, or {code, args}. A value in args is shown as it is, or is itself a message, or is
// {join: "SEPARATOR_CODE", items: [...]} - its items, each shown the same way, joined by that code's text.
//
// The add-in (AiTaskPaneControl.cs) reads the same files for the text it draws itself (dialogs).
const i18n = (() => {
  let base = {}, table = {};

  async function read(lang) {
    const r = await fetch("locales/" + lang + ".json");
    if (!r.ok) throw new Error("no texts for " + lang);
    return r.json();
  }

  // `lang` is the one the add-in chose among the files there are (index.html?lang=..., Texts.LanguageFor);
  // should its file not load, Japanese.
  async function load(lang) {
    base = await read("ja");
    table = base;
    if (lang && lang !== "ja") {
      try { table = Object.assign({}, base, await read(lang)); } catch (e) { lang = "ja"; }
    }
    document.documentElement.lang = lang || "ja";
  }

  function text(code, args) {
    const s = Object.prototype.hasOwnProperty.call(table, code) ? table[code] : code;
    return s.replace(/\{(\d+)\}/g, (_, i) => value(args[+i]));
  }
  function value(v) {
    if (v == null) return "";
    if (typeof v === "object" && v.code) return t(v);
    if (typeof v === "object" && v.join) return (v.items || []).map(value).join(text(v.join, []));
    return String(v);
  }
  // t("CODE"), t({code, args}) or t({join, items})
  function t(msg) {
    if (typeof msg === "string") return text(msg, []);
    if (msg && msg.code) return text(msg.code, msg.args || []);
    return msg && msg.join ? value(msg) : "";
  }
  function fill(code, ...args) { return text(code, args); }

  // data-i18n (text), data-i18n-placeholder, data-i18n-title, data-i18n-label (aria-label).
  function applyText(root) {
    for (const el of root.querySelectorAll("[data-i18n]")) el.textContent = t(el.dataset.i18n);
    for (const el of root.querySelectorAll("[data-i18n-placeholder]")) el.placeholder = t(el.dataset.i18nPlaceholder);
    for (const el of root.querySelectorAll("[data-i18n-title]")) el.title = t(el.dataset.i18nTitle);
    for (const el of root.querySelectorAll("[data-i18n-label]")) el.setAttribute("aria-label", t(el.dataset.i18nLabel));
  }

  return { load, t, fill, applyText };
})();
