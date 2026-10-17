// Converts one document the user attached into text the model can read line by line, off the page's
// thread (anydoc runs synchronously). Started by docs.js; one message in, one message out:
//   in:  { name, bytes: Uint8Array }
//   out: { ok: true, format, text, images: [{ index, mediaType, data: Uint8Array }] }
//     or { ok: false, reason: "unsupported" | "encrypted" | "needsOcr" | "malformed" | "resourceLimit" | "missingPart" | "error", detail }
//
// The text is Markdown-like and dense: one block per line (no blank lines between them), headings as
// "# ", list items as "- " / "1. ", tables as "| a | b |" rows, so a line number points at one paragraph,
// heading, item or table row. An embedded image is marked where it sits as [[img:N]] (N = its index in
// images); the add-in turns that into the image's name (d1_img3). PDFs do not come here (pdfdoc.js);
// CSV has no document model in anydoc, so its text is anydoc's Markdown as is.
import init, { formatFromBytes, toDocument, toMarkdownBytes } from "./vendor/anydoc/anydoc_wasm.js";

const ready = init();

self.onmessage = async (e) => {
  const { name, bytes } = e.data;
  try {
    await ready;
    self.postMessage(convert(name, bytes), transferables);
  } catch (err) {
    self.postMessage({ ok: false, reason: (err && err.code) || "error", detail: String((err && err.message) || err) });
  } finally {
    transferables = [];
  }
};

let transferables = [];

function convert(name, bytes) {
  const ext = (name.match(/\.([^.]+)$/) || [, ""])[1].toLowerCase();
  let format = formatFromBytes(bytes);
  if (!format && ext === "csv") format = "csv";
  if (!format && ["txt", "md", "markdown", "text"].includes(ext)) {
    return { ok: true, format: "text", text: dense(decodeText(bytes)), images: [] };
  }
  if (!format || format === "pdf") return { ok: false, reason: "unsupported", detail: format ? "PDFs are read by pdfdoc.js" : "unknown format" };
  if (format === "csv") {
    return { ok: true, format, text: dense(toMarkdownBytes(bytes, format)), images: [] };
  }
  const doc = toDocument(bytes, format);
  const text = render(doc);
  const images = doc.assets
    .filter(a => a.mediaType && a.mediaType.startsWith("image/"))
    .map(a => { transferables.push(a.data.buffer); return { index: a.id, mediaType: a.mediaType, data: a.data }; });
  return { ok: true, format, text, images };
}

// UTF-8 when it is valid UTF-8, otherwise Shift_JIS (older Japanese text files).
function decodeText(bytes) {
  try { return new TextDecoder("utf-8", { fatal: true }).decode(bytes); }
  catch { return new TextDecoder("shift_jis").decode(bytes); }
}

// Markdown from anydoc (PDF, CSV) or a text file: drop the blank lines and trailing spaces.
function dense(markdown) {
  return markdown.replace(/\r\n?/g, "\n").split("\n").map(l => l.replace(/\s+$/, "")).filter(l => l.length > 0).join("\n");
}

// ---- document model -> text -----------------------------------------------------------------------

function render(doc) {
  const out = [];
  blocks(doc.blocks, out, "");
  for (const n of doc.notes || []) {
    const body = [];
    blocks(n.blocks, body, "");
    out.push(`[^${n.id}]: ` + body.join(" "));
  }
  return out.filter(l => l.trim().length > 0).join("\n");
}

function blocks(list, out, indent) {
  for (const b of list || []) block(b, out, indent);
}

function block(b, out, indent) {
  switch (b.kind) {
    case "heading":
      out.push(indent + "#".repeat(Math.min(Math.max(b.level || 1, 1), 6)) + " " + oneLine(inlines(b.content)));
      break;
    case "paragraph":
      for (const line of inlines(b.content).split("\n")) out.push(indent + line);
      break;
    case "list": {
      const l = b.list;
      l.items.forEach((item, i) => {
        const marker = item.markerLabel ? item.markerLabel + " " : l.marker === "bullet" ? "- " : `${l.start + i}. `;
        const inner = [];
        blocks(item.blocks, inner, "");
        if (inner.length === 0) inner.push("");
        out.push(indent + marker + inner[0]);
        for (const rest of inner.slice(1)) out.push(indent + "  " + rest);
      });
      break;
    }
    case "table":
      table(b.table, out, indent);
      break;
    case "blockQuote": {
      const inner = [];
      blocks(b.blocks, inner, "");
      for (const line of inner) out.push(indent + "> " + line);
      break;
    }
    case "codeBlock":
      out.push(indent + "```" + (b.lang || ""));
      for (const line of (b.text || "").split("\n")) out.push(indent + line);
      out.push(indent + "```");
      break;
    case "math":
      out.push(indent + "$$ " + oneLine(b.text || "") + " $$");
      break;
    case "rule":
      out.push(indent + "---");
      break;
  }
}

// A data table becomes Markdown rows; a layout table (text boxes, positioning) is just its cells' content.
function table(t, out, indent) {
  if (t.kind === "layout") {
    for (const row of t.grid) for (const slot of row) if (slot.kind === "origin") blocks(slot.cell.blocks, out, indent);
    return;
  }
  const rows = t.grid.map(row => row.map(slot => {
    if (slot.kind !== "origin") return "";
    const inner = [];
    blocks(slot.cell.blocks, inner, "");
    return inner.join(" / ").replace(/\|/g, "\\|");
  }));
  if (rows.length === 0) return;
  const cols = Math.max(...rows.map(r => r.length));
  const line = r => indent + "| " + Array.from({ length: cols }, (_, i) => r[i] || "").join(" | ") + " |";
  out.push(line(rows[0]));
  out.push(indent + "|" + " --- |".repeat(cols));
  for (const r of rows.slice(1)) out.push(line(r));
}

function inlines(list) {
  let s = "";
  for (const x of list || []) {
    switch (x.kind) {
      case "text": {
        const t = x.text || "";
        s += x.style && x.style.code ? "`" + t + "`" : x.style && x.style.bold && t.trim() ? "**" + t + "**" : t;
        break;
      }
      case "link": {
        const t = inlines(x.content);
        s += x.target && x.target.kind === "external" && x.target.value !== t ? `${t} (${x.target.value})` : t;
        break;
      }
      case "image":
        if (x.source && x.source.kind === "asset") s += `[[img:${x.source.assetId}]]` + (x.alt ? ` (${x.alt})` : "");
        else if (x.alt) s += `[image: ${x.alt}]`;
        break;
      case "noteRef": s += `[^${x.noteId}]`; break;
      case "lineBreak": s += "\n"; break;
      case "math": s += "$" + (x.text || "") + "$"; break;
      case "checkbox": s += x.checked ? "[x] " : "[ ] "; break;
    }
  }
  return s;
}

function oneLine(s) {
  return s.replace(/\s*\n\s*/g, " ");
}
