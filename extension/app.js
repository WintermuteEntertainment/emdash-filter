// app.js — UI glue: file in, decision cards, cleaned file out.
"use strict";

const $ = (id) => document.getElementById(id);

// Dropdown options: value -> { repl, cap, label }
const OPTIONS = [
  { value: "keep",   repl: null, cap: false, label: "keep the dash" },
  { value: "comma",  repl: ", ", cap: false, label: "comma  “, ”" },
  { value: "colon",  repl: ": ", cap: false, label: "colon  “: ”" },
  { value: "period", repl: ". ", cap: true,  label: "period “. ” (new sentence)" },
  { value: "hyphen", repl: " - ", cap: false, label: "hyphen “ - ”" },
  { value: "range",  repl: "-",  cap: false, label: "bare hyphen “-”" },
  { value: "openp",  repl: " (", cap: false, label: "open paren “ (”" },
  { value: "closep", repl: ") ", cap: false, label: "close paren “) ”" },
  { value: "space",  repl: " ",  cap: false, label: "just a space" },
  { value: "delete", repl: "",   cap: false, label: "delete entirely" },
  { value: "custom", repl: "",   cap: false, label: "custom text…" },
];

function optionFor(repl, cap) {
  if (repl === null || repl === undefined) return "keep";
  const hit = OPTIONS.find((o) => o.repl === repl && o.cap === cap && o.value !== "custom");
  return hit ? hit.value : "custom";
}

const state = {
  fileName: null,
  fileBytes: null,  // ArrayBuffer of the original file
  isDocx: false,
  zipFiles: null,   // Map for docx
  doc: null,
  recs: [],
  suggestions: [],
};

// ------------------------------------------------------------------ file in

async function handleFile(file) {
  $("error").hidden = true;
  $("doneNote").hidden = true;
  const name = file.name;
  const ext = (name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
  if (ext === ".pdf" || ext === ".doc") {
    return showError(ext + " isn't supported here — open it in Word and save as .docx, then try again.");
  }
  try {
    state.fileName = name;
    state.fileBytes = await file.arrayBuffer();
    state.isDocx = ext === ".docx";
    await analyze();
  } catch (err) {
    showError("Couldn't read that file: " + err.message);
  }
}

async function analyze() {
  if (state.isDocx) {
    state.zipFiles = await unzip(state.fileBytes);
    const xmlBytes = state.zipFiles.get("word/document.xml");
    if (!xmlBytes) throw new Error("no word/document.xml inside — is this a real .docx?");
    state.doc = docxDocument(new TextDecoder().decode(xmlBytes));
  } else {
    state.doc = textDocument(new TextDecoder().decode(state.fileBytes));
  }
  state.recs = findMatches(state.doc, $("allDashes").checked);
  state.suggestions = smartDecide(state.doc.ltext, state.recs);
  render();
}

function showError(msg) {
  const el = $("error");
  el.textContent = msg;
  el.hidden = false;
  $("results").hidden = true;
}

// ------------------------------------------------------------------ render

function el(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
}

function render() {
  const n = state.recs.length;
  $("results").hidden = false;
  $("summary").textContent = n === 0
    ? `No em dashes found in ${state.fileName}. Clean!`
    : `${n} dash${n === 1 ? "" : "es"} found in ${state.fileName} — each has a suggested fix below.`;
  const list = $("list");
  list.replaceChildren();
  state.recs.forEach((rec, i) => {
    const sug = state.suggestions[i];
    const sugVal = optionFor(sug.repl, sug.cap);
    const ctx = contextFor(state.doc, rec);
    const card = el("div", "card");

    const loc = el("div", "loc");
    loc.append(
      el("span", "num", "#" + (i + 1)),
      ` · ${labelFor(state.doc, rec)} · ${DASH_NAMES[rec.grp] || "DASH"}`,
      el("span", "reason",
        `suggested: ${describeReplacement(sug.repl, sug.cap)} — ${sug.reason}`));

    const ctxDiv = el("div", "ctx");
    ctxDiv.append(ctx.pre, el("mark", null, ctx.mid), ctx.post);

    const sel = el("select");
    sel.dataset.i = i;
    for (const o of OPTIONS) {
      const opt = el("option", null, o.label + (o.value === sugVal ? "  ✓ suggested" : ""));
      opt.value = o.value;
      if (o.value === sugVal) opt.selected = true;
      sel.append(opt);
    }

    const custom = el("input", "custom");
    custom.dataset.i = i;
    custom.placeholder = "replacement text";
    custom.hidden = sugVal !== "custom";
    if (sugVal === "custom") custom.value = sug.repl;
    sel.addEventListener("change", () => {
      custom.hidden = sel.value !== "custom";
      if (sel.value === "custom") custom.focus();
    });

    card.append(loc, ctxDiv, sel, custom);
    list.append(card);
  });
}

// ------------------------------------------------------------------ decisions

function currentDecisions() {
  const decisions = new Array(state.recs.length).fill(null);
  document.querySelectorAll("#list select").forEach((sel) => {
    const i = Number(sel.dataset.i);
    if (sel.value === "keep") { decisions[i] = null; return; }
    if (sel.value === "custom") {
      const custom = document.querySelector(`input.custom[data-i="${i}"]`);
      decisions[i] = { repl: custom.value, cap: false };
      return;
    }
    const opt = OPTIONS.find((o) => o.value === sel.value);
    decisions[i] = { repl: opt.repl, cap: opt.cap };
  });
  return decisions;
}

function setAll(valueFor) {
  document.querySelectorAll("#list select").forEach((sel) => {
    const i = Number(sel.dataset.i);
    sel.value = valueFor(i);
    const custom = document.querySelector(`input.custom[data-i="${i}"]`);
    custom.hidden = sel.value !== "custom";
  });
}

// ------------------------------------------------------------------ file out

function buildOutput() {
  const decisions = currentDecisions();
  const changed = decisions.filter((d) => d !== null).length;
  const edited = applyDecisions(state.doc, state.recs, decisions);
  const dot = state.fileName.lastIndexOf(".");
  const outName = dot === -1
    ? state.fileName + ".nodash"
    : state.fileName.slice(0, dot) + ".nodash" + state.fileName.slice(dot);
  let blob;
  if (state.isDocx) {
    const files = new Map(state.zipFiles);
    files.set("word/document.xml", new TextEncoder().encode(edited));
    blob = new Blob([zipStore(files)], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  } else {
    blob = new Blob([edited], { type: "text/plain;charset=utf-8" });
  }
  return { blob, outName, changed, total: state.recs.length };
}

function download() {
  if (!state.recs.length) return;
  const { blob, outName, changed, total } = buildOutput();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = outName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  const note = $("doneNote");
  note.textContent = `Saved ${outName} to your Downloads (${changed} of ${total} replaced, ` +
    `${total - changed} kept). Your original file was not touched.`;
  note.hidden = false;
}

// ------------------------------------------------------------------ wiring

$("fileInput").addEventListener("change", (e) => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});
const drop = $("drop");
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("dragover"); });
drop.addEventListener("dragleave", () => drop.classList.remove("dragover"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("dragover");
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
$("allDashes").addEventListener("change", () => { if (state.fileBytes) analyze(); });
$("acceptAll").addEventListener("click", () =>
  setAll((i) => optionFor(state.suggestions[i].repl, state.suggestions[i].cap)));
$("keepAll").addEventListener("click", () => setAll(() => "keep"));
$("download").addEventListener("click", download);
