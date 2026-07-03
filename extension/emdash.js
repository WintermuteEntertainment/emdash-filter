// emdash.js — the em dash engine. Direct port of find_emdashes.py:
// find every dash, decide a context-aware replacement, apply edits.
// No dependencies; runs entirely in the browser.
"use strict";

const EM_DASH = "—";

const DASH_NAMES = {
  "—": "EM DASH",
  "–": "EN DASH",
  "―": "HORIZONTAL BAR",
  "−": "MINUS SIGN",
  "--": "DOUBLE HYPHEN",
};

// words that, right after a dash, mean the dash is acting like a comma
const CONNECTIVES = new Set([
  "and", "but", "or", "nor", "yet", "so", "because", "which", "who", "whose",
  "whom", "though", "although", "while", "not", "even", "especially",
  "particularly", "including", "like", "unless", "if", "as", "just",
  "perhaps", "maybe", "namely", "whether", "despite", "unlike",
]);
// words that usually start a full independent clause → period reads best
const CLAUSE_STARTERS = new Set([
  "it", "he", "she", "they", "we", "i", "you", "there", "then", "this",
]);

const SENT_END = /[.!?\n]/;

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ---------------------------------------------------------------- documents
// A "document" = { kind, raw, regions, ltext, lmap }
//   raw     the text we actually edit (plain text, or raw XML for docx)
//   regions [start, end) spans of raw that are real content
//   ltext   readable "logical" text the smart rules analyze
//   lmap    ltext index -> raw index (-1 for inserted breaks); null = identity

function textDocument(text) {
  return { kind: "text", raw: text, regions: [[0, text.length]], ltext: text, lmap: null };
}

function docxDocument(xml) {
  const regions = [];
  const re = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const start = m.index + m[0].length - m[1].length - "</w:t>".length;
    regions.push([start, start + m[1].length]);
  }
  const paraEnds = [];
  let p = xml.indexOf("</w:p>");
  while (p !== -1) { paraEnds.push(p); p = xml.indexOf("</w:p>", p + 1); }

  // Stitch runs into readable text; paragraph breaks become "\n".
  const chunks = [];
  const lmap = [];
  let lastPara = null;
  let pi = 0;
  for (const [rs, rend] of regions) {
    while (pi < paraEnds.length && paraEnds[pi] < rs) pi++;
    if (lastPara !== null && pi !== lastPara) { chunks.push("\n"); lmap.push(-1); }
    lastPara = pi;
    chunks.push(xml.slice(rs, rend));
    for (let k = rs; k < rend; k++) lmap.push(k);
  }
  return { kind: "docx", raw: xml, regions, ltext: chunks.join(""), lmap };
}

function regionAt(doc, pos) {
  let lo = 0, hi = doc.regions.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (doc.regions[mid][0] <= pos) lo = mid; else hi = mid - 1;
  }
  return doc.regions[lo];
}

// ---------------------------------------------------------------- matching

function findMatches(doc, allDashes) {
  const targets = (allDashes ? Object.keys(DASH_NAMES) : [EM_DASH])
    .sort((a, b) => b.length - a.length);
  const pat = new RegExp(targets.map(escapeRegExp).join("|"), "g");
  const recs = [];
  let m;
  while ((m = pat.exec(doc.ltext)) !== null) {
    const ls = m.index, le = ls + m[0].length;
    let s, e;
    if (!doc.lmap) {
      s = ls; e = le;
    } else {
      s = doc.lmap[ls];
      let contiguous = s >= 0;
      for (let k = ls; contiguous && k < le; k++) {
        if (doc.lmap[k] !== s + (k - ls)) contiguous = false;
      }
      if (!contiguous) continue; // split across XML nodes; leave it alone
      e = s + (le - ls);
    }
    const [rs, rend] = regionAt(doc, s);
    recs.push({ ls, le, s, e, grp: m[0], rs, rend });
  }
  return recs;
}

function labelFor(doc, rec) {
  const upto = doc.ltext.slice(0, rec.ls);
  const line = (upto.match(/\n/g) || []).length + 1;
  if (doc.kind === "docx") return "para " + line;
  const col = rec.ls - upto.lastIndexOf("\n");
  return "line " + line + ", col " + col;
}

function contextFor(doc, rec, width) {
  width = width || 60;
  const a = Math.max(0, rec.ls - width);
  const b = Math.min(doc.ltext.length, rec.le + width);
  return {
    pre: (a > 0 ? "…" : "") + doc.ltext.slice(a, rec.ls).replace(/\n/g, " "),
    mid: doc.ltext.slice(rec.ls, rec.le),
    post: doc.ltext.slice(rec.le, b).replace(/\n/g, " ") + (b < doc.ltext.length ? "…" : ""),
  };
}

// ---------------------------------------------------------------- smart engine

function segAfter(ltext, pos) {
  const win = ltext.slice(pos, pos + 300);
  const m = win.search(SENT_END);
  return m === -1 ? win : win.slice(0, m);
}

// One decision per match: { repl: string|null (null = keep), cap: bool, reason }
// Rules, in priority order — this is the whole "AI":
//   numeric range | line-start attribution | trailing dash | paired dashes |
//   connective word | list intro | short afterthought | new clause | elaboration
function smartDecide(ltext, recs) {
  const decisions = new Array(recs.length).fill(null);
  let i = 0;
  while (i < recs.length) {
    const { ls, le } = recs[i];
    const prev = ls > 0 ? ltext[ls - 1] : "\n";
    const nxt = le < ltext.length ? ltext[le] : "\n";
    const seg = segAfter(ltext, le);
    const words = seg.match(/[A-Za-z0-9']+/g) || [];
    const first = words.length ? words[0].toLowerCase() : "";

    let j = ls;
    while (j > 0 && ltext[j - 1] === " ") j--;

    if (/[0-9]/.test(prev) && /[0-9]/.test(nxt)) {
      decisions[i] = { repl: "-", cap: false, reason: "number range" };
    } else if (j === 0 || ltext[j - 1] === "\n") {
      decisions[i] = { repl: null, cap: false, reason: "starts the line (attribution/bullet)" };
    } else if (!words.length) {
      decisions[i] = { repl: null, cap: false, reason: "trailing dash / end of sentence" };
    } else if (i + 1 < recs.length
        && recs[i + 1].ls - le <= 80
        && !SENT_END.test(ltext.slice(le, recs[i + 1].ls))) {
      const inner = ltext.slice(le, recs[i + 1].ls);
      if (inner.includes(",")) {
        decisions[i] = { repl: " (", cap: false, reason: "paired aside with inner commas" };
        decisions[i + 1] = { repl: ") ", cap: false, reason: "closes the aside" };
      } else {
        decisions[i] = { repl: ", ", cap: false, reason: "paired aside" };
        decisions[i + 1] = { repl: ", ", cap: false, reason: "closes the aside" };
      }
      i += 2;
      continue;
    } else if (CONNECTIVES.has(first)) {
      decisions[i] = { repl: ", ", cap: false, reason: "followed by '" + words[0] + "'" };
    } else if (seg.includes(",") && /\b(and|or)\b/.test(seg)) {
      decisions[i] = { repl: ": ", cap: false, reason: "introduces a list" };
    } else if (words.length <= 5) {
      decisions[i] = { repl: ", ", cap: false, reason: "short afterthought" };
    } else if (CLAUSE_STARTERS.has(first)) {
      decisions[i] = { repl: ". ", cap: true, reason: "new sentence starting with '" + words[0] + "'" };
    } else {
      decisions[i] = { repl: ": ", cap: false, reason: "introduces an elaboration" };
    }
    i++;
  }
  return decisions;
}

function describeReplacement(repl, cap) {
  const names = {
    ", ": "comma “, ”", " - ": "hyphen “ - ”", " ": "space", "": "delete",
    ": ": "colon “: ”", ". ": "period “. ”", "-": "hyphen “-”",
    " (": "open paren “ (”", ") ": "close paren “) ”",
  };
  if (repl === null || repl === undefined) return "keep";
  let d = names[repl] !== undefined ? names[repl] : "custom “" + repl + "”";
  if (cap) d += " + capitalize";
  return d;
}

// ---------------------------------------------------------------- editing

function xmlEscapeText(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// decisions: array aligned with recs, each { repl, cap } or null (= keep).
// Returns the edited raw text. Spaces around the dash are eaten so ", "
// lands as "word, word" whether or not the dash was spaced.
function applyDecisions(doc, recs, decisions) {
  const edits = []; // [start, end, replacement, regionStart, regionEnd, eatSpaces]
  recs.forEach((rec, idx) => {
    const dec = decisions[idx];
    if (!dec || dec.repl === null || dec.repl === undefined) return;
    edits.push([rec.s, rec.e, dec.repl, rec.rs, rec.rend, true]);
    if (dec.cap) { // uppercase the first letter after the dash (now a new sentence)
      let k = rec.le;
      while (k < doc.ltext.length && doc.ltext[k] === " ") k++;
      const ch = doc.ltext[k];
      if (ch && ch === ch.toLowerCase() && ch !== ch.toUpperCase()) {
        const lraw = doc.lmap ? doc.lmap[k] : k;
        if (lraw >= 0) {
          const [rs, rend] = regionAt(doc, lraw);
          edits.push([lraw, lraw + 1, ch.toUpperCase(), rs, rend, false]);
        }
      }
    }
  });
  edits.sort((a, b) => b[0] - a[0]); // back-to-front so offsets stay valid
  let text = doc.raw;
  for (let [s, e, repl, rs, rend, eat] of edits) {
    if (eat) {
      while (s > rs && text[s - 1] === " ") s--;
      while (e < rend && text[e] === " ") e++;
    }
    const ins = doc.kind === "docx" ? xmlEscapeText(repl) : repl;
    text = text.slice(0, s) + ins + text.slice(e);
  }
  return text;
}

// Node (for tests)
if (typeof module !== "undefined") {
  module.exports = {
    EM_DASH, DASH_NAMES, textDocument, docxDocument, findMatches,
    labelFor, contextFor, smartDecide, describeReplacement, applyDecisions,
  };
}
