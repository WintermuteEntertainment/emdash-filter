#!/usr/bin/env python3
"""
find_emdashes.py — scan a document for em dashes, and optionally filter them out.

Modes:
    (default)          just list every em dash with location + highlighted context
    -s, --smart        auto-fix: pick a context-aware replacement for each dash
    -r, --remove-all   replace every em dash with the same text (--replace)
    -i, --interactive  step through each em dash one at a time and decide
                       (each one comes with a smart suggestion — Enter accepts it)

The original file is NEVER modified. Cleaned output goes to a new file
(yourdoc.nodash.docx next to the original) unless you pass -o.
Add --dry to preview decisions without writing anything.

Supports:
    .txt / .md / any plain-text file   scan + filter  (no dependencies)
    .docx                              scan + filter  (no dependencies)
    .pdf                               scan only      (requires: pip install pypdf)
"""

import argparse
import bisect
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape as xml_escape

EM_DASH = "—"

DASHES = {
    "—": "EM DASH",
    "–": "EN DASH",
    "―": "HORIZONTAL BAR",
    "−": "MINUS SIGN",
    "--": "DOUBLE HYPHEN",
}

CONTEXT = 40   # characters of context shown on each side of a match
SENT_END = re.compile(r"[.!?\n]")

# words that, right after a dash, mean the dash is acting like a comma
CONNECTIVES = {
    "and", "but", "or", "nor", "yet", "so", "because", "which", "who", "whose",
    "whom", "though", "although", "while", "not", "even", "especially",
    "particularly", "including", "like", "unless", "if", "as", "just",
    "perhaps", "maybe", "namely", "whether", "despite", "unlike",
}
# words that usually start a full independent clause → period reads best
CLAUSE_STARTERS = {"it", "he", "she", "they", "we", "i", "you", "there", "then", "this"}


def enable_ansi():
    """Make ANSI colors work in the classic Windows console."""
    if os.name == "nt":
        os.system("")  # side effect: enables VT processing


# ---------------------------------------------------------------- documents

class TextDoc:
    """Plain text. .text is the whole buffer; the single region is all of it."""
    editable = True

    def __init__(self, path):
        # newline="" preserves the file's own line endings (LF vs CRLF)
        with open(path, encoding="utf-8", errors="replace", newline="") as f:
            self.text = f.read()
        self.regions = [(0, len(self.text))]

    def logical(self):
        """(readable_text, map_to_raw_offsets). Identity for plain text."""
        return self.text, None

    def label(self, pos):
        line = self.text.count("\n", 0, pos) + 1
        col = pos - (self.text.rfind("\n", 0, pos) + 1) + 1
        return f"line {line}, col {col}"

    def context(self, s, e, hi, reset):
        a = max(0, s - CONTEXT)
        b = min(len(self.text), e + CONTEXT)
        snip = self.text[a:s] + hi + self.text[s:e] + reset + self.text[e:b]
        return (("…" if a > 0 else "")
                + snip.replace("\r", "").replace("\n", " ")
                + ("…" if b < len(self.text) else ""))

    def prepare_replacement(self, repl):
        return repl

    def save(self, out_path):
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            f.write(self.text)


class PdfDoc(TextDoc):
    """PDF via pypdf. Scan only — no clean way to write text back into a PDF."""
    editable = False

    def __init__(self, path):
        try:
            from pypdf import PdfReader
        except ImportError:
            sys.exit("PDF support needs pypdf. Install it with:  pip install pypdf")
        reader = PdfReader(path)
        parts, self.page_starts, pos = [], [], 0
        for page in reader.pages:
            t = page.extract_text() or ""
            self.page_starts.append(pos)
            parts.append(t)
            pos += len(t) + 1
        self.text = "\n".join(parts)
        self.regions = [(0, len(self.text))]

    def label(self, pos):
        page = bisect.bisect_right(self.page_starts, pos)
        col = pos - (self.text.rfind("\n", 0, pos) + 1) + 1
        line = self.text.count("\n", self.page_starts[page - 1], pos) + 1
        return f"page {page}, line {line}, col {col}"


class DocxDoc:
    """Word .docx. We search/edit only inside <w:t> text nodes of document.xml,
    then rezip everything else untouched."""
    editable = True

    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        self.text = self.parts["word/document.xml"].decode("utf-8", errors="replace")
        self.regions = [
            (m.start(2), m.end(2))
            for m in re.finditer(r"(<w:t[^>]*>)(.*?)(</w:t>)", self.text, re.S)
        ]

    def logical(self):
        """Stitch the <w:t> runs into readable text (runs in the same paragraph
        are contiguous; paragraph breaks become newlines) plus a char-by-char
        map back to raw XML offsets, so smart analysis can see whole sentences."""
        para_ends = [m.start() for m in re.finditer(r"</w:p>", self.text)]
        chunks, lmap, last_para = [], [], None
        for rs, rend in self.regions:
            para = bisect.bisect_right(para_ends, rs)
            if last_para is not None and para != last_para:
                chunks.append("\n")
                lmap.append(None)
            last_para = para
            chunks.append(self.text[rs:rend])
            lmap.extend(range(rs, rend))
        return "".join(chunks), lmap

    def label(self, pos):
        return f"para {self.text.count('</w:p>', 0, pos) + 1}"

    def context(self, s, e, hi, reset):
        # Grab a wide window, strip the XML tags, then trim to CONTEXT chars.
        w0 = max(0, s - 1500)
        pre = self.text[w0:s]
        if w0 > 0:  # window may start mid-tag; drop the partial tag
            cut = pre.find(">")
            if cut != -1 and "<" not in pre[:cut]:
                pre = pre[cut + 1:]
        post = self.text[e:e + 1500]
        lt = post.rfind("<")
        if lt != -1 and ">" not in post[lt:]:  # window may end mid-tag
            post = post[:lt]
        pre = re.sub(r"<[^>]+>", "", pre)[-CONTEXT:]
        post = re.sub(r"<[^>]+>", "", post)[:CONTEXT]
        return ("…" + pre + hi + self.text[s:e] + reset + post + "…").replace("\n", " ")

    def prepare_replacement(self, repl):
        return xml_escape(repl)

    def save(self, out_path):
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                if name == "word/document.xml":
                    data = self.text.encode("utf-8")
                z.writestr(name, data)


def load_doc(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return DocxDoc(path)
    if ext == ".pdf":
        return PdfDoc(path)
    if ext == ".doc":
        sys.exit("Old-style .doc isn't supported — open it in Word and save as .docx first.")
    return TextDoc(path)


# ---------------------------------------------------------------- matching

def find_matches(doc, pattern):
    """Search the readable (logical) text; map every hit back to raw offsets.
    Returns (logical_text, logical_map, [match dicts])."""
    ltext, lmap = doc.logical()
    starts = [r[0] for r in doc.regions]
    recs = []
    for m in pattern.finditer(ltext):
        ls, le = m.span()
        if lmap is None:
            s, e = ls, le
        else:
            raws = [lmap[k] for k in range(ls, le)]
            if None in raws or raws != list(range(raws[0], raws[0] + len(raws))):
                continue  # match split across XML nodes; leave it alone
            s, e = raws[0], raws[-1] + 1
        ri = bisect.bisect_right(starts, s) - 1
        rs, rend = doc.regions[ri]
        recs.append({"ls": ls, "le": le, "s": s, "e": e,
                     "grp": m.group(), "rs": rs, "rend": rend})
    return ltext, lmap, recs


# ---------------------------------------------------------------- smart engine

def seg_after(ltext, pos):
    """Text from pos to the end of the current sentence (or 300 chars)."""
    m = SENT_END.search(ltext, pos, pos + 300)
    end = m.start() if m else min(len(ltext), pos + 300)
    return ltext[pos:end]


def smart_decide(ltext, recs):
    """One decision per match: (replacement or None-to-keep, capitalize_next, reason).
    Rules, in priority order — this is the whole 'AI':
      numeric range | line-start attribution | trailing dash | paired dashes |
      connective word | list intro | short afterthought | new clause | elaboration"""
    n = len(recs)
    decisions = [None] * n
    i = 0
    while i < n:
        ls, le = recs[i]["ls"], recs[i]["le"]
        prev = ltext[ls - 1] if ls > 0 else "\n"
        nxt = ltext[le] if le < len(ltext) else "\n"
        seg = seg_after(ltext, le)
        words = re.findall(r"[A-Za-z0-9']+", seg)
        first = words[0].lower() if words else ""

        j = ls
        while j > 0 and ltext[j - 1] == " ":
            j -= 1

        if prev.isdigit() and nxt.isdigit():
            decisions[i] = ("-", False, "number range")
        elif j == 0 or ltext[j - 1] == "\n":
            decisions[i] = (None, False, "starts the line (attribution/bullet)")
        elif not words:
            decisions[i] = (None, False, "trailing dash / end of sentence")
        elif i + 1 < n and len(ltext[le:recs[i + 1]["ls"]]) <= 80 \
                and not SENT_END.search(ltext, le, recs[i + 1]["ls"]):
            inner = ltext[le:recs[i + 1]["ls"]]
            if "," in inner:
                decisions[i] = (" (", False, "paired aside with inner commas")
                decisions[i + 1] = (") ", False, "closes the aside")
            else:
                decisions[i] = (", ", False, "paired aside")
                decisions[i + 1] = (", ", False, "closes the aside")
            i += 2
            continue
        elif first in CONNECTIVES:
            decisions[i] = (", ", False, f"followed by '{words[0]}'")
        elif "," in seg and re.search(r"\b(and|or)\b", seg):
            decisions[i] = (": ", False, "introduces a list")
        elif len(words) <= 5:
            decisions[i] = (", ", False, "short afterthought")
        elif first in CLAUSE_STARTERS:
            decisions[i] = (". ", True, f"new sentence starting with '{words[0]}'")
        else:
            decisions[i] = (": ", False, "introduces an elaboration")
        i += 1
    return decisions


def describe(repl, cap):
    names = {None: "keep", ", ": "comma ', '", " - ": "hyphen ' - '", " ": "space",
             "": "delete", ": ": "colon ': '", ". ": "period '. '", "-": "hyphen '-'",
             " (": "open paren ' ('", ") ": "close paren ') '"}
    d = names.get(repl, f"custom {repl!r}")
    if cap:
        d += " + capitalize"
    return d


# ---------------------------------------------------------------- editing

def build_edits(doc, ltext, lmap, recs, decisions):
    """Turn (replacement, capitalize) decisions into raw-text edits.
    Edit = (start, end, replacement, region_start, region_end, eat_spaces)."""
    starts = [r[0] for r in doc.regions]
    edits = []
    for rec, dec in zip(recs, decisions):
        if dec is None or dec[0] is None:
            continue
        repl, cap = dec[0], dec[1]
        edits.append((rec["s"], rec["e"], repl, rec["rs"], rec["rend"], True))
        if cap:  # uppercase the first letter after the dash (now a new sentence)
            k = rec["le"]
            while k < len(ltext) and ltext[k] == " ":
                k += 1
            if k < len(ltext) and ltext[k].islower():
                lraw = k if lmap is None else lmap[k]
                if lraw is not None:
                    ri = bisect.bisect_right(starts, lraw) - 1
                    rs, rend = doc.regions[ri]
                    edits.append((lraw, lraw + 1, ltext[k].upper(), rs, rend, False))
    return edits


def apply_edits(doc, edits):
    """Applied back-to-front so earlier offsets stay valid. Spaces around the
    dash are eaten so ', ' lands as 'word, word' whether or not it was spaced."""
    text = doc.text
    for s, e, repl, rs, rend, eat in sorted(edits, key=lambda x: -x[0]):
        if eat:
            while s > rs and text[s - 1] == " ":
                s -= 1
            while e < rend and text[e] == " ":
                e += 1
        text = text[:s] + doc.prepare_replacement(repl) + text[e:]
    doc.text = text


def default_output(path):
    root, ext = os.path.splitext(path)
    cand, n = f"{root}.nodash{ext}", 1
    while os.path.exists(cand):
        n += 1
        cand = f"{root}.nodash{n}{ext}"
    return cand


# ---------------------------------------------------------------- modes

def summarize(recs, targets):
    print("-" * 60)
    if not recs:
        print("No em dashes found. Clean!")
        return
    counts = {}
    for rec in recs:
        name = targets[rec["grp"]]
        counts[name] = counts.get(name, 0) + 1
    print(f"Total: {len(recs)}")
    for name, cnt in sorted(counts.items(), key=lambda kv: -kv[1]):
        print(f"  {name}: {cnt}")


def scan(doc, recs, targets, hi, reset):
    for i, rec in enumerate(recs, 1):
        print(f"[{i:>3}] {doc.label(rec['s'])} ({targets[rec['grp']]}):")
        print(f"      {doc.context(rec['s'], rec['e'], hi, reset)}\n")
    summarize(recs, targets)


def smart_report(doc, recs, suggs, hi, reset):
    for i, (rec, (repl, cap, reason)) in enumerate(zip(recs, suggs), 1):
        print(f"[{i:>3}] {doc.label(rec['s'])}:")
        print(f"      {doc.context(rec['s'], rec['e'], hi, reset)}")
        print(f"      → {describe(repl, cap)}   [{reason}]\n")


def read_choice(prompt):
    # strip BOM/null chars that show up when input is piped in on Windows
    return input(prompt).replace("﻿", "").replace("\x00", "").strip()


REPLS = {"c": ", ", "h": " - ", "s": " ", "x": ""}
MENU = ("      [Enter/a] accept suggestion   [k]eep   [c]omma   [h]yphen ' - '   "
        "[s]pace   [x]delete   [e]custom\n"
        "      [A] accept ALL suggestions   K/C/H/S/X/E = same for all remaining   "
        "[q] keep rest & save   [!] abort")


def interactive(doc, recs, suggs, targets, hi, reset):
    """Returns a decisions list aligned with recs (None = keep), or None if aborted."""
    decisions = [None] * len(recs)
    bulk = None  # "suggest" or a (repl, cap) tuple applied to everything remaining
    for i, rec in enumerate(recs):
        srepl, scap, sreason = suggs[i]
        if bulk is not None:
            decisions[i] = (srepl, scap) if bulk == "suggest" else bulk
            continue
        print(f"\n[{i + 1}/{len(recs)}] {doc.label(rec['s'])} ({targets[rec['grp']]})"
              f"   suggested: {describe(srepl, scap)} [{sreason}]")
        print(f"      {doc.context(rec['s'], rec['e'], hi, reset)}")
        print(MENU)
        while True:
            try:
                ans = read_choice("      > ")
            except (EOFError, KeyboardInterrupt):
                print("\nAborted — nothing written.")
                return None
            if ans in ("", "a"):
                decisions[i] = (srepl, scap)
                break
            if ans == "A":
                bulk = "suggest"
                decisions[i] = (srepl, scap)
                break
            if ans == "k":
                break
            if ans == "K":
                print("      Keeping this and all remaining.")
                return decisions
            if ans == "q":
                print("      Keeping all remaining; saving decisions so far.")
                return decisions
            if ans == "!":
                print("      Aborted — nothing written.")
                return None
            if ans in REPLS:
                decisions[i] = (REPLS[ans], False)
                break
            if ans in ("C", "H", "S", "X"):
                bulk = (REPLS[ans.lower()], False)
                decisions[i] = bulk
                break
            if ans in ("e", "E"):
                custom = input("      replacement text (spaces around the dash are "
                               "removed, so include your own): ").replace("﻿", "")
                decisions[i] = (custom, False)
                if ans == "E":
                    bulk = decisions[i]
                break
            print("      ? — pick one of the letters above")
    return decisions


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(
        description="Scan a document for em dashes; optionally filter them out.\n"
                    "The original file is NEVER modified — cleaned output goes to a\n"
                    "new copy (yourdoc.nodash.docx) unless you pick a spot with -o.",
        epilog="examples:\n"
               "  find_emdashes.py essay.docx                  list every em dash\n"
               "  find_emdashes.py essay.docx -s               smart auto-fix, save a copy\n"
               "  find_emdashes.py essay.docx -s --dry         preview smart decisions only\n"
               "  find_emdashes.py essay.docx -i               review each one (Enter = accept)\n"
               "  find_emdashes.py essay.docx -r --replace ' - '   dumb replace-all\n"
               "\n"
               "smart mode picks per dash: paired dashes -> commas/parens, before\n"
               "'but/which/and...' -> comma, list intro -> colon, short afterthought\n"
               "-> comma, new sentence -> period, elaboration -> colon.\n"
               "\n"
               "interactive keys (per em dash):\n"
               "  Enter/a accept suggestion   k keep   c comma   h hyphen   s space\n"
               "  x delete   e custom   A accept all   K/C/H/S/X/E same-for-all\n"
               "  q keep rest & save   ! abort\n"
               "\n"
               "supports: .txt/.md/plain text, .docx  |  .pdf is scan-only (pip install pypdf)",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", help="document to scan (.txt/.md/anything, .docx, .pdf)")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("-s", "--smart", action="store_true",
                      help="auto-pick a context-aware replacement for each dash")
    mode.add_argument("-r", "--remove-all", action="store_true",
                      help="replace every match with the same text (see --replace)")
    mode.add_argument("-i", "--interactive", action="store_true",
                      help="decide each match one at a time (with smart suggestions)")
    ap.add_argument("--replace", default=", ", metavar="TEXT",
                    help="replacement text for --remove-all (default: ', ')")
    ap.add_argument("--dry", action="store_true",
                    help="show what would change but don't write any file")
    ap.add_argument("-o", "--output", metavar="PATH",
                    help="where to write the cleaned copy (default: <file>.nodash.<ext>)")
    ap.add_argument("-a", "--all-dashes", action="store_true",
                    help="also match en dash, horizontal bar, minus sign, and '--'")
    ap.add_argument("--no-color", action="store_true", help="plain >>> <<< markers")
    if len(sys.argv) == 1:  # run with no arguments: show the full help, not an error
        ap.print_help()
        sys.exit(0)
    args = ap.parse_args()

    if not os.path.isfile(args.file):
        sys.exit(f"File not found: {args.file}")

    doc = load_doc(args.file)
    editing = args.smart or args.remove_all or args.interactive
    if editing and not doc.editable:
        sys.exit("PDFs are scan-only (no clean way to edit text inside a PDF).\n"
                 "Convert to .docx (Word: File > Open the PDF, then save as .docx) and re-run.")

    targets = dict(DASHES) if args.all_dashes else {EM_DASH: "EM DASH"}
    pattern = re.compile("|".join(re.escape(t) for t in sorted(targets, key=len, reverse=True)))

    use_color = not args.no_color and sys.stdout.isatty()
    if use_color:
        enable_ansi()
    hi = "\033[1;93;41m" if use_color else ">>>"
    reset = "\033[0m" if use_color else "<<<"

    ltext, lmap, recs = find_matches(doc, pattern)
    print(f"\nScanning: {args.file}\n" + "-" * 60)

    if not editing:
        scan(doc, recs, targets, hi, reset)
        print()
        return

    if not recs:
        print("No em dashes found — nothing to do, no file written.")
        return

    if args.remove_all:
        decisions = [(args.replace, False)] * len(recs)
    else:
        suggs = smart_decide(ltext, recs)
        if args.smart:
            smart_report(doc, recs, suggs, hi, reset)
            decisions = [(repl, cap) for repl, cap, _ in suggs]
        else:
            decisions = interactive(doc, recs, suggs, targets, hi, reset)
            if decisions is None:
                return

    edits = build_edits(doc, ltext, lmap, recs, decisions)
    changed = sum(1 for d in decisions if d is not None and d[0] is not None)

    print("-" * 60)
    print(f"{'Would replace' if args.dry else 'Replaced'} {changed} of {len(recs)} match(es)"
          f" ({len(recs) - changed} kept).")
    if args.dry:
        print("Dry run — no file written.\n")
        return
    if not edits:
        print("You kept everything — no changes, no file written.\n")
        return

    out = args.output or default_output(args.file)
    if os.path.abspath(out) == os.path.abspath(args.file):
        sys.exit("Refusing to overwrite the original file. Pick a different -o path.")
    if os.path.exists(out):
        sys.exit(f"Refusing to overwrite existing file: {out}. Pick a different -o path.")

    apply_edits(doc, edits)
    doc.save(out)
    print(f"Original untouched: {args.file}")
    print(f"Cleaned copy:       {out}\n")


if __name__ == "__main__":
    main()
