#!/usr/bin/env python3
"""
Build the user manual's generated parts (help/) and check the whole manual.

    python tools/make_help.py            build, then check (exit 1 on any problem)
    python tools/make_help.py --check    check only, write nothing

What it writes - every page stays hand-written; only the parts between markers are made here:
  - each page's left menu                    <!-- nav --> ... <!-- /nav -->
  - the home page's list of every page       <!-- pages --> ... <!-- /pages -->
  - the word list's columns, from src/ColumnHelp.cpp (the sheet's own tooltips), so the
    list never goes stale                     <!-- columns --> ... <!-- /columns -->
  - every picture:  <figure class="shot" data-img="x.png" data-kind="..." data-need="...">
    gets the picture when help/images/x.png exists, else a marked placeholder box
                                              <!-- img --> ... <!-- /img -->
  - help/search.js        the search index (page titles, headings, text, keywords)
  - help/print.html       the whole manual on one page, to print or save as PDF
  - help/images/NEEDED.txt the pictures still missing, for the owner to make

What it checks: the page set (the macros' "?" buttons open pages by name), each page's
sections, every link / anchor / picture, well-formed HTML, no internet links, no local
paths or user names, and it lists the <!-- CHECK: ... --> notes still open.
No libraries beyond Python's own.
"""
import html
import html.parser
import os
import re
import struct
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HELP = ROOT / "help"
IMAGES = HELP / "images"

# The menu: (heading, [(page, title in the menu)]). Every page of the manual is here.
NAV = [
    ("Start here", [("index", "Home and search"), ("start", "Getting started"), ("install", "Install the add-in")]),
    ("In Mastercam", [("dump", "Dump to Excel"), ("load", "Load the changes"), ("undo-load", "Undo a load")]),
    ("The workbook", [("workbook", "What is in the workbook"), ("summary", "Summary page"),
                      ("sheet", "The main sheet"), ("tools-page", "Tools page")]),
    ("Ribbon: Find", [("find", "Find an op"), ("slowest", "Slowest ops")]),
    ("Ribbon: Edit cells", [("set", "Set"), ("scale", "Scale %"), ("copy", "Copy from op")]),
    ("Ribbon: Selected ops", [("text", "Comments & text"), ("coolant", "Coolant"),
                              ("calculator", "Speed & feed"), ("inspection", "Inspection & inserts")]),
    ("Ribbon: Improve", [("target-time", "Hit a target time"), ("check", "Program check"), ("scenarios", "Scenarios")]),
    ("Ribbon: Undo and review", [("undo", "Undo"), ("changes", "List changes"), ("report", "Change report")]),
    ("How it works", [("estimate", "The cycle time estimate"), ("flips", "How insert flips are counted")]),
    ("Help", [("words", "Word list"), ("troubleshooting", "Troubleshooting"), ("print", "Print the whole manual")]),
]

# The pages the macros and the add-in open by name - these file names must not change.
REQUIRED = ["index", "start", "workbook", "summary", "sheet", "tools-page", "dump", "load", "undo-load", "find",
            "set", "scale", "copy", "text", "coolant", "calculator", "inspection", "target-time", "check",
            "slowest", "scenarios", "undo", "changes", "report", "words", "troubleshooting", "install",
            "flips", "estimate"]

# Every topic page has these sections (by id). The home page and the print page do not.
SECTIONS = ["what", "when", "steps", "example", "related"]

GENERATED = {"print"}           # written here, never by hand

# Columns whose word-list entry points on to a general word (an anchor in words.html).
SEE_ALSO = {
    "feed": "w-per-rev", "feed_mode": "w-per-rev", "plunge_mode": "w-per-rev", "retract_mode": "w-per-rev",
    "speed": "w-sfm", "speed_mode": "w-css", "max_ss": "w-max-rpm", "step": "w-depth-of-cut",
    "stepover": "w-stepover", "stock_x": "w-stock-to-leave", "stock_z": "w-stock-to-leave",
    "insp_time": "w-insp-time", "flips": "w-flip", "flips_est": "w-flip", "flips_part": "w-flip",
    "edge_limit": "w-edge-life", "mrr": "w-mrr", "mrr_avg": "w-mrr", "mrr_engaged": "w-mrr",
    "air_pct": "w-air-cutting", "needs_regen": "w-regenerate", "est_cycle_time": "w-estimate",
    "coolant": "w-coolant", "coolant_before": "w-coolant", "pt_rough_speed": "w-primeturning",
    "cut_dia": "w-sfm", "removed": "w-mrr",
}
# Their place on the sheet, for the word list's headings (from the section comments in ColumnHelp.cpp).

PAGE_PATHS = lambda: sorted(p for p in HELP.glob("*.html"))


# ============================================================ small helpers

def read(p):
    return p.read_text(encoding="utf-8")


def write_if_changed(p, text):
    old = p.read_text(encoding="utf-8") if p.exists() else None
    if old != text:
        p.write_text(text, encoding="utf-8", newline="\n")
        return True
    return False


def between(text, tag, new, page):
    """Replace what is between <!-- tag --> and <!-- /tag -->."""
    pat = re.compile(r"(<!-- " + re.escape(tag) + r" -->)(.*?)(<!-- /" + re.escape(tag) + r" -->)", re.S)
    if not pat.search(text):
        raise SystemExit(f"{page}: no <!-- {tag} --> ... <!-- /{tag} --> markers")
    return pat.sub(lambda m: m.group(1) + new + m.group(3), text, count=1)


def png_size(p):
    with open(p, "rb") as f:
        head = f.read(24)
    if head[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    return struct.unpack(">II", head[16:24])


def title_of(page):
    for _, items in NAV:
        for name, t in items:
            if name == page:
                return t
    return page


# ============================================================ the menu

def nav_html(current):
    out = ['\n<nav class="side" aria-label="Manual pages">']
    for head, items in NAV:
        out.append(f"<h2>{html.escape(head)}</h2>\n<ul>")
        for name, t in items:
            cur = ' aria-current="page"' if name == current else ""
            out.append(f'<li><a href="{name}.html"{cur}>{html.escape(t)}</a></li>')
        out.append("</ul>")
    out.append("</nav>\n")
    return "\n".join(out)


def description_of(text):
    m = re.search(r'<meta name="description" content="([^"]*)"', text)
    return html.unescape(m.group(1)) if m else ""


def pages_html(texts):
    out = ['\n<div class="all-pages">']
    for head, items in NAV:
        out.append(f"<h3>{html.escape(head)}</h3>\n<ul>")
        for name, t in items:
            if name == "index":
                continue
            d = description_of(texts.get(name, ""))
            out.append(f'<li><a href="{name}.html">{html.escape(t)}</a>' + (f" <span>&ndash; {html.escape(d)}</span>" if d else "") + "</li>")
        out.append("</ul>")
    out.append("</div>\n")
    return "\n".join(out)


# ============================================================ the word list's columns

def column_help():
    """[(section, [(column, text)])] in ColumnHelp.cpp's own order."""
    src = read(ROOT / "src" / "ColumnHelp.cpp")
    body = src[src.index("help = {"):src.index("};", src.index("help = {"))]
    sections, cur = [], None
    for line in body.splitlines():
        s = line.strip()
        m = re.match(r"^//\s*(.+)$", s)
        if m:
            cur = (m.group(1).strip(), [])
            sections.append(cur)
            continue
        m = re.match(r'^\{\s*L"([^"]+)",\s*L"((?:[^"\\]|\\.)*)"\s*\},?$', s)
        if m:
            if cur is None:
                cur = ("Columns", [])
                sections.append(cur)
            text = m.group(2).replace('\\"', '"')
            cur[1].append((m.group(1), text))
    return sections


def columns_html():
    out = ['\n<p class="note">This part is made from the sheet\'s own tooltips (the text you see when you hover a '
           'column name), so it always matches your sheet. A column not listed here (most of the Tool inspection and '
           'Filter settings, the planes, home and reference points) has its own tooltip on the sheet: hover its name.</p>']
    for sect, cols in column_help():
        if not cols:
            continue
        sid = "c-" + re.sub(r"[^a-z0-9]+", "-", sect.lower()).strip("-")
        out.append(f'<h3 id="{sid}">{html.escape(sect)}</h3>\n<dl class="words">')
        for name, text in sorted(cols, key=lambda c: c[0].lower()):
            also = SEE_ALSO.get(name)
            out.append(f'<dt id="col-{name}"><span class="col">{html.escape(name)}</span></dt>')
            dd = html.escape(text)
            if also:
                dd += f' <span class="also">See also: <a href="#{also}">{html.escape(also[2:].replace("-", " "))}</a>.</span>'
            out.append(f"<dd>{dd}</dd>")
        out.append("</dl>")
    out.append("")
    return "\n".join(out)


# ============================================================ pictures

FIG = re.compile(r'(<figure class="shot[^"]*"((?:\s+[\w-]+="[^"]*")*)\s*>)(.*?)(</figure>)', re.S)


def attrs(s):
    return {k: html.unescape(v) for k, v in re.findall(r'([\w-]+)="([^"]*)"', s)}


def figures(text, page, missing):
    def one(m):
        a = attrs(m.group(2))
        img, need = a.get("data-img", ""), a.get("data-need", "")
        alt = a.get("data-alt", need)
        p = IMAGES / img
        rest = re.sub(r"<!-- img -->.*?<!-- /img -->", "<!-- img --><!-- /img -->", m.group(3), flags=re.S)
        if img and p.exists():
            size = png_size(p)
            wh = f' width="{size[0]}" height="{size[1]}"' if size else ""
            inner = (f'<a class="zoom" href="images/{img}" title="Open the picture full size">'
                     f'<img src="images/{img}" alt="{html.escape(alt)}"{wh} loading="lazy"></a>')
            cls = "shot"
        else:
            kind = a.get("data-kind", "")
            how = {"mastercam": "a screenshot from Mastercam - to be added",
                   "screen": "a screenshot - to be added",
                   "dialog": "made by tools/make_help_images.ps1 - the add-in's own window, drawn without Mastercam",
                   "ribbon": "made by tools/make_help_images.ps1",
                   "window": "made by tools/make_help_images.ps1 once this part of the macros is finished",
                   "sheet": "made by tools/make_help_images.ps1"}.get(kind, "to be added")
            inner = (f'<div class="todo-box" role="img" aria-label="Picture not ready yet: {html.escape(need)}">'
                     f'<strong>Picture coming: {html.escape(img)}</strong>{html.escape(need)} '
                     f'<em>({how})</em></div>')
            cls = "shot todo"
            missing.append((kind, img, page, need))
        head = re.sub(r'class="shot[^"]*"', f'class="{cls}"', m.group(1), count=1)
        rest = rest.replace("<!-- img --><!-- /img -->", "<!-- img -->" + inner + "<!-- /img -->", 1)
        return head + rest + m.group(4)
    return FIG.sub(one, text)


def needed_txt(missing):
    lines = ["Pictures the manual still needs (written by tools/make_help.py - do not edit by hand).",
             "Each page shows a dashed placeholder box where the picture goes; once the PNG is in",
             "help/images under this name, run  python tools/make_help.py  and it is put in.", ""]
    groups = [("mastercam", "MASTERCAM SCREENS - capture by hand (crop to the window; no other windows; number",
               "callouts 1, 2, 3 as the page's steps say; no part paths or user names in the picture)"),
              ("screen", "OTHER SCREENS - capture by hand (crop to the window; callouts as above)", ""),
              ("dialog", "THE ADD-IN'S OWN WINDOWS - run tools/make_help_images.ps1 (tests/dialog_shots.cpp draws them,",
               "no Mastercam needed)"),
              ("ribbon", "EXCEL'S RIBBON - run tools/make_help_images.ps1 (an Excel of its own, off the screen)", ""),
              ("window", "EXCEL MACRO WINDOWS AND SHEETS OF THE NEW COMMANDS - run tools/make_help_images.ps1 after",
               "they are merged (it captures only that window / range; check its shot list fits the new layout)"),
              ("sheet", "SHEET PICTURES - run tools/make_help_images.ps1", ""),
              ("", "OTHER", "")]
    seen = set()
    for kind, h1, h2 in groups:
        rows = sorted({(img, page, need) for k, img, page, need in missing if (k or "") == kind and img not in seen})
        if not rows:
            continue
        lines.append(h1)
        if h2:
            lines.append(h2)
        lines.append("")
        for img, page, need in rows:
            if img in seen:
                continue
            seen.add(img)
            pages = sorted({p for k, i, p, n in missing if i == img})
            lines.append(f"  images/{img}")
            lines.append(f"      on: {', '.join(p + '.html' for p in pages)}")
            lines.append(f"      show: {need}")
            lines.append("")
    if not seen:
        lines.append("None - every picture is in.")
    return "\n".join(lines) + "\n"


# ============================================================ text of a page (search, checks)

class Text(html.parser.HTMLParser):
    """The page's <main>, split into sections at each <h2>: [(id, heading, text)].
    Left out: scripts, menus, diagrams, picture placeholders and the "Home" link."""
    SKIP_TAGS = {"script", "style", "nav", "svg", "noscript"}
    SKIP_CLASSES = {"todo-box", "crumb"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_main = 0
        self.skip = 0                   # > 0 while inside a left-out element
        self.depth = 0                  # open (non-void) elements
        self.skip_at = []               # the depths where left-out elements began
        self.sections = [["", "", []]]
        self.heading = None
        self.h1 = ""
        self.in_h1 = False

    def handle_starttag(self, tag, a):
        a = dict(a)
        if tag == "main":
            self.in_main += 1
        if tag not in VOID:
            self.depth += 1
            if tag in self.SKIP_TAGS or set((a.get("class") or "").split()) & self.SKIP_CLASSES:
                self.skip_at.append(self.depth)
                self.skip = len(self.skip_at)
        if not self.in_main:
            return
        if tag == "section" and a.get("id"):
            self.section_id = a["id"]
        if tag == "h1":
            self.in_h1 = True
        if tag == "h2":
            # The heading's own id, else its section's: the link a search result goes to.
            self.sections.append([a.get("id") or getattr(self, "section_id", "") or "", "", []])
            self.heading = self.sections[-1]
        if tag in ("p", "li", "td", "th", "dt", "dd", "h2", "h3", "figcaption", "caption", "br"):
            self.sections[-1][2].append(" ")

    def handle_endtag(self, tag):
        if tag == "main":
            self.in_main -= 1
        if tag not in VOID:
            if self.skip_at and self.skip_at[-1] == self.depth:
                self.skip_at.pop()
                self.skip = len(self.skip_at)
            self.depth -= 1
        if tag == "h1":
            self.in_h1 = False
        if tag == "h2":
            self.heading = None

    def handle_data(self, d):
        if not self.in_main or self.skip:
            return
        if self.in_h1:
            self.h1 += d
            return
        if self.heading is not None:
            self.heading[1] += d            # the heading is searched as the heading, not twice
            return
        self.sections[-1][2].append(d)


def page_text(text):
    t = Text()
    t.feed(text)
    out = []
    for sid, head, parts in t.sections:
        body = re.sub(r"\s+", " ", "".join(parts)).strip()
        if body or head:
            out.append((sid, re.sub(r"\s+", " ", head).strip(), body))
    return re.sub(r"\s+", " ", t.h1).strip(), out


def js_str(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ").replace("</", "<\\/") + '"'


def search_js(texts):
    # The home page's "I want to ..." tasks are how people put it: each task's words become
    # keywords of the page it leads to. The home page itself is not a result.
    task_words = {}
    for m in re.finditer(r'<a href="([\w-]+)\.html"><strong>(.*?)</strong><span>(.*?)</span></a>', texts.get("index", "")):
        plain = html.unescape(re.sub(r"<[^>]+>", "", m.group(2) + " " + m.group(3)))
        task_words[m.group(1)] = task_words.get(m.group(1), "") + " " + plain
    rows = []
    for name in [n for _, items in NAV for n, _ in items]:
        if name in GENERATED or name not in texts or name == "index":
            continue
        text = texts[name]
        h1, sections = page_text(text)
        kw = ""
        m = re.search(r'<meta name="keywords" content="([^"]*)"', text)
        if m:
            kw = html.unescape(m.group(1))
        kw = (kw + task_words.get(name, "")).strip()
        title = h1 or title_of(name)
        for i, (sid, head, body) in enumerate(sections):
            if not body or (name == "words" and sid in ("general", "columns")):
                continue
            rows.append("{" + f'p:{js_str(name + ".html")},a:{js_str(sid)},t:{js_str(title)},h:{js_str(head)},'
                        f'k:{js_str(kw if i == 0 else "")},x:{js_str(body)}' + "}")
        if name == "words":
            # The word list: one result per word, so "what is max_ss" lands on max_ss.
            for m in re.finditer(r'<dt id="([^"]+)">(.*?)</dt>\s*<dd>(.*?)</dd>', text, re.S):
                anchor = m.group(1)
                inner = re.search(r'<span id="(w-[^"]+)">', m.group(2))
                if inner:
                    anchor = inner.group(1)
                term = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>|<!--.*?-->", "", m.group(2)))).strip()
                body = re.sub(r"\s+", " ", html.unescape(re.sub(r"<!--.*?-->|<[^>]+>", " ", m.group(3)))).strip()
                rows.append("{" + f'p:{js_str("words.html")},a:{js_str(anchor)},t:{js_str("Word list")},h:{js_str(term)},'
                            f'k:{js_str(term.replace("_", " ") + " " + term)},x:{js_str(body)}' + "}")
    return ("// The manual's search index - made by tools/make_help.py from the pages. Do not edit.\n"
            "window.PT_INDEX = [\n" + ",\n".join(rows) + "\n];\n")


# ============================================================ the print page

def print_html(texts):
    order = [n for _, items in NAV for n, _ in items if n not in GENERATED and n in texts]
    parts = []
    for name in order:
        text = texts[name]
        m = re.search(r"<main[^>]*>(.*)</main>", text, re.S)
        body = m.group(1)
        body = re.sub(r'<p class="crumb">.*?</p>', "", body, flags=re.S)
        body = re.sub(r'<div class="big-search[^"]*">.*?</div>\s*<ul id="results"[^>]*></ul>', "", body, flags=re.S)
        # Ids and links made unique to the one page: "#steps" on coolant.html -> "#coolant--steps".
        body = re.sub(r'\bid="([^"]+)"', lambda x: f'id="{name}--{x.group(1)}"', body)

        def link(x):
            href = x.group(1)
            if href.startswith(("http:", "https:", "mailto:", "images/")):
                return x.group(0)
            if href.startswith("#"):
                return f'href="#{name}--{href[1:]}"'
            mm = re.match(r"([\w-]+)\.html(?:#(.*))?$", href)
            if mm and mm.group(1) not in GENERATED:
                return f'href="#{mm.group(1)}--{mm.group(2)}"' if mm.group(2) else f'href="#page-{mm.group(1)}"'
            return x.group(0)
        body = re.sub(r'href="([^"]*)"', link, body)
        body = re.sub(r'(<svg[^>]*?)\bid="[^"]*"', r"\1", body)
        # One <h1> on the print page: each page's headings one level down.
        for lo, hi in (("h3", "h4"), ("h2", "h3"), ("h1", "h2")):
            body = re.sub(rf"<{lo}\b", f"<{hi}", body)
            body = body.replace(f"</{lo}>", f"</{hi}>")
        body = body.replace("<h2", '<h2 class="page-title"', 1)
        parts.append(f'<section class="print-page" id="page-{name}">\n{body.strip()}\n</section>')
    toc = "\n".join(f'<li><a href="#page-{n}">{html.escape(title_of(n))}</a></li>' for n in order)
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Print the whole manual - Parameter Table Tool</title>
<meta name="description" content="Every page of the manual on one page - print it, or save it as a PDF.">
<link rel="stylesheet" href="style.css">
<script src="search.js" defer></script>
<script src="help.js" defer></script>
</head>
<body>
<!-- Made by tools/make_help.py from every page - do not edit. -->
<a class="skip" href="#main">Skip to the page</a>
<header class="top">
<a class="home" href="index.html">Parameter Table Tool <span>User manual</span></a>
<form class="search-mini" role="search" action="index.html"><input type="search" name="q" placeholder="Search the manual" aria-label="Search the manual"><div class="drop" hidden><ul class="results"></ul></div></form>
<button class="menu-btn" type="button" aria-expanded="false">Menu</button>
</header>
<div class="wrap">
<!-- nav --><!-- /nav -->
<main id="main">
<p class="crumb no-print"><a href="index.html">&larr; Home and search</a></p>
<h1>Parameter Table Tool &ndash; the whole manual</h1>
<p class="lead no-print">Every page on one page. Press <kbd>Ctrl</kbd>+<kbd>P</kbd> to print it, or pick
<span class="ui">Microsoft Print to PDF</span> (or <span class="ui">Save as PDF</span>) as the printer to make a PDF.
Each topic starts on a new sheet of paper.</p>
<h2 id="contents">Contents</h2>
<ol>
{toc}
</ol>
{chr(10).join(parts)}
</main>
</div>
<footer>Parameter Table Tool &ndash; user manual. Works with no internet.</footer>
</body>
</html>
"""


# ============================================================ checks

VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr",
        "path", "rect", "circle", "line", "polyline", "polygon", "ellipse", "use", "stop"}


class Shape(html.parser.HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack, self.problems, self.ids, self.hrefs, self.srcs, self.h1 = [], [], [], [], [], 0
        self.imgs_no_alt = 0

    def handle_starttag(self, tag, a):
        a = dict(a)
        if a.get("id"):
            self.ids.append(a["id"])
        if tag == "a" and a.get("href") is not None:
            self.hrefs.append(a["href"])
        if tag in ("img", "script", "link") and (a.get("src") or a.get("href")):
            self.srcs.append(a.get("src") or a.get("href"))
        if tag == "img" and not a.get("alt"):
            self.imgs_no_alt += 1
        if tag == "h1":
            self.h1 += 1
        if tag not in VOID:
            self.stack.append((tag, self.getpos()[0]))

    def handle_startendtag(self, tag, a):
        self.handle_starttag(tag, a)
        if tag not in VOID and self.stack and self.stack[-1][0] == tag:
            self.stack.pop()

    def handle_endtag(self, tag):
        if tag in VOID:
            return
        if not self.stack:
            self.problems.append(f"line {self.getpos()[0]}: </{tag}> with nothing open")
            return
        if self.stack[-1][0] == tag:
            self.stack.pop()
            return
        # <p> and <li> close themselves; anything else out of order is a mistake.
        names = [t for t, _ in self.stack]
        if tag in names:
            while self.stack and self.stack[-1][0] != tag:
                t, line = self.stack.pop()
                if t not in ("p", "li", "dt", "dd", "td", "th", "tr", "option"):
                    self.problems.append(f"line {line}: <{t}> not closed before </{tag}>")
            self.stack.pop()
        else:
            self.problems.append(f"line {self.getpos()[0]}: </{tag}> does not match <{self.stack[-1][0]}>")


def check(texts, files_written):
    problems, notes = [], []
    names = set(texts)
    for r in REQUIRED:
        if r not in names:
            problems.append(f"missing page {r}.html (the macros open it by name)")
    nav_names = {n for _, items in NAV for n, _ in items}
    for n in names:
        if n not in nav_names:
            problems.append(f"{n}.html is not in the menu (NAV in tools/make_help.py)")

    # The pages the macros' "?" buttons and the ribbon ask for.
    asked = set()
    for f in list((ROOT / "vba").glob("*.bas")) + list((ROOT / "vba").glob("*.vb")) + list((ROOT / "vba").glob("*.cls")):
        for m in re.finditer(r'(?:ShowHelp\s*\(?\s*|HelpLink\s+[^,\n]+,\s*)"([\w-]*)"', f.read_text(encoding="latin-1")):
            if m.group(1):
                asked.add((m.group(1), f.name))
    for topic, where in sorted(asked):
        if topic not in names:
            problems.append(f'{where} opens help page "{topic}" - there is no {topic}.html')

    user = os.environ.get("USERNAME", "") or os.environ.get("USER", "")
    bad_words = [re.compile(r"[A-Za-z]:\\\\?Users\\\\", re.I), re.compile(r"Dropbox", re.I)]
    if user and len(user) > 2:
        bad_words.append(re.compile(r"\b" + re.escape(user) + r"\b", re.I))

    ids_of = {}
    for name, text in texts.items():
        s = Shape()
        s.feed(text)
        s.close()
        for p in s.problems:
            problems.append(f"{name}.html {p}")
        for t, line in s.stack:
            if t not in ("p", "li", "dt", "dd", "td", "th", "tr", "html", "body"):
                problems.append(f"{name}.html line {line}: <{t}> never closed")
        dup = {i for i in s.ids if s.ids.count(i) > 1}
        if dup:
            problems.append(f"{name}.html: the same id twice: {', '.join(sorted(dup))}")
        if s.h1 != 1:
            problems.append(f"{name}.html: {s.h1} <h1> headings (one wanted)")
        if s.imgs_no_alt:
            problems.append(f"{name}.html: {s.imgs_no_alt} picture(s) with no alt text")
        ids_of[name] = set(s.ids)
        s.name = name
        texts[name] = (text, s)
        if name not in ("index",) and name not in GENERATED:
            for sec in SECTIONS:
                if f'id="{sec}"' not in text:
                    problems.append(f'{name}.html: no section id="{sec}"')
        if "<title>" not in text or 'name="description"' not in text:
            problems.append(f"{name}.html: no <title> or description")
        for rx in bad_words:
            if rx.search(text):
                problems.append(f"{name}.html: holds a local path or user name ({rx.pattern})")
        for m in re.finditer(r"<!--\s*CHECK:(.*?)-->", text if name not in GENERATED else "", re.S):
            notes.append(f"{name}.html: {re.sub(r'\s+', ' ', m.group(1)).strip()}")

    for name, (text, s) in texts.items():
        for href in s.hrefs + s.srcs:
            if re.match(r"^(https?:|//|mailto:|javascript:)", href):
                if href.startswith(("http:", "https:", "//")):
                    problems.append(f"{name}.html: an internet link ({href}) - the manual works offline")
                continue
            path, _, frag = href.partition("#")
            path = path.split("?")[0]
            if path == "":
                target = name
            elif path.endswith(".html"):
                target = path[:-5]
                if target not in texts:
                    problems.append(f"{name}.html: link to {path}, which does not exist")
                    continue
            else:
                if not (HELP / path).exists():
                    problems.append(f"{name}.html: {href} does not exist")
                continue
            if frag and frag not in ids_of.get(target, set()):
                problems.append(f"{name}.html: link {href} - no id \"{frag}\" on {target}.html")

    # The text files: no paths or names either.
    for p in [HELP / "search.js", HELP / "help.js", HELP / "style.css"]:
        if p.exists():
            t = p.read_text(encoding="utf-8")
            for rx in bad_words:
                if rx.search(t):
                    problems.append(f"{p.name}: holds a local path or user name")
    return problems, notes


# ============================================================ main

def main():
    only_check = "--check" in sys.argv
    files = {p.stem: p for p in PAGE_PATHS()}
    texts = {n: read(p) for n, p in files.items() if n not in GENERATED}
    changed, missing = [], []

    for name in list(texts):
        t = texts[name]
        t = between(t, "nav", nav_html(name), name)
        t = figures(t, name, missing)
        if name == "words":
            t = between(t, "columns", columns_html(), name)
        texts[name] = t
    texts["index"] = between(texts["index"], "pages", pages_html(texts), "index")

    pr = print_html(texts)
    pr = between(pr, "nav", nav_html("print"), "print")
    texts["print"] = pr

    if not only_check:
        for name, t in texts.items():
            if write_if_changed(HELP / f"{name}.html", t):
                changed.append(f"{name}.html")
        if write_if_changed(HELP / "search.js", search_js(texts)):
            changed.append("search.js")
        IMAGES.mkdir(exist_ok=True)
        if write_if_changed(IMAGES / "NEEDED.txt", needed_txt(missing)):
            changed.append("images/NEEDED.txt")
    else:
        for name, t in texts.items():
            p = HELP / f"{name}.html"
            if not p.exists() or read(p) != t:
                changed.append(f"{name}.html (out of date - run without --check)")

    problems, notes = check(dict(texts), changed)
    imgs = sorted(p.name for p in IMAGES.glob("*.png")) if IMAGES.exists() else []
    print(f"pages: {len(texts)}   pictures in: {len(imgs)}   placeholders: {len({m[1] for m in missing})}")
    if changed:
        print(("out of date: " if only_check else "written: ") + ", ".join(changed))
    if notes:
        print(f"\n{len(notes)} CHECK note(s) for the owner to confirm:")
        for n in notes:
            print("  - " + n)
    if problems:
        print(f"\n{len(problems)} PROBLEM(S):")
        for p in problems:
            print("  - " + p)
        return 1
    if only_check and changed:
        return 1
    print("\nall links, anchors, pictures and pages check out")
    return 0


if __name__ == "__main__":
    sys.exit(main())
