#!/usr/bin/env python3
"""
One-off importer for "God's Ways" (Tina Suwa) into the same structured JSON
schema consumed by scripts/build-book.js. Written for this manuscript's
specific shape, which differs from "Blessed by Association":

  - No Heading-styled paragraphs at all — chapters are marked only by a
    plain "Chapter One" / "Chapter Two" / ... line.
  - No scripture epigraph under the chapter title; the title is followed
    directly by body prose.
  - Scripture quotes appear inline through the body as italic line(s)
    immediately followed by a BOLD citation (e.g. "Exodus 33:13", no book
    abbreviation suffix) — but some italic passages (personal prayers) have
    no citation at all.
  - In-chapter subheadings are bold lines that are *not* citations (i.e.
    not immediately preceded by an italic quote line).

Usage: python3 scripts/convert-source-manuscript-gods-ways.py <source.docx> <output.json>
"""
import json
import re
import sys

import docx

CITE_RE = re.compile(r'^\d?\s?[A-Za-z]+(?:\s[A-Za-z]+)?\s\d+:\d+(-\d+)?\.?$')
CHAPTER_RE = re.compile(
    r'^Chapter (One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten|Eleven|Twelve|Thirteen|Fourteen|Fifteen)$'
)


def load_lines(source_path):
    d = docx.Document(source_path)
    lines = []
    for p in d.paragraphs:
        if not p.text.strip():
            continue
        runs = p.runs
        italic = bool(runs) and all(r.italic for r in runs if r.text.strip())
        bold = bool(runs) and all(r.bold for r in runs if r.text.strip())
        lines.append({"text": p.text.strip(), "italic": italic, "bold": bold})
    return lines


def classify(lines):
    n = len(lines)
    i = 0
    elements = []

    # ---- title page: title, tagline, author, publisher ----
    elements.append({
        "type": "title-page",
        "title": lines[0]["text"], "tagline": lines[1]["text"],
        "author": lines[2]["text"], "publisher": lines[3]["text"],
    })
    i = 4

    # ---- copyright page: up to "Contents" ----
    copyright_lines = []
    while lines[i]["text"] != "Contents":
        copyright_lines.append(lines[i]["text"])
        i += 1
    elements.append({"type": "copyright-page", "lines": copyright_lines})

    # ---- contents: skip placeholder list, insert a real Word TOC field ----
    i += 1
    while not CHAPTER_RE.match(lines[i]["text"]):
        i += 1
    elements.append({"type": "toc"})

    # ---- chapters ----
    while i < n:
        assert CHAPTER_RE.match(lines[i]["text"]), f"expected a chapter marker at {i}: {lines[i]}"
        label = lines[i]["text"]
        i += 1
        title = lines[i]["text"]
        i += 1

        chapter = {"type": "chapter", "label": label, "title": title, "epigraph": None, "citation": None, "tagline": None, "body": []}

        while i < n and not CHAPTER_RE.match(lines[i]["text"]):
            l = lines[i]

            if l["italic"]:
                quote_lines = []
                while i < n and lines[i]["italic"]:
                    quote_lines.append(lines[i]["text"])
                    i += 1
                citation = None
                if i < n and lines[i]["bold"] and CITE_RE.match(lines[i]["text"]):
                    citation = lines[i]["text"]
                    i += 1
                chapter["body"].append({"type": "block-quote", "lines": quote_lines, "citation": citation})
                continue

            if l["bold"]:
                chapter["body"].append({"type": "subheading", "text": l["text"]})
                i += 1
                continue

            chapter["body"].append({"type": "paragraph", "text": l["text"]})
            i += 1

        elements.append(chapter)

    elements.append({"type": "back-matter", "items": []})
    return elements


def report(elements):
    chapters = [e for e in elements if e["type"] == "chapter"]
    print("chapters:", len(chapters))
    for c in chapters:
        counts = {t: sum(1 for b in c["body"] if b["type"] == t) for t in ("paragraph", "block-quote", "subheading")}
        uncited = sum(1 for b in c["body"] if b["type"] == "block-quote" and not b["citation"])
        print(f"  {c['label']!r:15} title={c['title']!r:40} {counts} uncited_quotes={uncited}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 scripts/convert-source-manuscript-gods-ways.py <source.docx> <output.json>")
        sys.exit(1)
    source_path, output_path = sys.argv[1], sys.argv[2]
    lines = load_lines(source_path)
    elements = classify(lines)
    report(elements)
    json.dump(elements, open(output_path, "w"), indent=1)
    print(f"\nWrote {output_path}")
