#!/usr/bin/env python3
"""
One-off importer that turns an existing, inconsistently-styled source .docx
into structured JSON elements consumable by scripts/build-book.js, using
heuristics (paragraph style, italic runs, scripture-citation shape) rather
than a fixed input format.

Written for "Blessed by Association" (The Association Series, Book One);
reuse as a starting point for converting other legacy manuscripts, not as a
generic tool — the "Take these steps" / "A prayer" section markers and the
title-page/copyright-page layout below are specific to this manuscript's
front matter.

Usage: python3 scripts/convert-source-manuscript.py <source.docx> <output.json>
"""
import json
import re
import sys

import docx

CITE_RE = re.compile(r'^\W{0,3}\d?\s?[A-Za-z][A-Za-z ]{2,25}\d+:\d+(-\d+)?,?\s*KJV\.?\W{0,3}$')
QUOTE_START_RE = re.compile(r'^[“‘"\']')


def split_lines(text):
    return [p.strip() for p in text.split("\n") if p.strip()]


def load_tokens(source_path):
    """Flatten every paragraph into logical lines, splitting on embedded
    line breaks (w:br) so that e.g. a heading paragraph containing both the
    chapter label and its title on separate lines becomes two tokens."""
    d = docx.Document(source_path)
    tokens = []
    for p in d.paragraphs:
        is_heading2 = p.style.name == "Heading 2" if p.style else False
        for j, line in enumerate(split_lines(p.text)):
            tokens.append({"text": line, "is_heading2": is_heading2, "heading_line_index": j})
    return tokens


def classify(tokens):
    elements = []
    i = 0
    n = len(tokens)

    # ---- title page (series label, title, tagline, author, publisher) ----
    elements.append({
        "type": "title-page",
        "series": tokens[0]["text"], "title": tokens[1]["text"],
        "tagline": tokens[2]["text"], "author": tokens[3]["text"], "publisher": tokens[4]["text"],
    })
    i = 5

    # ---- copyright page: everything up to the "Contents" heading ----
    copyright_lines = []
    while tokens[i]["text"] != "Contents":
        copyright_lines.append(tokens[i]["text"])
        i += 1
    elements.append({"type": "copyright-page", "lines": copyright_lines})

    # ---- contents list: skipped for rendering (a real self-updating Word
    # TOC field is inserted instead) — just skip past it to the first chapter ----
    i += 1  # past "Contents"
    while not tokens[i]["is_heading2"]:
        i += 1
    elements.append({"type": "toc"})

    # ---- chapters ----
    while i < n:
        tok = tokens[i]
        assert tok["is_heading2"], f"expected a Heading 2 chapter marker at token {i}: {tok}"
        label = tok["text"]
        i += 1

        # Anomaly: some source chapter headings carry the title as a second
        # line of the *same* source paragraph (still tagged is_heading2).
        title = tokens[i]["text"]
        i += 1

        chapter = {
            "type": "chapter", "label": label, "title": title,
            "epigraph": None, "citation": None, "tagline": None, "body": [],
        }

        # Epigraph: either a scripture quote + citation line, or an
        # aphorism + a descriptive tagline line (no citation).
        line1 = tokens[i]["text"]; i += 1
        line2 = tokens[i]["text"] if i < n else None
        if line2 and CITE_RE.match(line2):
            chapter["epigraph"], chapter["citation"] = line1, line2
        else:
            chapter["epigraph"], chapter["tagline"] = line1, line2
        i += 1

        # Body: everything until the next chapter heading or end of document.
        while i < n and not (tokens[i]["is_heading2"] and tokens[i]["heading_line_index"] == 0):
            t = tokens[i]["text"]

            if t == "✶":
                chapter["body"].append({"type": "section-break"})
                i += 1
                continue

            if t.lower() == "take these steps":
                i += 1
                items = []
                while i < n and not tokens[i]["is_heading2"] and tokens[i]["text"].lower() != "a prayer":
                    items.append(tokens[i]["text"])
                    i += 1
                chapter["body"].append({"type": "steps", "heading": "Take These Steps", "items": items})
                continue

            if t.lower() == "a prayer":
                i += 1
                prayer_lines = []
                while i < n and not (tokens[i]["is_heading2"] and tokens[i]["heading_line_index"] == 0) and tokens[i]["text"] != "✶":
                    prayer_lines.append(tokens[i]["text"])
                    i += 1
                chapter["body"].append({"type": "prayer", "heading": "A Prayer", "lines": prayer_lines})
                continue

            # Mid-body scripture block quote: a quote-opening line
            # immediately followed by a citation-shaped line.
            nxt = tokens[i + 1]["text"] if i + 1 < n else None
            if QUOTE_START_RE.match(t) and nxt and CITE_RE.match(nxt):
                chapter["body"].append({"type": "block-quote", "text": t, "citation": nxt})
                i += 2
                continue

            chapter["body"].append({"type": "paragraph", "text": t})
            i += 1

        elements.append(chapter)

    # ---- back matter ----
    # Everything after the last chapter's own content (About the Author,
    # Also in the Series, colophon) landed inside the final chapter's body
    # above, tagged generically — split it out here.
    last_chapter = elements[-1]
    new_body, back_matter = [], []
    mode = "body"
    for item in last_chapter["body"]:
        if item.get("type") == "paragraph" and item["text"] in ("About the Author", "Also in the Association Series"):
            mode = "back"
        (new_body if mode == "body" else back_matter).append(item)
    last_chapter["body"] = new_body
    elements.append({"type": "back-matter", "items": back_matter})

    # The final chapter's closing "A prayer" (personal salvation prayer) and
    # the book's closing "A Prayer of Blessing Over You" benediction both
    # landed in one merged prayer block (neither is a Heading 2 nor a "✶").
    # Split them back into two distinct blocks.
    HEADING_MARKER = "A Prayer of Blessing Over You"
    split_body = []
    for item in last_chapter["body"]:
        if item.get("type") == "prayer" and HEADING_MARKER in item["lines"]:
            idx = item["lines"].index(HEADING_MARKER)
            split_body.append({"type": "prayer", "heading": item["heading"], "lines": item["lines"][:idx]})
            split_body.append({"type": "prayer", "heading": HEADING_MARKER, "lines": item["lines"][idx + 1:]})
        else:
            split_body.append(item)
    last_chapter["body"] = split_body

    return elements


def report(elements):
    chapters = [e for e in elements if e["type"] == "chapter"]
    print("chapters:", len(chapters))
    for c in chapters:
        counts = {t: sum(1 for b in c["body"] if b["type"] == t) for t in ("paragraph", "block-quote", "steps", "prayer")}
        print(f"  {c['label']!r:22} title={c['title']!r:45} has_citation={bool(c['citation'])} {counts}")
    print("back-matter items:", len(elements[-1]["items"]))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 scripts/convert-source-manuscript.py <source.docx> <output.json>")
        sys.exit(1)
    source_path, output_path = sys.argv[1], sys.argv[2]
    tokens = load_tokens(source_path)
    elements = classify(tokens)
    report(elements)
    json.dump(elements, open(output_path, "w"), indent=1)
    print(f"\nWrote {output_path}")
