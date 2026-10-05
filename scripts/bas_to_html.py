#!/usr/bin/env python3
"""Convert a QB64 program (+ optional captured output) into Word-friendly HTML.

Why the HTML looks "old-fashioned": Microsoft Word's HTML import ignores most
modern CSS (flexbox, grid, CSS variables, external/class styles are unreliable).
So the output uses:
  * tables for layout and centering (align="center" + margin auto),
  * inline style="" attributes only,
  * pt units, explicit margins/line-height, 'Courier New' for code,
  * bgcolor attributes in addition to CSS backgrounds.
Open the HTML in a browser, Ctrl+A, Ctrl+C, paste into Word (Keep Source Formatting),
or open the .html file directly in Word.

Usage:
  python scripts/bas_to_html.py samples/grade_calculator.bas \
      --output-text samples/grade_calculator_output.txt -o output/grade_calculator.html
"""
from __future__ import annotations

import argparse
import html
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from qb64_parse import KEYWORDS, parse_steps, program_title  # noqa: E402

BODY_FONT = "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222222;"
CODE_FONT = "font-family:'Courier New',Courier,monospace;font-size:10pt;"
P = f"margin:0 0 6pt 0;line-height:14pt;mso-line-height-rule:exactly;{BODY_FONT}"
H1 = "margin:0 0 12pt 0;font-family:Calibri,Arial,sans-serif;font-size:20pt;color:#1F3864;"
H2 = "margin:18pt 0 6pt 0;font-family:Calibri,Arial,sans-serif;font-size:14pt;color:#1F3864;"
H3 = "margin:12pt 0 4pt 0;font-family:Calibri,Arial,sans-serif;font-size:12pt;color:#2F5496;"

TOKEN_RE = re.compile(r'("[^"]*")|(\b[A-Za-z_][A-Za-z0-9_]*\b)')


def highlight(line: str) -> str:
    """Escape and colour one line of QB64 code with inline spans."""
    stripped = line.lstrip()
    if stripped.startswith("'") or stripped.upper().startswith("REM "):
        return f'<span style="color:#008000;font-style:italic;">{html.escape(line)}</span>'
    out, pos = [], 0
    for m in TOKEN_RE.finditer(line):
        out.append(html.escape(line[pos:m.start()]))
        token = m.group(0)
        if m.group(1):
            out.append(f'<span style="color:#A31515;">{html.escape(token)}</span>')
        elif token.upper() in KEYWORDS:
            out.append(f'<span style="color:#0000CC;font-weight:bold;">{html.escape(token)}</span>')
        else:
            out.append(html.escape(token))
        pos = m.end()
    out.append(html.escape(line[pos:]))
    return "".join(out)


TAG_SPLIT_RE = re.compile(r"(<[^>]+>)")


def preserve_spaces(fragment: str) -> str:
    """Turn leading spaces and runs of 2+ spaces into &nbsp; (text only, never inside tags).

    HTML (and Word's importer) collapses whitespace, which would destroy code
    indentation and column alignment such as PRINT USING output.
    """
    out = []
    for part in TAG_SPLIT_RE.split(fragment):
        if part.startswith("<"):
            out.append(part)
        else:
            out.append(re.sub(r"^ | {2,}", lambda m: "&nbsp;" * len(m.group(0)), part))
    return "".join(out)


def code_block(code: str, numbered: bool = False, highlight_syntax: bool = True) -> str:
    """Monospace code block as a one-cell shaded table (Word keeps shading on tables).

    Each line is its own <p> with exact line-height and zero margins, so Word does
    not add "space after" between code lines. Leading spaces become &nbsp; so the
    indentation survives the paste.
    """
    lines = code.splitlines() or [""]
    width = len(str(len(lines)))
    rows = []
    for n, line in enumerate(lines, start=1):
        body = preserve_spaces(highlight(line) if highlight_syntax else html.escape(line))
        prefix = ""
        if numbered:
            prefix = (f'<span style="color:#888888;">{str(n).rjust(width).replace(" ", "&nbsp;")}'
                      f'&nbsp;&nbsp;</span>')
        rows.append(
            f'<p style="margin:0;line-height:12pt;mso-line-height-rule:exactly;{CODE_FONT}">'
            f'{prefix}{body or "&nbsp;"}</p>'
        )
    inner = "\n".join(rows)
    return (
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="width:100%;border-collapse:collapse;margin:0 0 10pt 0;">'
        '<tr><td bgcolor="#F4F4F4" style="background:#F4F4F4;border:1pt solid #CCCCCC;'
        f'padding:6pt 8pt;">\n{inner}\n</td></tr></table>'
    )


def step_diagram(steps) -> str:
    """Centered vertical flow diagram built only from nested tables.

    Outer table: align="center" (Word) + margin auto (browsers). Each step is a
    bordered cell; arrows are their own centered rows. No flexbox, no grid.
    """
    rows = []
    for idx, step in enumerate(steps):
        rows.append(
            '<tr><td align="center" style="padding:0;">'
            '<table align="center" width="320" cellpadding="0" cellspacing="0" border="0" '
            'style="width:240pt;border-collapse:collapse;margin:0 auto;">'
            '<tr><td align="center" valign="middle" bgcolor="#DEEAF6" '
            'style="background:#DEEAF6;border:1.5pt solid #2F5496;padding:6pt 8pt;">'
            f'<p align="center" style="margin:0;text-align:center;line-height:14pt;'
            f'mso-line-height-rule:exactly;{BODY_FONT}"><b>Step {step.number}</b></p>'
            f'<p align="center" style="margin:0;text-align:center;line-height:14pt;'
            f'mso-line-height-rule:exactly;{BODY_FONT}">{html.escape(step.title)}</p>'
            '</td></tr></table></td></tr>'
        )
        if idx < len(steps) - 1:
            rows.append(
                '<tr><td align="center" style="padding:0;">'
                '<p align="center" style="margin:2pt 0;text-align:center;line-height:16pt;'
                'mso-line-height-rule:exactly;font-family:Arial,sans-serif;font-size:14pt;'
                'color:#2F5496;">&#9660;</p></td></tr>'
            )
    return (
        '<table align="center" cellpadding="0" cellspacing="0" border="0" '
        'style="margin:0 auto 12pt auto;border-collapse:collapse;">\n'
        + "\n".join(rows)
        + "\n</table>"
    )


def build_html(source: str, output_text: str | None = None, title: str | None = None) -> str:
    steps = parse_steps(source)
    title = title or program_title(source)
    parts = [
        f'<h1 style="{H1}">{html.escape(title)}</h1>',
        f'<p style="{P}">This document walks through a QB64 (QuickBASIC-compatible) program '
        f'step by step. The diagram shows the order in which the program runs; each step '
        f'is then explained with its code.</p>',
    ]
    if steps:
        parts.append(f'<h2 style="{H2}">Program flow</h2>')
        parts.append(step_diagram(steps))
        parts.append(f'<h2 style="{H2}">Step-by-step code</h2>')
        for step in steps:
            parts.append(f'<h3 style="{H3}">Step {step.number}: {html.escape(step.title)}</h3>')
            parts.append(code_block(step.code))
    parts.append(f'<h2 style="{H2}">Full listing</h2>')
    parts.append(code_block(source.rstrip("\n"), numbered=True))
    if output_text is not None:
        parts.append(f'<h2 style="{H2}">Program output</h2>')
        parts.append(code_block(output_text.rstrip("\n"), highlight_syntax=False))
    body = "\n".join(parts)
    return (
        '<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n'
        f'<title>{html.escape(title)}</title>\n</head>\n'
        f'<body style="margin:36pt;{BODY_FONT}">\n{body}\n</body>\n</html>\n'
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("source", help="QB64 .bas file")
    ap.add_argument("--output-text", help="captured program output to include")
    ap.add_argument("--title", help="override document title")
    ap.add_argument("-o", "--out", default=None, help="output .html (default: stdout)")
    args = ap.parse_args(argv)

    with open(args.source, encoding="utf-8") as fh:
        source = fh.read()
    output_text = None
    if args.output_text:
        with open(args.output_text, encoding="utf-8") as fh:
            output_text = fh.read()
    doc = build_html(source, output_text, args.title)
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(doc)
        print(f"wrote {args.out} ({len(parse_steps(source))} steps)")
    else:
        sys.stdout.write(doc)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
