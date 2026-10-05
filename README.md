# QB64 → Word-friendly HTML / DOCX formatting demo

Turn a QuickBASIC / QB64 program into a document that pastes cleanly into
Microsoft Word: a **centered step-by-step flow diagram**, consistent paragraph
spacing, and **monospace, syntax-coloured code blocks** that keep their
indentation.

```
samples/grade_calculator.bas        QB64 program; each "' STEP n: ..." comment = one diagram box
samples/grade_calculator_output.txt captured program output
scripts/qb64_parse.py               splits the source into STEP sections
scripts/bas_to_html.py              -> Word-friendly HTML (tables + inline styles only)
scripts/bas_to_docx.py              -> native .docx via python-docx (optional)
output/                             generated examples (.html and .docx)
tests/                              unittest suite
```

## Quick start

```bash
pip install -r requirements.txt          # only needed for the .docx script
python scripts/bas_to_html.py samples/grade_calculator.bas \
    --output-text samples/grade_calculator_output.txt -o output/grade_calculator.html
python scripts/bas_to_docx.py samples/grade_calculator.bas \
    --output-text samples/grade_calculator_output.txt -o output/grade_calculator.docx
python -m unittest discover -s tests -v
```

Open `output/grade_calculator.html` directly in Word (File → Open), or open it
in a browser, select all, copy, and paste into Word with **Keep Source
Formatting**. Or just open the generated `.docx`.

## Why the HTML looks "old-school"

Word's HTML importer is closer to a 2003-era browser than to a modern one:

| Works in a browser | What Word does | What this demo uses instead |
|---|---|---|
| `display:flex` / `grid` for centering | ignored – boxes stack left-aligned | nested `<table align="center">` (+ `margin:0 auto` for browsers) |
| `<style>` blocks, classes, CSS variables | partly or wholly dropped on paste | inline `style=""` on every element |
| `<pre>` / collapsed whitespace | indentation and column alignment lost | one `<p>` per code line, leading/double spaces → `&nbsp;` |
| default `<p>` margins | Word adds its own "space after" (often 8–10 pt) | explicit `margin:0 0 6pt 0` and `line-height` in pt with `mso-line-height-rule:exactly` |
| `background` on `<div>` | often lost | shading on a table cell (`bgcolor` **and** `background`) |
| px / rem units | inconsistent conversion | `pt` everywhere (Word thinks in points) |

## Word spacing and centering tips

- **Centering a diagram:** center the *table* (Table Properties → Alignment:
  Center), not the text inside it. In HTML that is `align="center"` on the
  `<table>`; in python-docx it is `table.alignment = WD_TABLE_ALIGNMENT.CENTER`.
  Center the paragraphs inside each box separately.
- **Fixed box width:** give the step boxes an explicit width (here 240 pt) so
  every box is the same size regardless of its text length.
- **Arrows between boxes:** put them in their own centered table rows; a
  Unicode ▼ (U+25BC) survives copy/paste everywhere, unlike images or CSS
  borders-as-arrows.
- **No gaps between code lines:** set *Spacing Before/After = 0* and *Line
  spacing = Exactly 12 pt* for code paragraphs. If pasted code looks
  double-spaced, that's Word's default "space after" coming back.
- **Consistent spacing elsewhere:** define spacing once (Normal style: 6 pt
  after) instead of pressing Enter for blank lines. Toggle ¶ (Ctrl+Shift+8) to
  see stray empty paragraphs.
- **Monospace font:** use Courier New or Consolas – both ship with Word;
  exotic fonts silently fall back to Times New Roman on another machine.
- **Keep a step together:** select a heading + its code block → Paragraph →
  Line and Page Breaks → *Keep with next*, so a step never splits across pages.

## How I'd help in 1:1 sessions

I'm Mihai, a software engineer who tutors programming and practical tooling.
For a "make my code look right in a Word report" problem I'd keep it hands-on:

1. **Session 1 – the program (45–60 min).** Walk through the QB64 program
   together, mark its logical steps with `' STEP n:` comments, and run it
   (QB64 or QB64-PE) to capture the output.
2. **Session 2 – why Word breaks layouts (45 min).** Paste a flexbox page into
   Word live, see it collapse, then rebuild it with tables + inline styles.
   You leave knowing the handful of rules in the table above.
3. **Session 3 – automate it (60 min).** Read and extend the Python converter:
   add your own program, change colours/fonts, or generate a `.docx` directly
   with python-docx. Write a test for one thing you changed.
4. **Follow-up.** Async review of your final document/report and a checklist
   you can reuse for future assignments.

Every session ends with a small exercise and a working file you keep.

## License

MIT – see [LICENSE](LICENSE).
