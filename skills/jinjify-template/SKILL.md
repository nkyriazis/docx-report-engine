---
name: jinjify-template
description: Adapt a new organisation DOCX report template to the report engine — write the project's template_prep.py so markdown drafts compile into it. Use when a project brings a new official .docx template (usually containing author instructions) that the engine must fill.
---

# Jinjify a DOCX template for the report engine

You are given an organisation's official report template (a `.docx`, usually
with written instructions for authors inside it) and must produce the
project's `template_prep.py` — the script that rewrites that template into a
Jinja-instrumented one satisfying the engine's template contract
(`engine/docs/TEMPLATE_CONTRACT.md`, read it first).

A complete worked example ships with the engine's reference project: the
SEAVISTA `template_prep.py`. Base your script on it; adapt, don't rewrite.

## Step 1 — Inventory the template

Unzip-inspect before writing any code:

```python
from docx import Document
doc = Document("documents/<template>.docx")
for i, p in enumerate(doc.paragraphs[:200]):
    if p.style.style_id != "Normal" or p.text.strip():
        print(i, [p.style.style_id], p.text[:80])
for i, t in enumerate(doc.tables):
    print("TABLE", i, len(t.rows), "x", len(t.columns),
          [c.text[:20] for c in t.rows[0].cells])
print([s.style_id for s in doc.styles if s.type == 1][:60])
```

Record:
- **Section map** — which headings delimit: author instructions (to delete),
  document-control front matter (to keep + instrument), TOC/LOF/LOT area,
  and where the body placeholder starts.
- **Front-matter tables** — index, header-row count, and column meaning of
  each (authors, revisions, acronyms, …).
- **Style inventory** — heading styles, body style, bullet/numbered list
  styles (and their `numId`s in `word/numbering.xml`), the figure-title
  style, the table-caption style, TOC styles.
- **Written constraints** — read the instruction text inside the template
  (font rules, mandatory sections, caption conventions, fixed structure).
  These constraints are requirements for your prep script and, sometimes,
  for the draft's structure; note the ones the prep script cannot enforce.

## Step 2 — Decide the mapping

Fill this table before coding (it becomes the header comment of your prep
script):

| Contract element | This template |
| --- | --- |
| headings h1–h4 | e.g. `Heading1`…`Heading4` |
| body paragraph | e.g. `Normal` |
| bullet levels 1–4 | style + `numId`/`ilvl` per level |
| numbered list | style + `numId` |
| figure title (`RENDER_OPTIONS.figure_title_style`) | e.g. `FImageTitle` |
| table caption | e.g. `Caption` |
| LOF/LOT `TOC \t` selectors | style *names* (not IDs), e.g. `"F.Image Title,1"` |
| front-matter tables | index → VAR name → columns |
| sections to delete | heading texts |
| body placeholder start | heading text |

Templates with a **fixed body structure** (pre-set chapters that must be
filled rather than replaced) still fit: instead of deleting the body and
inserting one content loop, insert one loop per fillable region, each bound
to its own freeform VAR (e.g. `{%p for item in chapter_intro %}`). The
contract's strict context validation then forces the draft to provide
exactly those VAR blocks.

## Step 3 — Write template_prep.py

Use the generic helpers — `docx_builder.template_tools` — for all mechanics:
`find_heading`, `remove_range`, `remove_styled_and_empty_between`,
`instrument_table`, `insert_toc_sections`, `insert_content_loop`,
`make_para`, `get_or_make_spacing`. Your script should mostly be *data*:
marker texts, table configs, the content-loop tuple list, style overrides.

Requirements:
- export `instrument_template(input_path, output_path)` and
  `RENDER_OPTIONS`;
- be deterministic (no timestamps, no randomness);
- fail loudly when a structural marker is missing (`RuntimeError` listing
  what was not found) — templates change between organisation versions;
- keep the content loop complete: every `item.type` from the contract needs
  a block, including the `has_math` variants, even if the first draft uses
  no math — drafts evolve;
- end with a printed structure report (see the worked example's `_report`)
  so a human can eyeball the instrumented body.

## Step 4 — Verify

1. `python -m docx_builder instrument-template --template <official>.docx
   --output template_jinja.docx --prep template_prep.py` — read the report,
   confirm section removal and loop placement.
2. Run it twice; byte-compare the outputs (determinism).
3. Build a kitchen-sink draft that exercises every feature: h1–h4, bold /
   italic / code / colored spans, bullets to level 4, numbered lists,
   a GFM table with `:tab{}:` caption, a figure with `[50%]` width hint,
   a mermaid figure, inline + display math, `:ref{}:`/`:fig{}:`/`:tab{}:`
   cross-references, acronyms. Build and open the DOCX; check every
   feature landed with the template's styles.
4. Check the build's own validation passed: no leftover sentinels, no
   unresolved references, cross-ref field count matches draft refs (the
   build fails or warns on these).
5. In Word (or LibreOffice), select-all → F9 to refresh fields: TOC, LOF,
   LOT and section numbers must populate; figure/table links must jump.

## Pitfalls seen in practice

- **docxtpl cannot render RichText/subdocs inside `{%p if %}`** — never put
  formatted content in the loop directly; the contract's placeholder
  attributes exist for this reason.
- **`\t` TOC switches take style *names*** (`"F.Image Title,1"`), while
  everything else uses style *IDs* (`FImageTitle`). Confusing them yields
  empty LOF/LOT.
- **List numbering lives in `word/numbering.xml`**, not the style: nested
  levels may need explicit `numId`/`ilvl` on the paragraph (see the worked
  example's level-2+ bullets) and sometimes numbering-definition surgery
  (bullet glyph fonts, indents).
- **Headings that precede the body** (e.g. "List of Figures") can consume
  the chapter numbering counter; suppress numbering on them explicitly.
- **`sectPr` must stay the last body child** — insert the content loop
  *before* it, never after.
