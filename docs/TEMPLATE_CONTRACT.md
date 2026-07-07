# Template Contract

What the engine expects from a project's template layer. A project ships two
things next to its draft:

1. **`template_prep.py`** — a script that turns the organisation's official
   DOCX template into a Jinja-instrumented template (written once per
   template, usually by an AI following the `jinjify-template` skill).
2. **The instrumented template** (`<template-stem>_jinja.docx`) — its
   deterministic output, regenerated automatically by the build when missing
   or older than the source template or the prep script.

The engine (`docx_builder`) never contains template-specific knowledge:
heading texts, style IDs, table layouts and TOC instructions all live in the
prep script. This document is the interface between the two.

---

## 1. Prep script interface

The build loads the prep script given by `--prep` (default:
`./template_prep.py`) as a Python module. It must export:

```python
def instrument_template(input_path: str, output_path: str) -> None:
    """Read the official template, write the Jinja-instrumented copy."""
```

and may export:

```python
RENDER_OPTIONS: dict   # keyword overrides for docx_builder.render.render()
```

Recognised `RENDER_OPTIONS` keys:

| Key | Default | Meaning |
| --- | --- | --- |
| `figure_title_style` | `FImageTitle` | Style ID of figure-title paragraphs. The engine's figure post-processing (bookmark insertion, run bolding, borderless-table wrapping) keys on it, and the LOF field should select it. |

`instrument_template` must be **deterministic**: same input template → byte
-identical output. The instrumented template is a build artifact, not a
source file.

Generic building blocks for prep scripts live in
`docx_builder.template_tools` (find/remove template sections, instrument
tables for row loops, insert Word TOC fields, insert the content loop).

## 2. Context variables

The draft supplies context through VAR blocks:

```markdown
<!-- VAR:authors TYPE:table COLUMNS:name,beneficiary,email -->
Jane Doe|ACME|jane@acme.org
<!-- ENDVAR -->

<!-- VAR:content TYPE:freeform -->
## Executive Summary
...body markdown...
<!-- ENDVAR -->
```

- `TYPE:table` → list of dicts keyed by `COLUMNS` (rows are pipe-delimited);
  render with a `{%tr for ... %}` row loop (`template_tools.instrument_table`).
- `TYPE:freeform` → the parsed document body (list of content nodes);
  one per fillable region (a fixed-structure template may have several);
  render with the content loop (§3).
- `TYPE:yaml` → the block parsed as YAML (a ```yaml fence around the payload
  is tolerated and keeps the markdown preview clean). For nested data that
  pipe-rows cannot express — e.g. a per-WP dossier with a deliverables list —
  rendered with block-table loops (§4a) and nested `{%tr for %}` loops.
  Booleans follow YAML 1.2: only `true`/`false` — `no:`/`yes:` stay strings,
  so they're safe as field names. Trailing newlines of folded/literal
  scalars are stripped (they would render as spurious breaks in cells).
- `TYPE:checklist` → GitHub task-list lines (`- [x] key` / `- [ ] key`,
  an optional dash/colon note after the key is ignored) parsed into a
  `key → bool` dict; drives native checkboxes (§4b) via `{{ name.key }}`.
- any other `TYPE` → plain string.

In any non-freeform value, a string that is exactly a markdown image —
`![alt](path)` with an optional `[50%]`/`[7cm]` width hint — is replaced by a
native inline image at render time (missing files degrade to an italic
`[image not found: …]` note).

The engine always adds `acronyms`: a list of `{abbrev, definition}` dicts for
the acronyms actually used, sorted case-insensitively.

**Validation is strict and bidirectional**: every Jinja variable in the
template must exist in the context, and every context variable must be used
by the template. A mismatch fails the build. The template therefore defines,
by construction, which VAR blocks a draft must provide.

## 3. The content loop

The document body renders through a union-type loop over `content` (or
whatever the freeform VAR is named). The instrumented template must contain,
in body position:

```
{%p for item in content %}
  ...one {%p if item.type == "<T>" %} ... {%p endif %} block per type...
{%p endfor %}
```

Types and the attributes each item exposes:

| `item.type` | Attributes | Notes |
| --- | --- | --- |
| `h1`…`h4` | `text` | Map to the template's heading styles. Word numbering comes from the styles, not the engine. |
| `p` | `text`, `has_math`, `math_placeholder` | Gate on `has_math`: plain paragraphs render `{{ item.text }}`; math paragraphs render `{{ item.math_placeholder }}` (replaced with OMML in post-processing). |
| `bullet` | `text`, `level` (1–4), `has_math`, `math_placeholder` | One block per level; attach the template's list style / `numId`+`ilvl` per level. |
| `numbered` | `text`, `has_math`, `math_placeholder` | Ordered list. |
| `figure` | `fig_start`, `image`, `title`, `caption`, `fig_end` | Five paragraphs in this order: sentinel, image (centered), title (in `figure_title_style`), caption, sentinel. Post-processing wraps the group into a borderless 2-row table. |
| `table` | `caption`, `subdoc` | Caption paragraph (use the style your LOT field selects) then `{{ item.subdoc }}`, which renders a `TABLE_PLACEHOLDER_N` sentinel replaced with a real Word table in post-processing. |
| `math_display` | `math_placeholder` | Centered paragraph; replaced with OMML via pandoc. |

### Why placeholders instead of direct content

docxtpl cannot render `RichText` or subdocuments inside `{%p if %}` blocks.
The engine therefore renders plain sentinel strings and rewrites them in an
OOXML post-processing pass. **Do not** put formatted content directly in the
loop; use the attributes above and let post-processing do its work.

### Sentinels reserved by the engine

The engine finds and replaces these in the rendered document — templates and
drafts must not produce them by other means:

`TABLE_PLACEHOLDER_N`, `MATH_DISP_N`, `MATH_PARA_N`, `MATH_CAP_N`,
`FIGURE_START_N` / `FIGURE_END_N`, `@@INLINEFMT:n@@`, `@@SECLABEL:id@@`,
`@@SECREF:id@@`, `@@FIGREF:id@@`, `@@TABREF:id@@`, `@@CHK:0@@` / `@@CHK:1@@`.

## 4. Front matter

Anything outside the content loop is free-form and template-specific;
typical prep-script work:

- delete "instructions for authors" sections from the official template;
- instrument document-control tables (authors, revisions, acronyms) with
  `{%tr for %}` row loops (`template_tools.instrument_table`);
- keep the template's own TOC field; insert List of Figures / List of
  Tables as Word-native `TOC` fields (`template_tools.insert_toc_sections`)
  whose `\t` switch selects the figure-title and table-caption styles;
- apply organisation-specific style corrections (alignment, bullets, fonts).

Fields are inserted dirty, so Word regenerates TOC/LOF/LOT on open.

### 4a. Replicated table blocks

Templates that say "table to be replicated for each WP" are handled by
`template_tools.wrap_block_loop`: a `{%p for wp in <var> %}` paragraph before
the block (caption paragraph + table) and `{%p endfor %}` after it. docxtpl
repeats everything in between per item and removes the marker paragraphs.
Inside the block use `{{ wp.field }}` in cells, `{{ loop.index }}` for the
"Table 2.X" numbering, and nested `{%tr for d in wp.deliverables %}` row
loops for variable-length sub-lists. Cells whose template text spans several
guidance paragraphs should be cleared with `set_cell_text_full`.

### 4b. Native checkboxes

`w14:checkbox` content controls stay live: the prep script binds each one to
a Jinja boolean expression with `template_tools.jinjify_checkboxes`, which
makes the control render a `@@CHK:1@@`/`@@CHK:0@@` sentinel; the engine's
checkbox post-pass then writes `w14:checked` and the control's own declared
state glyph — exactly what Word writes when a user ticks the box. Drive the
expressions from `TYPE:checklist` dicts (`checks.key` / `not checks.key` for
Yes/No pairs) or from enum-valued plain VARs (`report_type == 'interim'`).

## 5. Cross-references

- Headings labelled `{#id}` in the draft get a Word bookmark `sec_<id>`.
- `:ref{id}:` renders as `Section ` + a `REF \n \h` field → live section
  number.
- `:fig{id}:` / `:tab{id}:` references render as `REF \h` fields with static
  `Figure N` / `Table N` display text, hyperlinked to bookmarks placed on
  the figure-title / table-caption paragraphs.

The engine needs no template cooperation for this beyond
`figure_title_style` and heading styles with Word-native numbering.

## 6. Comments

Blockquotes are reserved in this dialect for author comments — nothing
else may use `>` syntax. The first line inside the blockquote must read
`comment: <author>`; every other paragraph inside becomes one paragraph of
the comment body:

```markdown
> comment: Nikos
>
> This attaches to the paragraph/heading immediately below.

This is the commented text.
```

The comment attaches to whatever content unit follows it — a heading,
plain paragraph, bullet, or numbered item (not figures, tables,
math-display blocks, or paragraphs containing inline math). It renders as
a real Word (OOXML) comment — a margin note attached to that text, with
`author` set from the blockquote header — not as visible body content.

Parsing is strict, by design: a blockquote that doesn't start with
`comment: <author>`, a comment with no body, two comments in a row with
nothing between them, a comment with nothing following it, or a comment
attached to an unsupported node type, all fail the build with a clear
error rather than silently degrading. This is deliberate — comments are
meta-content about the draft, not draft content, so a malformed one
should never quietly end up as (or vanish from) the shipped document.

That strictness only constrains *what type* of node a comment may attach
to — it says nothing about *where* in the document that's a good idea.
Attachment is forward-only and purely positional (whatever content node
comes next), so a comment worded as if it refers to preceding text (e.g.
"needs a citation for the claim above") will still mechanically attach to
whatever follows it — a real authoring footgun the parser has no way to
catch, since it can't read intent out of the comment body. To make this
checkable rather than just documented, every build reports each
`(target, comment)` pairing — `build_report.json`'s
`metrics.comment_assignments`, a list of `{author, target_type,
target_preview, comment_preview}` — so a build reviewer can eyeball
whether each comment landed where its text implies it should, without
opening the rendered DOCX.

The engine needs no template cooperation for this feature.

## 7. The markdown dialect (source side, for reference)

`:KEY:` acronyms (expanded on first use), `:fig{id}:` / `:tab{id}:` /
`:ref{id}:` cross-references, `{#label}` heading anchors,
`**:fig{id}: Title** [width]` + image/mermaid + `*:fig{id}: caption*` figure
groups, `**:tab{id}: Caption**` + GFM table groups, `$...$` / `$$...$$`
math, ` ```mermaid ` fences, `> comment: Author` blockquotes (§6). All
degrade gracefully in a plain markdown preview — that property is a design
goal; keep it when extending the dialect.
