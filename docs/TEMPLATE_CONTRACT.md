# Template Contract

What the engine expects from a project's template layer. A project ships two
things next to its draft:

1. **`template_prep.py`** — a script that turns the organisation's official
   DOCX template into a Jinja-instrumented template (written once per
   template, usually by an AI following the `jinjify-template` skill).
2. **The instrumented template** (`template_jinja.docx`) — its deterministic
   output, regenerated automatically by the build when missing.

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
  exactly one per draft; render with the content loop (§3).
- any other `TYPE` → plain string.

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
`@@SECREF:id@@`, `@@FIGREF:id@@`, `@@TABREF:id@@`.

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

## 5. Cross-references

- Headings labelled `{#id}` in the draft get a Word bookmark `sec_<id>`.
- `:ref{id}:` renders as `Section ` + a `REF \n \h` field → live section
  number.
- `:fig{id}:` / `:tab{id}:` references render as `REF \h` fields with static
  `Figure N` / `Table N` display text, hyperlinked to bookmarks placed on
  the figure-title / table-caption paragraphs.

The engine needs no template cooperation for this beyond
`figure_title_style` and heading styles with Word-native numbering.

## 6. The markdown dialect (source side, for reference)

`:KEY:` acronyms (expanded on first use), `:fig{id}:` / `:tab{id}:` /
`:ref{id}:` cross-references, `{#label}` heading anchors,
`**:fig{id}: Title** [width]` + image/mermaid + `*:fig{id}: caption*` figure
groups, `**:tab{id}: Caption**` + GFM table groups, `$...$` / `$$...$$`
math, ` ```mermaid ` fences. All degrade gracefully in a plain markdown
preview — that property is a design goal; keep it when extending the
dialect.
