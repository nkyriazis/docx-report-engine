# docx-report-engine

Compile WYSIWYG-friendly Markdown drafts into organisation DOCX templates.

The engine implements a small Markdown dialect (acronym macros, figure /
table / section cross-references, math, mermaid) whose sources read cleanly
in any Markdown preview *and* compile into a fully styled Word document:
live REF-field cross-references, OMML math, Word-native TOC/LOF/LOT fields,
template styles throughout.

## Architecture

Three layers, only the first of which lives in this repository:

1. **Engine** (`docx_builder/`) — template-agnostic. Parses the dialect,
   renders through a Jinja-instrumented template (docxtpl), then rewrites
   the OOXML in post-processing (real tables, math, bookmarks and REF
   fields, figure wrapping).
2. **Template layer** (per project) — `template_prep.py`, a script that
   adapts the organisation's official template to the engine's contract.
   Written once per template, usually by an AI following
   `skills/jinjify-template/SKILL.md`. Contract:
   `docs/TEMPLATE_CONTRACT.md`.
3. **Project** (per report) — the draft (`draft/*.md`), `acronyms.yml`,
   images, and the official template.

## Usage

From a project directory (draft, template, `template_prep.py`, configs):

```bash
python -m docx_builder build \
  --draft draft/draft.md \
  --template documents/template.docx \
  --output report.docx \
  --compiled compiled.md
```

Also available: `instrument-template` (force template re-instrumentation),
`split` / `assemble` (edit the draft as per-section files).

Runtime dependencies beyond `pyproject.toml`: `pandoc` on PATH (math → OMML)
and `mmdc` (mermaid-cli) for mermaid figures.

## Consuming from a downstream project

Add as a git submodule and install:

```bash
git submodule add <this-repo> engine
pip install -e engine   # or PYTHONPATH=engine
```

The provided Dockerfile bakes the engine at `/opt/engine` and prefers a
mounted `/app/engine` checkout, so image and submodule stay interchangeable.
