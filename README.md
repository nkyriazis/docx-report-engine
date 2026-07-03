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

## Docker service

The published image bundles everything a build needs — Python deps, pandoc,
mermaid-cli and its Chromium — so downstream projects need no local
toolchain. It's rebuilt and pushed by CI
(`.github/workflows/publish-image.yml`, authenticated with the workflow's
own `GITHUB_TOKEN`) on every push to `main` that touches the engine or the
Dockerfile. Tags: `latest` plus the commit SHA for pinning. The image name
follows `github.repository`, so a fork's CI publishes to that fork's own
`ghcr.io/<owner>/docx-report-engine` package — no edits needed after
forking.

Note: GHCR packages start out private regardless of repository visibility.
Until the package is made public (Package settings → Danger Zone → Change
visibility), pulls need a `docker login ghcr.io` with a token that has
`read:packages`.

## Consuming from a downstream project

Add as a git submodule:

```bash
git submodule add https://github.com/nkyriazis/docx-report-engine.git engine
```

The repository is public, so cloning the submodule needs no credentials —
in CI, `actions/checkout` with `submodules: true` just works.

### Via Docker Compose (recommended)

`docker-compose.yml` at the engine's repo root is the intended entry point
for a downstream project — it resolves the image to run, pulling the
published one and falling back to a local build only if that tag isn't on
the registry (verified with Compose v2: a failed pull for a `build:`-backed
service triggers an automatic local build, no flag needed). Wire it up in
the project's own build script:

```bash
#!/usr/bin/env bash
set -euo pipefail
export DOCX_ENGINE_PROJECT_DIR="$PWD"
export DOCX_ENGINE_TAG="$(git -C engine rev-parse --short HEAD)"
exec docker compose -f engine/docker-compose.yml run --rm engine build \
  --draft draft/draft.md \
  --template documents/template.docx \
  --output report.docx \
  --compiled compiled.md \
  "$@"
```

Pinning `DOCX_ENGINE_TAG` to the submodule's own checked-out commit ties the
image version to the source version by construction — the CI that
publishes the image tags it with the same short SHA `git rev-parse --short`
produces, so they're always the same string. **Updating the engine is then
an explicit, reviewable `git submodule update --remote` + commit** — a
step that shows up as an ordinary diff — rather than a floating `:latest`
that silently drifts once Compose has it cached locally (Compose only
pulls when the resolved tag isn't already present locally).

Other variables the compose file reads (see the file's header comment for
the full contract): `DOCX_ENGINE_IMAGE` overrides the image (e.g. for a
fork's own published package), `DOCX_ENGINE_PROJECT_DIR` is mounted at
`/app` and must be an absolute path — Compose resolves relative volume
paths against its own project directory (the compose file's location by
default), not the caller's `$PWD`, so a bare `.` would mount the wrong tree
when the compose file lives one level down at `engine/`.

### Via pip / PYTHONPATH (local, non-docker builds)

```bash
pip install -e engine   # or PYTHONPATH=engine
```

The Dockerfile bakes the engine at `/opt/engine` but puts a mounted
`/app/engine` first on `PYTHONPATH`, so when a project mounts its repo at
`/app` the submodule checkout takes precedence over the baked copy — image
and submodule stay interchangeable, and the image never needs a rebuild to
test local engine changes (this applies to both the raw `docker run` form
and the Compose service above).

## AI skill: adapting a new template

`skills/jinjify-template/SKILL.md` teaches an AI agent to write a project's
`template_prep.py` for any new organisation template. In a downstream Claude
Code project, expose it with a symlink so it is discoverable as a skill:

```bash
mkdir -p .claude/skills
ln -s ../../engine/skills/jinjify-template .claude/skills/jinjify-template
```

Then `/jinjify-template` (or just asking to adapt a new template) walks the
agent through inventorying the DOCX, deciding the contract mapping, writing
the prep script, and verifying the result against
`docs/TEMPLATE_CONTRACT.md`.
