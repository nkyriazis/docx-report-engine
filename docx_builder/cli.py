"""Command-line interface for the report build pipeline.

Usage:
    python -m docx_builder [OPTIONS]

Pipeline:
    1. Preflight checks (labels, refs, acronyms, figures)
    2. Expand acronym macros (:KEY:) and inject abbreviations table
    3. Expand figure macros (:fig{ID}:) and inject List of Figures
    4. Render Mermaid diagrams to PNG
    5. Compile to DOCX via Jinja template pipeline

Template-specific instrumentation lives OUTSIDE the engine, in a per-project
prep script (default: ./template_prep.py) that exports
``instrument_template(input_path, output_path)`` and optionally
``RENDER_OPTIONS``.  See docs/TEMPLATE_CONTRACT.md.
"""
from __future__ import annotations

import csv
import datetime
import importlib.util
import json
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from .expansion import (
    expand_acronyms as _expand_acronyms,
    inject_acronym_table,
    expand_figures as _expand_figures,
    load_yaml,
    check_acronyms as _check_acronyms,
    check_plain_acronyms,
    check_figures as _check_figures,
    RE_FIG_CAPTION,
)
from .mermaid_render import render_mermaid_figures
from .md_parser import parse_body
from .render import render as render_jinja


# ---------------------------------------------------------------------------
# Per-project template prep loading
# ---------------------------------------------------------------------------

def load_prep_module(prep_path: str):
    """Load a project's template-prep script as a Python module.

    The script must export instrument_template(input_path, output_path);
    it may export RENDER_OPTIONS (dict of render() keyword overrides).
    """
    path = Path(prep_path)
    if not path.exists():
        raise FileNotFoundError(
            f"Template prep script not found: {path} — every project needs one "
            f"(see the engine's TEMPLATE_CONTRACT.md and the jinjify-template skill)"
        )
    spec = importlib.util.spec_from_file_location("_project_template_prep", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    if not hasattr(mod, "instrument_template"):
        raise AttributeError(f"{path} does not export instrument_template(input, output)")
    return mod


# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

_MBLOCK = re.compile(r"```mermaid\n(.*?)```", re.DOTALL)
_MCAPTION = re.compile(r"\*\*Figure\s+(\d+):[^*]*\*\*\s*$")
_VAR_OPEN = re.compile(r'^<!--\s*VAR:(\S+)\s+TYPE:(\S+)(?:\s+COLUMNS:([\w,]+))?\s*-->$')
_VAR_CLOSE = re.compile(r'^<!--\s*ENDVAR\s*-->$')
_SEC_LABEL = re.compile(r"^\s*#{1,6}.*\{#([^}]+)\}\s*$", re.MULTILINE)
_SEC_REF = re.compile(r":ref\{([^}]+)\}:")
_STATIC_SECTION = re.compile(r"\b(Section|Sections|Chapter|Chapters)\s+[0-9]+(\.[0-9]+)*\b")


# ---------------------------------------------------------------------------
# Build report
# ---------------------------------------------------------------------------

class BuildReport:
    def __init__(self):
        self.data = {
            "timestamp_utc": datetime.datetime.now(datetime.timezone.utc)
                .replace(microsecond=0).isoformat().replace("+00:00", "Z"),
            "status": "running",
            "cwd": str(Path.cwd()),
            "steps": [],
            "metrics": {},
            "warnings": [],
            "errors": [],
        }
        self._step_start = {}

    def start_step(self, name):
        self._step_start[name] = time.perf_counter()
        self.data["steps"].append({"name": name, "status": "running"})

    def end_step(self, name, status="ok", details=None):
        elapsed = None
        if name in self._step_start:
            elapsed = round(time.perf_counter() - self._step_start.pop(name), 3)
        for step in reversed(self.data["steps"]):
            if step["name"] == name and step["status"] == "running":
                step["status"] = status
                if elapsed is not None:
                    step["seconds"] = elapsed
                if details:
                    step["details"] = details
                return

    def add_metric(self, key, value):
        self.data["metrics"][key] = value

    def warn(self, msg):
        self.data["warnings"].append(msg)
        print(f"  ⚠  {msg}", file=sys.stderr)

    def error(self, msg):
        self.data["errors"].append(msg)
        print(f"  ✖  {msg}", file=sys.stderr)

    def finalize(self, success):
        self.data["status"] = "ok" if success else "failed"

    def write(self, path):
        Path(path).write_text(json.dumps(self.data, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Preflight
# ---------------------------------------------------------------------------

def preflight_check(draft_text: str) -> dict:
    labels = _SEC_LABEL.findall(draft_text)
    refs = _SEC_REF.findall(draft_text)

    dupes = sorted({lab for lab in labels if labels.count(lab) > 1})
    unknown_refs = sorted({rid for rid in refs if rid not in set(labels)})

    static_internal = []
    for ln, line in enumerate(draft_text.splitlines(), start=1):
        if not _STATIC_SECTION.search(line):
            continue
        if re.search(r"\bD\d+\.\d+\b[,;]?\s+(section|chapter)\s+\d", line, re.IGNORECASE):
            continue
        static_internal.append((ln, line.strip()))

    return {
        "label_count": len(labels),
        "unique_label_count": len(set(labels)),
        "ref_count": len(refs),
        "duplicate_labels": dupes,
        "unknown_refs": unknown_refs,
        "static_internal_refs": static_internal,
    }


# ---------------------------------------------------------------------------
# Acronym / figure wrappers (preserve print side-effects)
# ---------------------------------------------------------------------------

def expand_acronyms(text: str, defs: dict[str, str]) -> tuple[str, set[str]]:
    text, seen = _expand_acronyms(text, defs)
    text = inject_acronym_table(text, seen, defs)
    print(f"   {len(seen)} acronyms expanded")
    return text, seen


def expand_figures(text: str) -> tuple[str, dict[int, str]]:
    captions = len({m.group(1) for m in RE_FIG_CAPTION.finditer(text)})
    text, num_to_fid = _expand_figures(text)
    print(f"   {len(num_to_fid)} figures, {captions} captioned")
    return text, num_to_fid


# ---------------------------------------------------------------------------
# Mermaid rendering + compiled MD replacement
# ---------------------------------------------------------------------------

def render_mermaid(text: str, num_to_fid: dict[int, str], content_nodes: list,
                   figures_dir: str = "figures") -> str:
    render_mermaid_figures(content_nodes, output_dir=str(Path.cwd() / figures_dir))

    errors, chunks, prev = [], [], 0
    for i, m in enumerate(_MBLOCK.finditer(text), 1):
        pre = text[: m.start()].rstrip().split("\n")[-1].strip()
        cm = _MCAPTION.match(pre)
        if cm:
            fig_num = int(cm.group(1))
            fid = num_to_fid.get(fig_num, f"fig{i:02d}")
        else:
            fid = f"fig{i:02d}"
        cap = re.sub(r"\*\*", "", pre).strip() if cm else fid
        png = f"figures/{fid}.png"

        chunks.append(text[prev : m.start()])
        if os.path.exists(png):
            chunks.append(f"![{cap}]({png})")
        else:
            errors.append(fid)
            chunks.append(m.group(0))
        prev = m.end()
    chunks.append(text[prev:])

    if errors:
        print(f"Mermaid errors (missing PNG): {errors}", file=sys.stderr)
        sys.exit(1)

    return "".join(chunks)


# ---------------------------------------------------------------------------
# DOCX validation
# ---------------------------------------------------------------------------

def validate_docx_integrity(docx_path: str) -> dict:
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xml_ns = {"w": ns_w}

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")

    root = ET.fromstring(doc_xml)

    text_nodes = [el.text or "" for el in root.findall(".//w:t", xml_ns)]
    doc_text = "".join(text_nodes)

    leftover_sentinels = doc_text.count("@@SECREF:") + doc_text.count("@@SECLABEL:") + doc_text.count("@@FIGREF:") + doc_text.count("@@TABREF:")
    leftover_raw_ref = doc_text.count(":ref{") + doc_text.count(":fig{") + doc_text.count(":tab{")
    unresolved_refs = len(re.findall(r"\[\?[^\]]+\]", doc_text))

    bookmark_names = [
        el.attrib.get(f"{{{ns_w}}}name", "")
        for el in root.findall(".//w:bookmarkStart", xml_ns)
    ]
    sec_bookmarks = [b for b in bookmark_names if b.startswith("sec_")]

    instr_texts = [el.text or "" for el in root.findall(".//w:instrText", xml_ns)]
    xref_fields = [t for t in instr_texts
                   if "HYPERLINK" in t and "\\l" in t and "sec_" in t]

    return {
        "docx_sec_bookmark_count": len(sec_bookmarks),
        "docx_xref_field_count": len(xref_fields),
        "docx_leftover_sentinel_count": leftover_sentinels,
        "docx_leftover_raw_ref_count": leftover_raw_ref,
        "docx_unresolved_placeholder_count": unresolved_refs,
    }


# ---------------------------------------------------------------------------
# VAR block extraction
# ---------------------------------------------------------------------------

_VAR_OPEN  = re.compile(r"^<!--\s*VAR:(\S+)\s+TYPE:(\S+)(?:\s+COLUMNS:([\w,]+))?\s*-->$")
_VAR_CLOSE = re.compile(r"^<!--\s*ENDVAR\s*-->$")


def _extract_vars(lines: list[str]) -> dict:
    """Extract VAR blocks from draft lines.

    Returns dict of name -> (type, columns, inner_lines).
    """
    blocks: dict = {}
    current = None  # (name, vtype, columns)
    buf: list[str] = []
    for line in lines:
        m = _VAR_OPEN.match(line)
        if m:
            current = (m.group(1), m.group(2), m.group(3) or "")
            buf = []
            continue
        if _VAR_CLOSE.match(line) and current is not None:
            blocks[current[0]] = (current[1], current[2], buf)
            current = None
            buf = []
            continue
        if current is not None:
            buf.append(line)
    return blocks


def _parse_csv_row(line: str, columns: tuple) -> dict:
    """Parse a pipe-delimited row into a dict keyed by columns.

    Uses csv semantics: cells are positional, so empty cells are preserved
    (a bare split-and-drop would shift later columns left).  Optional
    leading/trailing pipes are tolerated.
    """
    row = next(csv.reader([line.strip().strip("|")], delimiter="|"))
    cells = [c.strip() for c in row]
    return {col: (cells[i] if i < len(cells) else "") for i, col in enumerate(columns)}


# ---------------------------------------------------------------------------
# Jinja DOCX build
# ---------------------------------------------------------------------------

def build_jinja_docx(context: dict, body_vars: list[str], template_src: str, template_jinja: str,
                     output_docx: str, prep_path: str) -> None:
    prep = load_prep_module(prep_path)
    render_options = getattr(prep, "RENDER_OPTIONS", {})

    if not Path(template_jinja).exists():
        print("  Instrumenting template...")
        prep.instrument_template(template_src, template_jinja)

    render_jinja(template_jinja, context, body_vars, output_docx,
                 image_base=str(Path.cwd()), **render_options)
    print(f"  Done: {Path(output_docx).name}")


# ---------------------------------------------------------------------------
# Main build function (clize entry point)
# ---------------------------------------------------------------------------

def build(
    *,
    draft: str = "draft/draft.md",
    acronyms: str = "acronyms.yml",
    template: str = "documents/template.docx",
    output: str = "output.docx",
    compiled: str = "compiled.md",
    report: str = "build_report.json",
    no_docx: bool = False,
    figures_dir: str = "figures",
    prep: str = "template_prep.py",
) -> None:
    """Build a DOCX from a Markdown draft."""

    report_obj = BuildReport()
    success = False

    try:
        # -- Preflight -------------------------------------------------------
        report_obj.start_step("preflight")
        draft_text = Path(draft).read_text(encoding="utf-8")
        preflight = preflight_check(draft_text)
        report_obj.add_metric("draft_label_count", preflight["label_count"])
        report_obj.add_metric("draft_unique_label_count",
                              preflight["unique_label_count"])
        report_obj.add_metric("draft_ref_count", preflight["ref_count"])

        if preflight["duplicate_labels"]:
            msg = f"Duplicate section labels: {preflight['duplicate_labels']}"
            report_obj.error(msg)
            raise RuntimeError(msg)
        if preflight["unknown_refs"]:
            msg = f"Unknown section references: {preflight['unknown_refs']}"
            report_obj.error(msg)
            raise RuntimeError(msg)
        if preflight["static_internal_refs"]:
            for ln, line in preflight["static_internal_refs"][:10]:
                report_obj.warn(
                    f"Potential static internal reference at line {ln}: {line}"
                )
            if len(preflight["static_internal_refs"]) > 10:
                report_obj.warn(
                    f"... plus {len(preflight['static_internal_refs']) - 10} "
                    f"more potential static references"
                )

        acronyms_defs = load_yaml(acronyms)
        undefined_acronyms = _check_acronyms(draft_text, acronyms_defs)
        if undefined_acronyms:
            tokens = ", ".join(f":{t}:" for t in sorted(undefined_acronyms))
            report_obj.warn(f"Undefined acronyms: {tokens}")

        plain_acronyms = check_plain_acronyms(draft_text)
        if plain_acronyms:
            for word in sorted(plain_acronyms):
                lines = plain_acronyms[word]
                in_defs = " [DEFINED]" if word in acronyms_defs else ""
                sample = ", ".join(str(x) for x in lines[:5])
                extra = f" ... +{len(lines)-5} more" if len(lines) > 5 else ""
                report_obj.warn(
                    f"Bare acronym '{word}' at lines {sample}{extra}"
                    f"{in_defs} — use :{word}:"
                )
        report_obj.add_metric("draft_plain_acronyms", len(plain_acronyms))
        report_obj.add_metric("draft_acronyms_defined", len(acronyms_defs))

        # Check for manually numbered headings (agents should omit them —
        # the parser auto-numbers from hierarchy)
        _HEADING_NUM = re.compile(r'^(##|###|####)\s+\d+(\.\d+)*\.?\s+')
        manual_headings: list[int] = []
        for ln, line in enumerate(draft_text.splitlines(), 1):
            if _HEADING_NUM.match(line):
                manual_headings.append(ln)
        if manual_headings:
            sample = ", ".join(str(x) for x in manual_headings[:10])
            extra = f" ... +{len(manual_headings)-10} more" if len(manual_headings) > 10 else ""
            report_obj.warn(
                f"Manually numbered heading(s) at lines {sample}{extra}"
                f" — the parser auto-numbers from heading hierarchy"
            )
        report_obj.add_metric("draft_manual_headings", len(manual_headings))

        fig_errors = _check_figures(draft_text)
        if fig_errors:
            msg = f"Figure validation errors: {'; '.join(fig_errors)}"
            report_obj.error(msg)
            raise RuntimeError(msg)

        report_obj.end_step("preflight", details={
            "label_count": preflight["label_count"],
            "ref_count": preflight["ref_count"],
            "potential_static_refs": len(preflight["static_internal_refs"]),
            "acronyms_defined": len(acronyms_defs),
            "figure_errors": 0,
        })

        # -- Acronym expansion -----------------------------------------------
        print("1. Acronym expansion")
        report_obj.start_step("acronym_expansion")
        text, used_acronyms = expand_acronyms(draft_text, acronyms_defs)
        report_obj.end_step("acronym_expansion")

        # -- Parse VAR blocks → context variables ----------------------------
        blocks = _extract_vars(text.splitlines())
        # The DOCX abbreviations table lists only acronyms the draft uses,
        # matching inject_acronym_table's behaviour on the compiled-md side.
        context: dict = {
            'acronyms': [
                {"abbrev": k, "definition": acronyms_defs[k]}
                for k in sorted(used_acronyms, key=str.lower)
            ],
        }
        body_vars: list[str] = []
        for name, (vtype, columns, lines) in blocks.items():
            if vtype == "freeform":
                context[name] = parse_body(lines)
                body_vars.append(name)
            elif vtype == "table":
                cols = tuple(columns.split(",")) if columns else ()
                context[name] = [_parse_csv_row(l, cols) for l in lines if l.strip()]
            else:
                context[name] = "\n".join(lines).strip()
        print(f"   Context variables: {', '.join(sorted(context.keys()))}")

        # -- Figure expansion ------------------------------------------------
        print("2. Figure expansion")
        report_obj.start_step("figure_expansion")
        text, num_to_fid = expand_figures(text)
        report_obj.end_step("figure_expansion")

        # -- Mermaid rendering -----------------------------------------------
        print("3. Mermaid rendering")
        report_obj.start_step("mermaid_rendering")
        # Concatenate every freeform VAR's nodes — a fixed-structure template
        # (see the jinjify-template skill) has one content loop per chapter,
        # each bound to its own VAR, and a Mermaid figure can land in any of
        # them.
        content_nodes = [node for name in body_vars for node in context.get(name, [])]
        text = render_mermaid(text, num_to_fid, content_nodes, figures_dir)
        report_obj.end_step("mermaid_rendering")

        # -- Write compiled markdown -----------------------------------------
        report_obj.start_step("write_compiled_markdown")
        Path(compiled).write_text(text, encoding="utf-8")
        report_obj.end_step("write_compiled_markdown")

        # -- DOCX build ------------------------------------------------------
        if no_docx:
            print("   Skipping DOCX build (--no-docx)")
            report_obj.start_step("docx_build")
            report_obj.end_step("docx_build", details={"skipped": True})
            report_obj.start_step("docx_validation")
            report_obj.end_step("docx_validation", details={"skipped": True})
        else:
            print("4. Jinja DOCX")
            report_obj.start_step("docx_build")
            template_jinja = f"{Path(template).stem}_jinja.docx"
            build_jinja_docx(context, body_vars, template, template_jinja, output, prep)
            report_obj.end_step("docx_build")

            report_obj.start_step("docx_validation")
            integrity = validate_docx_integrity(output)
            for k, v in integrity.items():
                report_obj.add_metric(k, v)

            if integrity["docx_leftover_sentinel_count"] > 0:
                raise RuntimeError(
                    f"DOCX contains {integrity['docx_leftover_sentinel_count']} "
                    f"unresolved sentinels"
                )
            if integrity["docx_leftover_raw_ref_count"] > 0:
                raise RuntimeError(
                    f"DOCX contains {integrity['docx_leftover_raw_ref_count']} "
                    f"raw :ref{{...}} token(s)"
                )
            if integrity["docx_unresolved_placeholder_count"] > 0:
                raise RuntimeError(
                    f"DOCX contains "
                    f"{integrity['docx_unresolved_placeholder_count']} "
                    f"unresolved cross-reference placeholder(s)"
                )
            if integrity["docx_xref_field_count"] < preflight["ref_count"]:
                report_obj.warn(
                    f"DOCX cross-reference fields "
                    f"({integrity['docx_xref_field_count']}) are fewer than "
                    f"draft refs ({preflight['ref_count']})"
                )

            report_obj.end_step("docx_validation", details=integrity)

        success = True

    except Exception as exc:
        report_obj.error(str(exc))
        for step in ["preflight", "acronym_expansion", "figure_expansion",
                      "mermaid_rendering", "write_compiled_markdown",
                      "docx_build", "docx_validation"]:
            report_obj.end_step(step, status="failed")
        report_obj.finalize(False)
        report_obj.write(report)
        print(f"Build failed. See {report}", file=sys.stderr)
        raise
    finally:
        if success:
            report_obj.finalize(True)
            report_obj.write(report)
            print(f"Build report → {report}")
