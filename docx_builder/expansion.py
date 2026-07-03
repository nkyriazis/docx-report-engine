"""Shared macro expansion and validation for the build pipeline.

Used by both build.py (compiled MD pipeline) and md_parser.py (DOCX pipeline).
"""
from __future__ import annotations
import re

# ---------------------------------------------------------------------------
# YAML (simple key: value)
# ---------------------------------------------------------------------------

def load_yaml(path: str) -> dict[str, str]:
    """Parse a simple key: value YAML file (no nesting, no lists)."""
    defs: dict[str, str] = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            key, _, value = line.partition(":")
            key = key.strip().strip('"').strip("'")
            value = value.strip().strip('"').strip("'")
            if key:
                defs[key] = value
    return defs

# ---------------------------------------------------------------------------
# Acronym macros
# ---------------------------------------------------------------------------

RE_ACRONYM = re.compile(r":([A-Za-z][A-Za-z0-9/\-\.]*?)(?!\{):")
_AC_MARKER = "<!-- ACRONYM_TABLE -->"


def expand_acronyms(text: str, defs: dict[str, str]) -> tuple[str, set[str]]:
    """Expand :KEY: macros. First use → "Definition (KEY)", subsequent → "KEY".

    Returns (expanded_text, set_of_used_keys).
    """
    seen: set[str] = set()

    def _repl(m: re.Match) -> str:
        token = m.group(1)
        key = (
            token
            if token in defs
            else (token[:-1] if token.endswith("s") and token[:-1] in defs else None)
        )
        if key is None:
            return m.group(0)
        if key not in seen:
            seen.add(key)
            if token != key:
                return f"{defs[key]}s ({token})"
            return f"{defs[key]} ({key})"
        return token

    text = RE_ACRONYM.sub(_repl, text)
    return text, seen


def inject_acronym_table(text: str, used: set[str], defs: dict[str, str]) -> str:
    """Replace <!-- ACRONYM_TABLE --> marker with a sorted table of used acronyms."""
    if _AC_MARKER not in text:
        return text
    rows = sorted(((k, defs[k]) for k in used if k in defs), key=lambda kv: kv[0].lower())
    table = "\n".join([
        "## Abbreviations and acronyms",
        "",
        "| **Abbreviation / Acronym** | **Definition** |",
        "| --- | --- |",
        *[f"| {k} | {v} |" for k, v in rows],
    ]) + "\n"
    return text.replace(_AC_MARKER, table)


def check_acronyms(text: str, defs: dict[str, str]) -> dict[str, list[int]]:
    """Return {token: [line_numbers]} for acronyms used in text but missing from defs."""
    undefined: dict[str, list[int]] = {}
    for i, line in enumerate(text.splitlines(), 1):
        for m in RE_ACRONYM.finditer(line):
            token = m.group(1)
            key = token if token in defs else (
                token[:-1] if token.endswith("s") and token[:-1] in defs else None
            )
            if key is None:
                undefined.setdefault(token, []).append(i)
    return undefined


_UNIVERSALS_PATH = "universals.yml"


def _load_universals(path: str = _UNIVERSALS_PATH) -> frozenset[str]:
    """Load well-known terms from universals.yml (keys only)."""
    try:
        defs = load_yaml(path)
    except FileNotFoundError:
        return frozenset()
    return frozenset(defs.keys())


def check_plain_acronyms(text: str, universals: frozenset[str] | None = None) -> dict[str, list[int]]:
    """Return {word: [line_numbers]} for bare uppercase sequences (2+ chars)
    in the main body that should be wrapped as :KEY: macros instead.

    Skips front matter, code/mermaid blocks, already-wrapped :TOKEN: patterns,
    mermaid diagram labels (:fig{...}:), section refs (:ref{...}:), table refs
    (:tab{...}:), inline code spans, and cross-reference labels ({#...}).
    Terms in *universals* are also skipped (well-known, no definition needed).
    """
    if universals is None:
        universals = _load_universals()
    lines = text.split("\n")

    # Find main body start — first ## heading that is not the TOC placeholder.
    body_start = 0
    for i, line in enumerate(lines):
        if re.match(r"^##\s+\d", line):
            body_start = i
            break

    in_code_block = False
    results: dict[str, list[int]] = {}

    for ln in range(body_start, len(lines)):
        line = lines[ln]
        stripped = line.strip()

        # Track code / mermaid fences
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            continue

        # Skip figure caption lines (**...**) and italic caption lines (*...*)
        if stripped.startswith("*:fig{") or stripped.startswith("**:fig{"):
            continue

        # Remove already-wrapped :TOKEN: macros (including :fig{...}: / :ref{...}: / :tab{...}:)
        cleaned = RE_ACRONYM.sub("", line)
        cleaned = re.sub(r":(fig|ref|tab)\{[^}]*\}:?", "", cleaned)
        # Remove HTML comments
        cleaned = re.sub(r"<!--.*?-->", "", cleaned)
        # Remove inline code spans (backtick-protected content)
        cleaned = re.sub(r"`[^`]*`", "", cleaned)
        # Remove {#section-label} cross-reference anchors
        cleaned = re.sub(r"\{#[^}]+\}", "", cleaned)

        # Find bare sequences of 2+ consecutive uppercase letters
        for m in re.finditer(r'(?<![A-Za-z])([A-Z]{2,})(?![A-Za-z])', cleaned):
            word = m.group(1)
            # Skip universals — well-known terms that don't need a definition
            if word in universals:
                continue
            # Skip when followed by a digit: WP3, WP4, etc
            after = cleaned[m.end() : m.end() + 1] if m.end() < len(cleaned) else ""
            if after and after.isdigit():
                continue
            # Skip pattern D3.1, T4.1 etc — those are task / deliverable refs
            ctx_before = cleaned[max(0, m.start() - 2) : m.start()]
            if re.match(r'[DT]\d$', ctx_before):
                continue
            results.setdefault(word, []).append(ln + 1)

    return results

# ---------------------------------------------------------------------------
# Figure macros
# ---------------------------------------------------------------------------

RE_FIG = re.compile(r":fig\{([^}]+)\}:")
RE_FIG_RENDER = re.compile(r":fig\{([^}]+)\}:([a-z])?")
RE_FIG_CAPTION = re.compile(r"^\*\*:fig\{([^}]+)\}:\s*(.+?)\*\*\s*$", re.MULTILINE)
_FIG_MARKER = "<!-- FIGURE_TABLE -->"


def expand_figures(text: str) -> tuple[str, dict[int, str]]:
    """Expand :fig{ID}: macros to sequential numbers.

    Returns (expanded_text, num_to_fid_map) where num_to_fid maps
    figure number (int) → original figure ID (str).
    """
    nums: dict[str, int] = {}
    n = 0
    for m in RE_FIG.finditer(text):
        fid = m.group(1)
        if fid not in nums:
            n += 1
            nums[fid] = n

    captions: list[tuple[int, str]] = []
    seen: set[str] = set()
    for m in RE_FIG_CAPTION.finditer(text):
        fid, cap = m.group(1), m.group(2).rstrip()
        if fid in nums and fid not in seen:
            seen.add(fid)
            captions.append((nums[fid], cap))
    captions.sort(key=lambda x: x[0])

    if _FIG_MARKER in text:
        lines = ["**List of Figures**", ""]
        for num, cap in captions:
            lines.append(f"- Figure {num}: {cap}")
        text = text.replace(_FIG_MARKER, "\n".join(lines) + "\n")

    def _repl(m: re.Match) -> str:
        fid, suffix = m.group(1), m.group(2)
        if fid not in nums:
            return m.group(0)
        return f"Figure {nums[fid]}{suffix}" if suffix else f"Figure {nums[fid]}:"

    text = RE_FIG_RENDER.sub(_repl, text)
    num_to_fid = {v: k for k, v in nums.items()}
    return text, num_to_fid


def check_figures(text: str) -> list[str]:
    """Return list of errors for uncaptioned figures, duplicate captions, missing marker."""
    errors: list[str] = []
    used_ids: set[str] = set()
    for m in RE_FIG.finditer(text):
        used_ids.add(m.group(1))

    captioned: dict[str, int] = {}
    for m in RE_FIG_CAPTION.finditer(text):
        fid = m.group(1)
        captioned[fid] = captioned.get(fid, 0) + 1

    for fid, count in captioned.items():
        if count > 1:
            errors.append(f"Duplicate bold caption for :fig{{{fid}}}: ({count} times)")

    for fid in sorted(used_ids - set(captioned)):
        errors.append(f":fig{{{fid}}}: used but has no bold caption definition")

    return errors
