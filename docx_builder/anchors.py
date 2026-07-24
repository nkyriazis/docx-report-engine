"""Anchored numbers and ids — one source of truth, everything else derived.

The problem this solves: a value stated in more than one place (the id "E17", the
count "eighteen campaigns") is a *restatement* of a fact defined elsewhere. Typed
by hand, the copies drift the moment one is edited — the "edit one, the other goes
stale" bug the draft used to track with `KEEP IN SYNC` comments.

Three mechanisms, all resolved in a single pass over the draft (so the compiled-MD
and the DOCX, which branch downstream of it, stay identical):

* **Series labels** — a per-series counter composed with an alias. The item is
  tagged once at its definition site; its id follows document order, and every
  reference expands to the same value::

      :label{E:endpoint-sub}:        ->  E17   (17th :label{E:...}: in order)
      :id{endpoint-sub}:             ->  E17   (anywhere in prose)

  A trailing format spec varies the rendering: ``|n`` emits the bare number
  ("roadmap item :id{scale-llm|n}:" -> "roadmap item 2"), and ``|-`` on a
  :label{} registers the item for the counter/count while rendering nothing —
  for tagging an inline enumeration (the six models, the seven substitutions)
  that has no registry table of its own.

* **Category count** — how high a series counter reached; the labels ARE the
  count, so nothing separate needs syncing::

      :count{E}:  -> 18        :count{E|word}: -> eighteen        :count{E|Word}: -> Eighteen

* **Check** — an assertion in an HTML comment (never rendered), reported at build
  time. `seriescount(S)` and integer arithmetic (`+ - *`, parentheses) are
  available, so self-summing tallies can be pinned::

      <!-- @check seriescount(E) == 18 -->
      <!-- @warn  77 + 1 + 4 == 84 -->

  `@check` failures are build errors; `@warn` failures are warnings.
"""
from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import NamedTuple

# ---------------------------------------------------------------------------
# Number → words (0–99 is more than enough for anything derived from structure)
# ---------------------------------------------------------------------------

_ONES = [
    "zero", "one", "two", "three", "four", "five", "six", "seven", "eight",
    "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
    "sixteen", "seventeen", "eighteen", "nineteen",
]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
         "eighty", "ninety"]


def to_words(n: int) -> str:
    """Spell a non-negative integer in [0, 99]. Raises above that range."""
    if not 0 <= n <= 99:
        raise ValueError(f"to_words only covers 0..99, got {n}")
    if n < 20:
        return _ONES[n]
    tens, ones = divmod(n, 10)
    return _TENS[tens] if ones == 0 else f"{_TENS[tens]}-{_ONES[ones]}"


# ---------------------------------------------------------------------------
# Emit macro :count{SERIES|word}:
# ---------------------------------------------------------------------------
# Emits how high a series counter has reached — the total of a tagged category
# (:count{C}: -> 9). The labels ARE the count, so no separate number to sync.

RE_COUNT = re.compile(r":count\{([^}|]+?)(?:\|(word|Word|num))?\}:")


def expand_counts(text: str) -> tuple[str, list[str]]:
    """Replace ``:count{SERIES}:`` emit macros with the series total.

    Returns (expanded_text, errors). An unresolved macro is left in place and
    reported, mirroring how figure errors fail the build.
    """
    errors: list[str] = []

    def _repl(m: re.Match) -> str:
        spec, fmt = m.group(1), m.group(2) or "num"
        n = series_count(text, spec)
        if n == 0:
            errors.append(f":count{{{spec}}}: — no :label{{{spec}:...}}: definitions found")
            return m.group(0)
        if fmt == "num":
            return str(n)
        word = to_words(n)
        return word[:1].upper() + word[1:] if fmt == "Word" else word

    return RE_COUNT.sub(_repl, text), errors


# ---------------------------------------------------------------------------
# Series labels — counter + alias (the EXX / CXX ids)
# ---------------------------------------------------------------------------
# Two orthogonal primitives, composed:
#   * a per-series COUNTER that steps in document order (C1, C2, C3 …), and
#   * an ALIAS that binds a stable key to the resulting label, so every other
#     mention expands to the same value.
# The definition site — the registry table's #-cell — carries the counter
# step; prose refers to the item by key and never types the number:
#
#     #-cell :  :label{C:baltic}:            -> C6   (6th :label{C:...}: in order)
#     prose  :  the :id{baltic}: recording   -> the C6 recording
#
# Numbering follows document order of the :label{} definitions, so the table
# is authoritative: insert a row and every later label — and every :id{} that
# references it — renumbers on the next build.

# A trailing |FMT chooses how the value renders:
#   (none)  full label       E3   (series prefix + number)
#   |n      bare number      3    (ordinals — "roadmap item 3")
#   |-      define only       ''   (register the item for the count/order, render
#                                   nothing — for tagging an inline enumeration)
RE_LABEL = re.compile(r":label\{([A-Za-z][\w.]*):([\w.\-]+)(?:\|(n|full|-))?\}:")
RE_ID = re.compile(r":id\{([\w.\-]+)(?:\|(n|full))?\}:")


def _render(series: str, n: int, fmt: str | None) -> str:
    if fmt == "-":
        return ""
    if fmt == "n":
        return str(n)
    return f"{series}{n}"


def _build_series(text: str) -> tuple[dict[str, tuple[str, int]], list[str]]:
    """Assign each :label{series:key}: its (series, number) in document order."""
    counters: dict[str, int] = {}
    info: dict[str, tuple[str, int]] = {}
    errors: list[str] = []
    for m in RE_LABEL.finditer(text):
        series, key = m.group(1), m.group(2)
        if key in info:
            errors.append(f":label{{{series}:{key}}}: — duplicate definition of '{key}'")
            continue
        counters[series] = counters.get(series, 0) + 1
        info[key] = (series, counters[series])
    return info, errors


def series_count(text: str, series: str) -> int:
    """Number of distinct :label{series:...}: definitions in the text."""
    return len({m.group(2) for m in RE_LABEL.finditer(text) if m.group(1) == series})


def expand_series_labels(text: str) -> tuple[str, list[str]]:
    """Expand :label{series:key}: (define+emit) and :id{key}: (reference)."""
    info, errors = _build_series(text)

    def _lab(m: re.Match) -> str:
        key = m.group(2)
        if key not in info:
            return m.group(0)
        series, n = info[key]
        return _render(series, n, m.group(3))

    text = RE_LABEL.sub(_lab, text)

    def _id(m: re.Match) -> str:
        key = m.group(1)
        if key not in info:
            errors.append(f":id{{{key}}}: — no :label{{...:{key}}}: defines it")
            return m.group(0)
        series, n = info[key]
        return _render(series, n, m.group(2))

    return RE_ID.sub(_id, text), errors


# ---------------------------------------------------------------------------
# Check directives  <!-- @check <expr> == <expr> -->
# ---------------------------------------------------------------------------

RE_CHECK = re.compile(r"<!--\s*@(check|warn)\s+(.+?)\s*-->", re.DOTALL)
_RE_DERIV = re.compile(r"\bseriescount\(\s*([^)]+?)\s*\)")


def _eval_arith(expr: str) -> int:
    """Evaluate an integer arithmetic expression through a whitelisted AST."""
    node = ast.parse(expr.strip(), mode="eval").body

    def ev(n: ast.AST) -> int:
        if isinstance(n, ast.Constant) and isinstance(n.value, int):
            return n.value
        if isinstance(n, ast.BinOp) and isinstance(n.op, (ast.Add, ast.Sub, ast.Mult)):
            a, b = ev(n.left), ev(n.right)
            return a + b if isinstance(n.op, ast.Add) else a - b if isinstance(n.op, ast.Sub) else a * b
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, ast.USub):
            return -ev(n.operand)
        raise ValueError("only integer + - * and parentheses are allowed")

    return ev(node)


def run_checks(raw_text: str) -> list[dict]:
    """Evaluate every ``@check``/``@warn`` directive in the raw draft.

    Returns a list of dicts: {severity, expr, ok, detail}. Derivations are
    substituted for their integer value first; the remainder is pure
    ``<int> == <int>`` and evaluated safely.
    """
    results: list[dict] = []
    for m in RE_CHECK.finditer(raw_text):
        severity, expr = m.group(1), m.group(2).strip()
        entry = {"severity": severity, "expr": expr, "ok": False, "detail": ""}
        try:
            resolved = _RE_DERIV.sub(
                lambda d: str(series_count(raw_text, d.group(1).strip())),
                expr,
            )
            if "==" not in resolved:
                raise ValueError("check must contain '=='")
            lhs, rhs = resolved.split("==", 1)
            lv, rv = _eval_arith(lhs), _eval_arith(rhs)
            entry["ok"] = lv == rv
            entry["detail"] = f"{expr}  ⇒  {lv} == {rv}"
        except (ValueError, SyntaxError) as exc:
            entry["detail"] = f"{expr}  ⇒  {exc}"
        results.append(entry)
    return results


# ---------------------------------------------------------------------------
# Feature registry — the single source of truth for `features` help
# ---------------------------------------------------------------------------
# ⚠ MAINTENANCE CONTRACT (for anyone, human or agent, editing this module):
# every macro this module implements has ONE entry in FEATURES below, carrying
# a runnable example. `verify_features()` executes each example through the real
# pipeline and also scans this file for macro regexes, so:
#   * change a macro's behaviour without updating its example  -> verify fails;
#   * add a new :macro{}/@directive without a FEATURES entry    -> verify fails.
# Keep them in step; `python -m docx_builder features --verify` must stay green.

class Feature(NamedTuple):
    token: str            # the macro/directive token, e.g. ":label"
    syntax: str
    summary: str
    example: str          # a self-contained snippet
    renders: str | None = None    # expected output of the emit pipeline, or…
    check_ok: bool | None = None   # …expected pass/fail for a @check/@warn


def _apply(text: str) -> str:
    """Run the emit pipeline (counts then labels) over a snippet, as the build does."""
    text, _ = expand_counts(text)
    text, _ = expand_series_labels(text)
    return text


FEATURES: list[Feature] = [
    Feature(
        ":label", ":label{SERIES:key}:  (opt |n bare number, |- silent)",
        "Define a series item; its number follows document order of :label{SERIES:…}:.",
        ":label{E:alpha}: :label{E:beta}:", renders="E1 E2",
    ),
    Feature(
        ":id", ":id{key}:  (opt |n bare number)",
        "Reference a defined item; expands to the same value everywhere.",
        ":label{E:alpha}: :label{E:beta}: — see :id{alpha}: and :id{beta|n}:",
        renders="E1 E2 — see E1 and 2",
    ),
    Feature(
        ":count", ":count{SERIES}:  (opt |word, |Word)",
        "How high a series counter reached — the total of a tagged category.",
        "tag a:label{M:a|-}: b:label{M:b|-}: c:label{M:c|-}: → :count{M}: / :count{M|word}:",
        renders="tag a b c → 3 / three",
    ),
    Feature(
        "@check", "<!-- @check EXPR == EXPR -->",
        "Build-time assertion (error on mismatch). seriescount(S) + integer arithmetic.",
        ":label{E:a}: :label{E:b}: <!-- @check seriescount(E) == 2 -->", check_ok=True,
    ),
    Feature(
        "@warn", "<!-- @warn EXPR == EXPR -->",
        "Same as @check but a warning, not a build error.",
        "<!-- @warn 2 + 2 == 5 -->", check_ok=False,
    ),
]


def _macro_tokens_in_source() -> set[str]:
    """Macro tokens actually implemented (scanned from this file's compiled regexes)."""
    src = Path(__file__).read_text(encoding="utf-8")
    toks: set[str] = set()
    for m in re.finditer(r're\.compile\(\s*r"([^"]*)"', src):
        pat = m.group(1)
        toks.update(":" + t for t in re.findall(r":([a-z]+)\\\{", pat))
        for a, b in re.findall(r"@\((\w+)\|(\w+)\)", pat):
            toks.update({"@" + a, "@" + b})
    return toks


def verify_features() -> list[str]:
    """Return a list of drift problems; empty means help and behaviour agree."""
    problems: list[str] = []
    for f in FEATURES:
        if f.renders is not None:
            got = _apply(f.example)
            if got != f.renders:
                problems.append(f"{f.token}: example renders {got!r}, FEATURES says {f.renders!r}")
        if f.check_ok is not None:
            res = run_checks(f.example)
            got = bool(res and res[0]["ok"])
            if got != f.check_ok:
                problems.append(f"{f.token}: check example ok={got}, FEATURES says {f.check_ok}")
    documented = {f.token for f in FEATURES}
    for tok in _macro_tokens_in_source():
        if tok not in documented:
            problems.append(f"macro {tok} is implemented but has no FEATURES entry")
    return problems


def help_text() -> str:
    """Structured reference for the anchor macros, generated from FEATURES."""
    out = [
        "docx_builder — anchored numbers & ids",
        "=====================================",
        "",
        "One source of truth for every repeated number or id in a draft: edit the",
        "definition and all references follow on the next build. Resolved in a single",
        "pass, so the compiled MD and the DOCX stay identical.",
        "",
    ]
    for f in FEATURES:
        out.append(f"  {f.token:7} {f.syntax}")
        out.append(f"          {f.summary}")
        out.append(f"          e.g.  {f.example}")
        if f.renders is not None:
            out.append(f"           ->   {f.renders}")
        elif f.check_ok is not None:
            out.append(f"           ->   check {'passes' if f.check_ok else 'fails'}")
        out.append("")
    out += [
        "This text is generated from FEATURES in anchors.py; each example is executed",
        "by verify_features(). After editing a macro, run:",
        "    python -m docx_builder features --verify",
    ]
    return "\n".join(out)


def features(*, verify: bool = False) -> None:
    """Print the anchor-macro reference; with --verify, run the examples."""
    print(help_text())
    if verify:
        problems = verify_features()
        if problems:
            print("\nVERIFY FAILED — help is out of step with the code:")
            for p in problems:
                print(f"  - {p}")
            raise SystemExit(1)
        print(f"\nverify: all {len(FEATURES)} feature examples pass.")
