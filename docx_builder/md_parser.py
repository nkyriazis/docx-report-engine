"""Parse markdown lines into a list of ContentNode structural elements.

Two-pass pipeline:
  1. Structural: markdown-it-py AST → ContentNodes (headings, paragraphs, lists, tables, code fences)
  2. Macros: detect figure/table patterns and transform

Inline formatting (bold, italic, code, math) is taken directly from the
markdown-it inline token tree — there is no re-parsing of serialized text.
The only regex layer left at the inline level handles syntax that is not
markdown: :ref{}/:fig{}/:tab{} cross-references and color <span>s.

Generic API:
  parse_body(lines) → list[ContentNode]
"""
from __future__ import annotations
import re
from markdown_it import MarkdownIt
from mdit_py_plugins.dollarmath import dollarmath_plugin

from .expansion import RE_ACRONYM
from .schema import Run, ContentNode

# ---------------------------------------------------------------------------
# markdown-it-py instance (GFM tables via default preset; $/$$ math via
# dollarmath, which protects math content from emphasis parsing)
# ---------------------------------------------------------------------------
_mdit = MarkdownIt("default").use(dollarmath_plugin)

# ---------------------------------------------------------------------------
# Compiled patterns — non-markdown inline syntax (cross-refs, color spans)
# ---------------------------------------------------------------------------
RE_SECREF = re.compile(r':ref\{([^}]+)\}:')
RE_FIGREF = re.compile(r':fig\{([^}]+)\}:')
RE_TABREF = re.compile(r':tab\{([^}]+)\}:')

_RE_INLINE_FEATURE = re.compile(
    r'(?P<secref>:ref\{(?P<secid>[^}]+)\}:)'
    r'|(?P<figref>:fig\{(?P<figid>[^}]+)\}:)'
    r'|(?P<tabref>:tab\{(?P<tabid>[^}]+)\}:)'
    r'|(?P<span_open><span\s+style="[^"]*color:\s*(?P<color>[^;"]+)[^"]*"[^>]*>)'
    r'|(?P<span_close></span>)'
    r'|(?P<footref>@@FOOTREF:(?P<footkey>[^@]+)@@)'
)

# ---------------------------------------------------------------------------
# Compiled patterns — macro detection
# ---------------------------------------------------------------------------
RE_FIG_MACRO = re.compile(r'^:fig\{([^}]+)\}:\s*(.*?)\s*(?:\[([^\]]+)\])?$')
RE_TAB_MACRO = re.compile(r'^:tab\{([^}]+)\}:\s*(.*?)$')
RE_WIDTH_HINT = re.compile(r'^(\d+(?:\.\d+)?)(%|cm)$')
RE_HEADING_ID = re.compile(r'\s*\{#([^}]+)\}\s*$')
RE_COMMENT_HEADER = re.compile(r'^comment:\s*(\S.*)$')

# Footnotes — `[^key]` inline references and `[^key]: body` definition
# paragraphs. Definitions are extracted from the raw lines before the
# markdown-it pass (a bare-URL body would otherwise be swallowed as a link
# reference definition); references are converted to @@FOOTREF:key@@
# sentinels at the same time so bracket-parsing can't split them.
RE_FOOTNOTE_DEF = re.compile(r'^\[\^([A-Za-z0-9_.:+\-]+)\]:\s+(\S.*)$')
RE_FOOTNOTE_REF = re.compile(r'\[\^([A-Za-z0-9_.:+\-]+)\]')

# Node types a blockquote comment is allowed to attach to — plain-text nodes
# whose rendering goes through the @@INLINEFMT:n@@ placeholder path. Math
# paragraphs, figures, tables, and math_display blocks are out of scope.
_COMMENTABLE_TYPES = {'h1', 'h2', 'h3', 'h4', 'p', 'bullet', 'numbered'}

# ---------------------------------------------------------------------------
# Mapping from Markdown heading depth to schema type
# ---------------------------------------------------------------------------
_HEADING_TYPE = {2: 'h1', 3: 'h2', 4: 'h3', 5: 'h4', 6: 'h4'}


def _refs_to_sentinels(text: str) -> str:
    """Convert :ref{id}:, :fig{id}:, and :tab{id}: tokens to sentinels."""
    text = RE_SECREF.sub(lambda m: f'@@SECREF:{m.group(1)}@@', text)
    text = RE_FIGREF.sub(lambda m: f'@@FIGREF:{m.group(1)}@@', text)
    text = RE_TABREF.sub(lambda m: f'@@TABREF:{m.group(1)}@@', text)
    return text


# ---------------------------------------------------------------------------
# Inline token tree → Runs
# ---------------------------------------------------------------------------

class _InlineState:
    """Formatting state while walking an inline token tree."""
    __slots__ = ('bold', 'italic', 'colors')

    def __init__(self) -> None:
        self.bold = 0
        self.italic = 0
        self.colors: list[str] = []

    def make_run(self, text: str, **extra) -> Run:
        return Run(
            text=text,
            bold=self.bold > 0,
            italic=self.italic > 0,
            color=self.colors[-1] if self.colors else '',
            **extra,
        )


def _emit_text(text: str, state: _InlineState, runs: list[Run]) -> None:
    """Split a text fragment on refs / color-span tags and emit Runs."""
    text = RE_ACRONYM.sub(lambda m: m.group(1), text)
    pos = 0
    for m in _RE_INLINE_FEATURE.finditer(text):
        before = text[pos:m.start()]
        if before:
            runs.append(state.make_run(before))
        if m.group('secref'):
            rid = m.group('secid')
            runs.append(state.make_run(f'@@SECREF:{rid}@@', ref_id=rid))
        elif m.group('figref'):
            fid = m.group('figid')
            runs.append(state.make_run(f'@@FIGREF:{fid}@@', fig_ref_id=fid))
        elif m.group('tabref'):
            tid = m.group('tabid')
            runs.append(state.make_run(f'@@TABREF:{tid}@@', tab_ref_id=tid))
        elif m.group('footref'):
            fkey = m.group('footkey')
            runs.append(state.make_run(m.group(0), footnote_key=fkey))
        elif m.group('span_open'):
            state.colors.append(m.group('color').strip())
        elif m.group('span_close'):
            if state.colors:
                state.colors.pop()
        pos = m.end()
    tail = text[pos:]
    if tail:
        runs.append(state.make_run(tail))


def _merge_adjacent(runs: list[Run]) -> list[Run]:
    """Merge neighboring plain-text runs with identical formatting."""
    merged: list[Run] = []
    for run in runs:
        prev = merged[-1] if merged else None
        if (
            prev is not None
            and not (prev.ref_id or prev.fig_ref_id or prev.tab_ref_id or prev.footnote_key)
            and not (run.ref_id or run.fig_ref_id or run.tab_ref_id or run.footnote_key)
            and (prev.bold, prev.italic, prev.code, prev.color)
            == (run.bold, run.italic, run.code, run.color)
        ):
            prev.text += run.text
        else:
            merged.append(run)
    return merged


def _runs_from_children(children: list) -> list[Run]:
    """Build Runs by walking an inline token's children."""
    state = _InlineState()
    runs: list[Run] = []
    for child in (children or []):
        ct = child.type
        if ct == 'text':
            _emit_text(child.content, state, runs)
        elif ct == 'code_inline':
            if child.content:
                runs.append(state.make_run(child.content, code=True))
        elif ct == 'strong_open':
            state.bold += 1
        elif ct == 'strong_close':
            state.bold -= 1
        elif ct == 'em_open':
            state.italic += 1
        elif ct == 'em_close':
            state.italic -= 1
        elif ct == 'math_inline':
            # kept verbatim; paragraphs with math are re-rendered from
            # math_raw by the pandoc post-processing step
            runs.append(state.make_run(f'${child.content}$'))
        elif ct in ('softbreak', 'hardbreak'):
            runs.append(state.make_run(' '))
        elif ct == 'html_inline':
            _emit_text(child.content, state, runs)
        elif ct == 'image':
            alt = child.content or ''
            src = dict(child.attrs or {}).get('src', '')
            runs.append(state.make_run(f'![{alt}]({src})'))
    return _merge_adjacent(runs)


def _has_inline_math(children: list) -> bool:
    return any(c.type == 'math_inline' for c in (children or []))


def parse_inline(text: str) -> list[Run]:
    """Parse an inline markdown string into Runs."""
    tokens = _mdit.parseInline(text or '')
    if not tokens:
        return []
    return _runs_from_children(tokens[0].children)


# ---------------------------------------------------------------------------
# Inline serializer — markdown source text for macro matching and pandoc
# ---------------------------------------------------------------------------

def _inline_text(inline_token) -> str:
    """Assemble markdown text from an inline token's children.

    One-way only: used for macro-pattern matching on raw lines, for table
    cell strings, and as pandoc input for paragraphs with inline math.
    Runs are never re-parsed from this.
    """
    parts: list[str] = []
    for child in (inline_token.children or []):
        ct = child.type
        if ct == 'text':
            parts.append(child.content)
        elif ct == 'code_inline':
            parts.append(f'`{child.content}`')
        elif ct in ('strong_open', 'strong_close'):
            parts.append('**')
        elif ct in ('em_open', 'em_close'):
            parts.append('*')
        elif ct == 'math_inline':
            parts.append(f'${child.content}$')
        elif ct in ('softbreak', 'hardbreak'):
            parts.append(' ')
        elif ct == 'html_inline':
            parts.append(child.content)
        elif ct == 'image':
            alt = child.content or ''
            src = dict(child.attrs or {}).get('src', '')
            parts.append(f'![{alt}]({src})')
    return ''.join(parts)


def _math_raw(assembled: str) -> str:
    """Pandoc input for a paragraph with inline math: acronym-stripped,
    cross-references converted to sentinels."""
    return _refs_to_sentinels(RE_ACRONYM.sub(lambda m: m.group(1), assembled))


# ---------------------------------------------------------------------------
# Structural pass — markdown-it-py AST → ContentNodes
# ---------------------------------------------------------------------------

def _parse_heading(tokens: list, i: int, hlevel: int) -> tuple[ContentNode, int]:
    """Parse heading_open + inline + heading_close. Return (node, new_i)."""
    inline_tok = tokens[i + 1]
    raw = _inline_text(inline_tok)

    sec_id = ""
    m = RE_HEADING_ID.search(raw)
    if m:
        sec_id = m.group(1)
        raw = raw[:m.start()]

    num_m = re.match(r"^((?:\d+\.?)+)\s+", raw.strip())
    sec_number = num_m.group(1) if num_m else ""
    display = re.sub(r"^(\d+\.?)+\s+", "", raw.strip())

    return (ContentNode(
        type=_HEADING_TYPE.get(hlevel, "h4"),
        runs=parse_inline(display),
        sec_id=sec_id,
        sec_number=sec_number,
    ), i + 3)


def _parse_paragraph(tokens: list, i: int) -> tuple[ContentNode, int]:
    """Parse paragraph_open + inline + paragraph_close. Return (node, i+3)."""
    inline_tok = tokens[i + 1]
    children = inline_tok.children or []
    assembled = _inline_text(inline_tok)

    # Inline image
    for child in children:
        if child.type == "image":
            return (ContentNode(
                type="_image",
                fig_path=dict(child.attrs or {}).get("src", ""),
                fig_alt=child.content or "",
                _raw=assembled,
            ), i + 3)

    # Regular paragraph
    node = ContentNode(type="p", runs=_runs_from_children(children), _raw=assembled)
    if _has_inline_math(children):
        node.has_math = True
        node.math_raw = _math_raw(assembled)
    return (node, i + 3)


def _parse_comment_blockquote(tokens: list, i: int) -> tuple[str, list[list[Run]], int]:
    """Parse a `> comment: Author` blockquote. Strict: this dialect reserves
    blockquotes for author comments — the first paragraph inside must read
    exactly 'comment: <author>', and every other paragraph inside becomes one
    paragraph of the comment body. Returns (author, body_paragraphs, i past
    the blockquote_close).
    """
    assert tokens[i].type == "blockquote_open"
    j = i + 1
    paragraphs: list[list[Run]] = []
    while j < len(tokens) and tokens[j].type != "blockquote_close":
        if tokens[j].type != "paragraph_open":
            raise ValueError(
                f"Blockquotes are reserved for comments in this dialect and may "
                f"only contain plain paragraphs starting with 'comment: <author>' "
                f"— found {tokens[j].type!r} inside one."
            )
        node, j = _parse_paragraph(tokens, j)
        if node.type != "p":
            raise ValueError(
                "Unsupported content in a comment blockquote (an image); "
                "comments may only contain plain text paragraphs."
            )
        paragraphs.append(node.runs)
    if j >= len(tokens):
        raise ValueError("Unclosed blockquote — missing the blank line that ends it.")

    if not paragraphs:
        raise ValueError(
            "Empty blockquote — comments must start with 'comment: <author>'."
        )

    header = "".join(r.text for r in paragraphs[0]).strip()
    m = RE_COMMENT_HEADER.match(header)
    if not m:
        raise ValueError(
            f"Blockquotes are reserved for comments in this dialect: the first "
            f"line must read 'comment: <author>' — found {header!r}."
        )
    author = m.group(1).strip()
    body = paragraphs[1:]
    if not body:
        raise ValueError(f"Comment by {author!r} has no body text.")

    return author, body, j + 1  # step past blockquote_close


def _parse_list_item(tokens: list, i: int, kind: str, depth: int) -> tuple[list[ContentNode], int]:
    """Parse list_item_open + content. Return (nodes, new_i).

    Extracts leading paragraph text. Stops at nested list or list_item_close.
    Nested lists are handled by the caller (main loop).
    """
    nodes: list[ContentNode] = []
    j = i + 1  # skip list_item_open

    while j < len(tokens):
        tk = tokens[j].type
        if tk == "paragraph_open":
            inline_tok = tokens[j + 1] if j + 1 < len(tokens) else None
            children = (inline_tok.children or []) if inline_tok else []
            runs = _runs_from_children(children)
            if runs:
                node = ContentNode(type=kind, runs=runs, level=depth)
                if _has_inline_math(children):
                    node.has_math = True
                    node.math_raw = _math_raw(_inline_text(inline_tok))
                nodes.append(node)
            j += 3  # paragraph_open + inline + paragraph_close
            break
        elif tk in ("bullet_list_open", "ordered_list_open"):
            break
        elif tk == "list_item_close":
            break
        j += 1

    return (nodes, j)


def _parse_table(tokens: list, i: int) -> tuple[ContentNode, int]:
    """Parse table_open ... tbody_close. Return (node, new_i)."""
    headers: list[str] = []
    rows: list[list[str]] = []
    current_row: list[str] = []
    in_header = False

    j = i + 1  # skip table_open
    while j < len(tokens):
        tk = tokens[j].type
        if tk == "thead_open":
            in_header = True
        elif tk == "thead_close":
            in_header = False
        elif tk == "tbody_close":
            j += 1
            break
        elif tk in ("th_open", "td_open"):
            if j + 1 < len(tokens) and tokens[j + 1].type == "inline":
                current_row.append(_inline_text(tokens[j + 1]).strip())
        elif tk == "tr_close":
            if current_row:
                if in_header:
                    headers = current_row
                else:
                    rows.append(current_row)
                current_row = []
        j += 1

    return (ContentNode(type="table", tbl_headers=headers, tbl_rows=rows), j)


def _structural_pass(tokens: list) -> list[ContentNode]:
    """Walk mdit AST tokens, produce structural ContentNodes.

    No macro detection. No pending state. Pure token→node mapping.
    """
    nodes: list[ContentNode] = []
    list_stack: list[str] = []
    i = 0

    while i < len(tokens):
        t = tokens[i].type

        if t == "heading_open":
            hlevel = int(tokens[i].tag[1])
            node, i = _parse_heading(tokens, i, hlevel)
            nodes.append(node)

        elif t == "paragraph_open":
            node, i = _parse_paragraph(tokens, i)
            nodes.append(node)

        elif t == "math_block":
            nodes.append(ContentNode(
                type="math_display",
                math_src=(tokens[i].content or "").strip(),
            ))
            i += 1

        elif t == "fence":
            if tokens[i].info and "mermaid" in tokens[i].info:
                nodes.append(ContentNode(
                    type="_mermaid",
                    mermaid_src=(tokens[i].content or "").rstrip("\n"),
                ))
            i += 1

        elif t in ("bullet_list_open", "ordered_list_open"):
            list_stack.append("bullet" if t == "bullet_list_open" else "numbered")
            i += 1

        elif t in ("bullet_list_close", "ordered_list_close"):
            if list_stack:
                list_stack.pop()
            i += 1

        elif t == "list_item_open":
            kind = list_stack[-1] if list_stack else "bullet"
            depth = len(list_stack)
            item_nodes, i = _parse_list_item(tokens, i, kind, depth)
            nodes.extend(item_nodes)

        elif t == "table_open":
            node, i = _parse_table(tokens, i)
            if node.tbl_headers and node.tbl_rows:
                nodes.append(node)

        elif t == "blockquote_open":
            author, body, i = _parse_comment_blockquote(tokens, i)
            nodes.append(ContentNode(type="_comment", comment_author=author, comment_body=body))

        else:
            i += 1

    return nodes


# ---------------------------------------------------------------------------
# Macro pass — figure/table detection
# ---------------------------------------------------------------------------

def _process_macros(nodes: list[ContentNode]) -> list[ContentNode]:
    """Detect and transform macro patterns: figures, tables, and comments.

    Linear scan. State: pending figures/tables/comment, current context ids.
    A `_comment` node (from a blockquote) attaches to whatever the *next*
    node dispatch actually appends to `result` — which may be several input
    nodes later if that next unit is itself a multi-paragraph figure/table
    macro still being assembled.
    """
    result: list[ContentNode] = []
    pending_fig: dict[str, ContentNode] = {}
    pending_tab: dict[str, ContentNode] = {}
    cur_fig: str | None = None
    cur_tab: str | None = None

    def _dispatch(node: ContentNode) -> None:
        nonlocal cur_fig, cur_tab
        t = node.type

        # --- Heading: reset figure context ---
        if t in ("h1", "h2", "h3", "h4"):
            cur_fig = None
            result.append(node)
            return

        # --- List items and display math: pass through ---
        if t in ("bullet", "numbered", "math_display"):
            result.append(node)
            return

        raw = (getattr(node, "_raw", "") or "").strip()

        # --- Figure label: **:fig{id}: Title** [width] ---
        fm = RE_FIG_MACRO.match(raw.strip("*").strip())
        if fm and raw.startswith("**") and raw.endswith("**"):
            fig_id = fm.group(1)
            title = fm.group(2).strip()
            width_raw = fm.group(3) or ""
            width = ""
            wm = RE_WIDTH_HINT.match(width_raw)
            if wm:
                width = f"{wm.group(1)}{wm.group(2)}"
            pending_fig[fig_id] = ContentNode(
                type="figure", fig_id=fig_id,
                fig_title=parse_inline(title), fig_width_hint=width,
            )
            cur_fig = fig_id
            return

        # --- Figure caption: *:fig{id}: Caption* ---
        if fm and raw.startswith("*") and not raw.startswith("**"):
            fig_id = fm.group(1)
            caption = re.sub(r"^[a-z]:\s+", "", fm.group(2).strip())
            if fig_id in pending_fig:
                fig = pending_fig.pop(fig_id)
                fig.fig_caption = parse_inline(caption)
                result.append(fig)
            cur_fig = None
            return

        # --- Table label: **:tab{id}: Caption** ---
        tm = RE_TAB_MACRO.match(raw.strip("*").strip())
        if tm and raw.startswith("**") and raw.endswith("**"):
            tab_id = tm.group(1)
            caption = tm.group(2).strip()
            pending_tab[tab_id] = ContentNode(
                type="table", tbl_id=tab_id, tbl_caption=parse_inline(caption),
            )
            cur_tab = tab_id
            return

        # --- Table caption: *:tab{id}: Caption* ---
        if tm and raw.startswith("*") and not raw.startswith("**"):
            tab_id = tm.group(1)
            if tab_id in pending_tab:
                result.append(pending_tab.pop(tab_id))
            cur_tab = None
            return

        # --- Image: attach to pending figure ---
        if t == "_image":
            if cur_fig and cur_fig in pending_fig:
                pending_fig[cur_fig].fig_path = node.fig_path
                pending_fig[cur_fig].fig_alt = node.fig_alt
            return

        # --- Mermaid: attach to pending figure ---
        if t == "_mermaid":
            if cur_fig and cur_fig in pending_fig:
                fig = pending_fig[cur_fig]
                fig.fig_path = f"_mermaid:{cur_fig}"
                fig.mermaid_src = node.mermaid_src
            return

        # --- Table: attach to pending table ---
        if t == "table":
            if cur_tab and cur_tab in pending_tab:
                tab = pending_tab.pop(cur_tab)
                tab.tbl_headers = node.tbl_headers
                tab.tbl_rows = node.tbl_rows
                result.append(tab)
                cur_tab = None
            elif node.tbl_headers and node.tbl_rows:
                result.append(node)
            return

        # --- Regular paragraph ---
        if t == "p":
            result.append(node)

    pending_comment: ContentNode | None = None
    for node in nodes:
        if node.type == "_comment":
            if pending_comment is not None:
                raise ValueError(
                    f"Two comments in a row with nothing between them (by "
                    f"{pending_comment.comment_author!r} and "
                    f"{node.comment_author!r}) — each comment must attach to "
                    f"a following heading, paragraph, bullet, or numbered item."
                )
            pending_comment = node
            continue

        before = len(result)
        _dispatch(node)

        if pending_comment is not None and len(result) > before:
            target = result[before]
            if target.type not in _COMMENTABLE_TYPES:
                raise ValueError(
                    f"Comment by {pending_comment.comment_author!r} attaches "
                    f"to a {target.type!r} node, which cannot carry a comment "
                    f"(only headings, paragraphs, bullets, and numbered items "
                    f"can — inline math within them is fine)."
                )
            target.comment_author = pending_comment.comment_author
            target.comment_body = pending_comment.comment_body
            pending_comment = None

    # Flush unclosed pending
    result.extend(pending_fig.values())
    result.extend(pending_tab.values())

    if pending_comment is not None:
        raise ValueError(
            f"Comment by {pending_comment.comment_author!r} has nothing "
            f"after it to attach to."
        )

    return result


# ---------------------------------------------------------------------------
# Footnote pre-pass — runs on raw lines, before markdown-it
# ---------------------------------------------------------------------------

def _extract_footnotes(lines: list[str]) -> tuple[list[str], list[tuple[str, str]]]:
    """Split footnote definitions out of the body lines and convert inline
    references to @@FOOTREF:key@@ sentinels.

    A definition is a line reading `[^key]: body`; following non-blank lines
    that are not themselves definitions, headings, or fences continue the
    body (joined with spaces). Fenced code blocks pass through untouched.
    Returns (body_lines, [(key, body_text), ...] in appearance order).
    """
    body: list[str] = []
    defs: list[tuple[str, str]] = []
    in_fence = False
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith('```') or stripped.startswith('~~~'):
            in_fence = not in_fence
            body.append(line)
            i += 1
            continue
        if in_fence:
            body.append(line)
            i += 1
            continue
        m = RE_FOOTNOTE_DEF.match(line)
        if m:
            key, text = m.group(1), m.group(2).strip()
            i += 1
            while i < n:
                cont = lines[i].strip()
                if (not cont or RE_FOOTNOTE_DEF.match(lines[i])
                        or cont.startswith(('#', '```', '~~~'))):
                    break
                text += ' ' + cont
                i += 1
            defs.append((key, text))
            continue
        body.append(RE_FOOTNOTE_REF.sub(lambda m: f'@@FOOTREF:{m.group(1)}@@', line))
        i += 1
    return body, defs


def _footnote_def_node(key: str, text: str) -> ContentNode:
    """Build a `_footnote_def` node, rejecting content a footnote can't hold."""
    runs = parse_inline(text)
    for run in runs:
        if run.ref_id or run.fig_ref_id or run.tab_ref_id:
            raise ValueError(
                f"Footnote [^{key}] contains a cross-reference "
                f"(:ref/:fig/:tab) — footnote bodies support only plain "
                f"formatted text."
            )
        if run.footnote_key:
            raise ValueError(
                f"Footnote [^{key}] references another footnote — nested "
                f"footnotes are not supported."
            )
    return ContentNode(type='_footnote_def', footnote_key=key, runs=runs)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _parse_body(lines: list[str], content: list[ContentNode]) -> None:
    """Parse body lines into ContentNode list via markdown-it-py AST."""
    lines, footnote_defs = _extract_footnotes(lines)
    tokens = _mdit.parse("\n".join(lines))
    structural = _structural_pass(tokens)
    content.extend(_process_macros(structural))
    for key, text in footnote_defs:
        content.append(_footnote_def_node(key, text))


def parse_body(lines: list[str]) -> list[ContentNode]:
    """Parse markdown lines into ContentNode list."""
    content: list[ContentNode] = []
    _parse_body(lines, content)
    return content
