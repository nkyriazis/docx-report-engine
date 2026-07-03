"""Parse markdown lines into a list of ContentNode structural elements.

Two-pass pipeline:
  1. Structural: markdown-it-py AST → ContentNodes (headings, paragraphs, lists, tables, code fences)
  2. Macros: detect figure/table/math patterns and transform

Generic API:
  parse_body(lines) → list[ContentNode]
"""
from __future__ import annotations
import re
from markdown_it import MarkdownIt

from .expansion import RE_ACRONYM
from .schema import Run, ContentNode

# ---------------------------------------------------------------------------
# markdown-it-py instance (GFM tables supported via default preset)
# ---------------------------------------------------------------------------
_mdit = MarkdownIt("default")

# ---------------------------------------------------------------------------
# Compiled patterns — inline parsing (used by parse_inline)
# ---------------------------------------------------------------------------
RE_SECREF = re.compile(r':ref\{([^}]+)\}:')
RE_FIGREF = re.compile(r':fig\{([^}]+)\}:')
RE_TABREF = re.compile(r':tab\{([^}]+)\}:')
RE_BOLD = re.compile(r'\*\*(.+?)\*\*')
RE_ITALIC = re.compile(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)')
RE_CODE = re.compile(r'`([^`]+)`')
RE_HTML_SPAN = re.compile(
    r'<span\s+style="([^"]*color:\s*([^;"]+))"[^>]*>(.+?)</span>',
    re.DOTALL,
)

# ---------------------------------------------------------------------------
# Compiled patterns — macro detection
# ---------------------------------------------------------------------------
RE_FIG_MACRO = re.compile(r'^:fig\{([^}]+)\}:\s*(.*?)\s*(?:\[([^\]]+)\])?$')
RE_TAB_MACRO = re.compile(r'^:tab\{([^}]+)\}:\s*(.*?)$')
RE_WIDTH_HINT = re.compile(r'^(\d+(?:\.\d+)?)(%|cm)$')
RE_HEADING_ID = re.compile(r'\s*\{#([^}]+)\}\s*$')
RE_HAS_INLINE_MATH = re.compile(r'\$[^$\n]+\$')

# ---------------------------------------------------------------------------
# Mapping from Markdown heading depth to schema type
# ---------------------------------------------------------------------------
_HEADING_TYPE = {2: 'h1', 3: 'h2', 4: 'h3', 5: 'h4', 6: 'h4'}


# ---------------------------------------------------------------------------
# Inline parsing
# ---------------------------------------------------------------------------

def _refs_to_sentinels(text: str) -> str:
    """Convert :ref{id}:, :fig{id}:, and :tab{id}: tokens to sentinels."""
    text = RE_SECREF.sub(lambda m: f'@@SECREF:{m.group(1)}@@', text)
    text = RE_FIGREF.sub(lambda m: f'@@FIGREF:{m.group(1)}@@', text)
    text = RE_TABREF.sub(lambda m: f'@@TABREF:{m.group(1)}@@', text)
    return text


def parse_inline(text: str) -> list[Run]:
    """Parse inline markdown into Runs."""
    text = RE_ACRONYM.sub(lambda m: m.group(1), text)

    runs: list[Run] = []
    i = 0
    while i < len(text):
        m_bold = RE_BOLD.search(text, i)
        m_ital = RE_ITALIC.search(text, i)
        m_ref  = RE_SECREF.search(text, i)
        m_fref = RE_FIGREF.search(text, i)
        m_tref = RE_TABREF.search(text, i)
        m_code = RE_CODE.search(text, i)
        m_span = RE_HTML_SPAN.search(text, i)

        candidates = []
        if m_bold:  candidates.append((m_bold.start(),  0, m_bold))
        if m_ital:  candidates.append((m_ital.start(),  1, m_ital))
        if m_code:  candidates.append((m_code.start(),  2, m_code))
        if m_span:  candidates.append((m_span.start(),  3, m_span))
        if m_ref:   candidates.append((m_ref.start(),   4, m_ref))
        if m_fref:  candidates.append((m_fref.start(),  5, m_fref))
        if m_tref:  candidates.append((m_tref.start(),  6, m_tref))

        if not candidates:
            tail = text[i:]
            if tail:
                runs.append(Run(text=tail))
            break

        _, _, earliest = min(candidates)
        before = text[i:earliest.start()]
        if before:
            runs.append(Run(text=before))

        if earliest is m_bold:
            bold_content = earliest.group(1)
            if RE_SECREF.search(bold_content) or RE_FIGREF.search(bold_content) or RE_TABREF.search(bold_content):
                inner_runs = parse_inline(bold_content)
                for r in inner_runs:
                    r.bold = True
                runs.extend(inner_runs)
            else:
                runs.append(Run(text=bold_content, bold=True))
        elif earliest is m_ital:
            italic_content = earliest.group(1)
            if RE_SECREF.search(italic_content) or RE_FIGREF.search(italic_content) or RE_TABREF.search(italic_content):
                inner_runs = parse_inline(italic_content)
                for r in inner_runs:
                    r.italic = True
                runs.extend(inner_runs)
            else:
                runs.append(Run(text=italic_content, italic=True))
        elif earliest is m_code:
            runs.append(Run(text=earliest.group(1), code=True))
        elif earliest is m_span:
            color_val = earliest.group(2).strip()
            inner_text = earliest.group(3)
            inner_runs = parse_inline(inner_text)
            for r in inner_runs:
                r.color = color_val
            runs.extend(inner_runs)
        elif earliest is m_ref:
            ref_id = earliest.group(1)
            runs.append(Run(text=f'@@SECREF:{ref_id}@@', ref_id=ref_id))
        elif earliest is m_fref:
            fig_id = earliest.group(1)
            runs.append(Run(text=f'@@FIGREF:{fig_id}@@', fig_ref_id=fig_id))
        else:
            tab_id = earliest.group(1)
            runs.append(Run(text=f'@@TABREF:{tab_id}@@', tab_ref_id=tab_id))

        i = earliest.end()

    return runs


# ---------------------------------------------------------------------------
# Inline text extraction from markdown-it-py tokens
# ---------------------------------------------------------------------------

def _inline_text(inline_token) -> str:
    """Assemble plain text from an inline token's children, preserving markdown syntax."""
    parts: list[str] = []
    for child in (inline_token.children or []):
        ct = child.type
        if ct == "text":
            parts.append(child.content)
        elif ct == "code_inline":
            parts.append(f"`{child.content}`")
        elif ct == "strong_open":
            parts.append("**")
        elif ct == "strong_close":
            parts.append("**")
        elif ct == "em_open":
            parts.append("*")
        elif ct == "em_close":
            parts.append("*")
        elif ct in ("softbreak", "hardbreak"):
            parts.append(" ")
        elif ct == "image":
            alt = child.content or ""
            src = dict(child.attrs or {}).get("src", "")
            parts.append(f"![{alt}]({src})")
    return "".join(parts)


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
    runs = parse_inline(assembled)
    node = ContentNode(type="p", runs=runs, _raw=assembled)
    if RE_HAS_INLINE_MATH.search(assembled):
        node.has_math = True
        node.math_raw = _refs_to_sentinels(RE_ACRONYM.sub(lambda m: m.group(1), assembled))
    return (node, i + 3)


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
            item_text = _inline_text(inline_tok) if inline_tok else ""
            runs = parse_inline(item_text)
            if runs:
                node = ContentNode(type=kind, runs=runs, level=depth)
                if RE_HAS_INLINE_MATH.search(item_text):
                    node.has_math = True
                    node.math_raw = _refs_to_sentinels(RE_ACRONYM.sub(lambda m: m.group(1), item_text))
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

        else:
            i += 1

    return nodes


# ---------------------------------------------------------------------------
# Macro pass — figure/table/math detection
# ---------------------------------------------------------------------------

def _process_macros(nodes: list[ContentNode]) -> list[ContentNode]:
    """Detect and transform macro patterns: figures, tables, display math.

    Linear scan. State: pending figures/tables, current context ids.
    """
    result: list[ContentNode] = []
    pending_fig: dict[str, ContentNode] = {}
    pending_tab: dict[str, ContentNode] = {}
    cur_fig: str | None = None
    cur_tab: str | None = None

    for node in nodes:
        t = node.type

        # --- Heading: reset figure context ---
        if t in ("h1", "h2", "h3", "h4"):
            cur_fig = None
            result.append(node)
            continue

        # --- List items: pass through ---
        if t in ("bullet", "numbered"):
            result.append(node)
            continue

        # --- Display math ---
        if t == "p" and node._raw.startswith("$$") and node._raw.endswith("$$"):
            result.append(ContentNode(type="math_display", math_src=node._raw[2:-2].strip()))
            continue

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
            continue

        # --- Figure caption: *:fig{id}: Caption* ---
        if fm and raw.startswith("*") and not raw.startswith("**"):
            fig_id = fm.group(1)
            caption = re.sub(r"^[a-z]:\s+", "", fm.group(2).strip())
            if fig_id in pending_fig:
                fig = pending_fig.pop(fig_id)
                fig.fig_caption = parse_inline(caption)
                result.append(fig)
            cur_fig = None
            continue

        # --- Table label: **:tab{id}: Caption** ---
        tm = RE_TAB_MACRO.match(raw.strip("*").strip())
        if tm and raw.startswith("**") and raw.endswith("**"):
            tab_id = tm.group(1)
            caption = tm.group(2).strip()
            pending_tab[tab_id] = ContentNode(
                type="table", tbl_id=tab_id, tbl_caption=parse_inline(caption),
            )
            cur_tab = tab_id
            continue

        # --- Table caption: *:tab{id}: Caption* ---
        if tm and raw.startswith("*") and not raw.startswith("**"):
            tab_id = tm.group(1)
            if tab_id in pending_tab:
                result.append(pending_tab.pop(tab_id))
            cur_tab = None
            continue

        # --- Image: attach to pending figure ---
        if t == "_image":
            if cur_fig and cur_fig in pending_fig:
                pending_fig[cur_fig].fig_path = node.fig_path
                pending_fig[cur_fig].fig_alt = node.fig_alt
            continue

        # --- Mermaid: attach to pending figure ---
        if t == "_mermaid":
            if cur_fig and cur_fig in pending_fig:
                fig = pending_fig[cur_fig]
                fig.fig_path = f"_mermaid:{cur_fig}"
                fig.mermaid_src = node.mermaid_src
            continue

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
            continue

        # --- Regular paragraph ---
        if t == "p":
            result.append(node)

    # Flush unclosed pending
    result.extend(pending_fig.values())
    result.extend(pending_tab.values())

    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _parse_body(lines: list[str], content: list[ContentNode]) -> None:
    """Parse body lines into ContentNode list via markdown-it-py AST."""
    tokens = _mdit.parse("\n".join(lines))
    structural = _structural_pass(tokens)
    content.extend(_process_macros(structural))


def parse_body(lines: list[str]) -> list[ContentNode]:
    """Parse markdown lines into ContentNode list."""
    content: list[ContentNode] = []
    _parse_body(lines, content)
    return content
