"""Content node schema for the DOCX builder pipeline."""
from __future__ import annotations
from dataclasses import dataclass, field


@dataclass
class Run:
    """A text run with inline formatting flags."""
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False  # inline code span → monospace font in DOCX
    color: str = ''  # hex color or named color (e.g. 'red', '#FF0000')
    ref_id: str = ''  # non-empty → section cross-reference; text holds @@SECREF:id@@ sentinel
    fig_ref_id: str = ''  # non-empty → figure cross-reference; text holds @@FIGREF:id@@ sentinel
    tab_ref_id: str = ''  # non-empty → table cross-reference; text holds @@TABREF:id@@ sentinel


@dataclass
class ContentNode:
    """A single content element in the document body.

    type values:
      'h1' 'h2' 'h3' 'h4'  — headings
      'p'                   — body paragraph
      'bullet'              — unordered list item
      'numbered'            — ordered (numbered) list item
      'figure'              — figure (image + caption)
      'table'               — data table
      'math_display'        — display equation ($$...$$)
    """
    type: str

    # Heading / paragraph / bullet
    runs: list[Run] = field(default_factory=list)
    level: int = 1  # nesting level for bullets

    # Figure fields
    fig_id: str = ''
    fig_number: int = 0         # assigned by DocumentContent.finalize()
    fig_path: str = ''          # filesystem path to PNG; '_mermaid:ID' for Mermaid-sourced
    fig_alt: str = ''
    fig_title: list[Run] = field(default_factory=list)
    fig_caption: list[Run] = field(default_factory=list)
    fig_width_hint: str = ''    # e.g. '50%', '10cm' — empty means full text width
    mermaid_src: str = ''       # raw Mermaid source (when fig_path starts with '_mermaid:')

    # Table fields
    tbl_id: str = ''
    tbl_number: int = 0         # assigned by DocumentContent.finalize()
    tbl_headers: list[str] = field(default_factory=list)
    tbl_rows: list[list[str]] = field(default_factory=list)
    tbl_caption: list[Run] = field(default_factory=list)

    # Math fields
    math_src: str = ''          # LaTeX source for 'math_display' nodes
    has_math: bool = False      # True for 'p' nodes containing $...$ inline math
    math_raw: str = ''          # Raw markdown of paragraph (for inline math post-processing)

    # Section cross-reference fields (headings only)
    sec_id: str = ''            # optional {#id} label from heading syntax
    sec_number: str = ''        # numeric prefix captured before stripping (e.g. '1.1')

    # Comment fields — set when a strict `> comment: Author` blockquote
    # immediately precedes this node in the draft. Renders as a native Word
    # (OOXML) comment attached to this node's text, not as body content.
    # Only h1-h4 / p / bullet / numbered nodes without inline math may carry
    # one; comment_author non-empty means "this node has a comment".
    comment_author: str = ''
    comment_body: list[list[Run]] = field(default_factory=list)  # one list per comment paragraph

    # Internal — raw text for macro processing pass
    _raw: str = ''


@dataclass
class DocumentContent:
    """Complete document content extracted from the Markdown draft.

    Deprecated: use (context_dict, body_var_name) tuple from parse_draft instead.
    Kept for backward compatibility; may be removed in a future release.
    """

    # Main body (all sections in reading order)
    content: list[ContentNode] = field(default_factory=list)


def finalize_content(content_nodes: list[ContentNode]) -> tuple:
    """Assign sequential numbers to figures and tables; build LOF/LOT lists.
    
    Returns:
        (figures_list, tables_list, fig_map, tab_map)
    """
    fig_num = 0
    tbl_num = 0
    figures_list: list[ContentNode] = []
    tables_list: list[ContentNode] = []
    fig_map: dict[str, int] = {}
    tab_map: dict[str, int] = {}
    
    for node in content_nodes:
        if node.type == 'figure':
            fig_num += 1
            node.fig_number = fig_num
            figures_list.append(node)
        elif node.type == 'table':
            tbl_num += 1
            node.tbl_number = tbl_num
            tables_list.append(node)
    
    for node in content_nodes:
        if node.type == 'figure' and node.fig_id:
            fig_map[node.fig_id] = node.fig_number
        elif node.type == 'table' and node.tbl_id:
            tab_map[node.tbl_id] = node.tbl_number
    
    return figures_list, tables_list, fig_map, tab_map
