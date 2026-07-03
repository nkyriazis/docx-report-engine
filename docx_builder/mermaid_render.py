"""Render Mermaid diagrams to PNG using shared mermaid_common module.

This uses the original figure IDs (e.g., 'architecture') so that FigureProxy
can find the PNGs by ID.

Usage:
    from docx_builder.mermaid_render import render_mermaid_figures
    render_mermaid_figures(content_nodes, output_dir='output/figures')
"""
from __future__ import annotations
import hashlib
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from .mermaid_common import render_mmd
from .schema import ContentNode


def _sha256(src: str) -> str:
    return hashlib.sha256(src.encode()).hexdigest()


def _fetch(fig_id: str, mermaid_src: str, out_path: Path) -> tuple[str, str | None, Exception | None]:
    via, err = render_mmd(mermaid_src, str(out_path))
    if via:
        return fig_id, via, None
    return fig_id, None, err


def render_mermaid_figures(
    content_nodes: list[ContentNode],
    output_dir: str = 'output/figures',
    cache_file: str | None = None,
    force: bool = False,
) -> dict[str, Path]:
    """Render all Mermaid figures in content_nodes to PNG files.

    Args:
        content_nodes:  List of ContentNode objects (from md_parser.parse_draft).
        output_dir:     Directory to write PNG files (created if needed).
        cache_file:     JSON file for content-hash caching (default: output_dir/.cache.json).
        force:          Re-render even if cached.

    Returns:
        dict mapping fig_id → Path of rendered PNG.
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    cache_path = Path(cache_file) if cache_file else out_dir / '.cache.json'

    try:
        cache: dict[str, str] = json.loads(cache_path.read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        cache = {}

    # Extract figure nodes
    figures = [n for n in content_nodes if n.type == 'figure']

    # Collect figures that need rendering
    to_render: list[tuple[str, str, Path]] = []
    for node in figures:
        if not node.fig_path.startswith('_mermaid:'):
            continue
        fig_id = node.fig_id
        src = node.mermaid_src
        if not src:
            continue
        digest = _sha256(src)
        out_path = out_dir / f'{fig_id}.png'
        if not force and cache.get(fig_id) == digest and out_path.exists():
            sz = out_path.stat().st_size // 1024
            print(f'  [{fig_id}] cached ({sz} KB)')
            continue
        to_render.append((fig_id, src, out_path))

    if not to_render:
        return {
            node.fig_id: out_dir / f'{node.fig_id}.png'
            for node in figures
            if node.fig_path.startswith('_mermaid:')
        }

    print(f'Rendering {len(to_render)} Mermaid diagram(s)...')
    results: dict[str, tuple[str | None, Exception | None]] = {}
    with ThreadPoolExecutor(max_workers=min(len(to_render), 4)) as pool:
        futs = {
            pool.submit(_fetch, fig_id, src, out_path): (fig_id, _sha256(src))
            for fig_id, src, out_path in to_render
        }
        for fut in as_completed(futs):
            fig_id, digest = futs[fut]
            _, via, err = fut.result()
            results[fig_id] = (via, err)
            if err:
                print(f'  [{fig_id}] FAILED: {err}')
            else:
                out_path = out_dir / f'{fig_id}.png'
                sz = out_path.stat().st_size // 1024
                print(f'  [{fig_id}] rendered ({sz} KB via {via})')
                cache[fig_id] = digest

    cache_path.write_text(json.dumps(cache, indent=2))

    rendered = {}
    for fig_id, src, out_path in to_render:
        if results.get(fig_id, (None, 'unknown'))[1] is None:
            rendered[fig_id] = out_path

    return rendered
