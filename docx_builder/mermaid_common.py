"""Shared mermaid rendering — single mmdc call, used by both build.py and docx_builder.

Render a mermaid source string to a PNG file using local mmdc with --scale 4.
"""
from __future__ import annotations
import subprocess
import shutil
from pathlib import Path


MMDC_SCALE = 4  # 4× for crisp output


def render_mmd(mermaid_src: str, png_path: str) -> tuple[str, str | None, Exception | None]:
    """Render mermaid_src to png_path using local mmdc.

    Returns (via, error) where via is 'mmdc' on success or None on failure.
    """
    mmdc = shutil.which("mmdc")
    if not mmdc:
        return None, Exception("mmdc not found — local mermaid-cli required")

    png = Path(png_path)
    try:
        tmp = png.with_suffix(".mmd")
        tmp.write_text(mermaid_src)
        pptr_cfg = png.with_suffix(".puppeteer.json")
        pptr_cfg.write_text('{"args":["--no-sandbox","--disable-setuid-sandbox"]}')
        subprocess.run(
            [mmdc, "-p", str(pptr_cfg), "-s", str(MMDC_SCALE),
             "-i", str(tmp), "-o", str(png)],
            capture_output=True, timeout=60, check=True,
        )
        pptr_cfg.unlink(missing_ok=True)
        tmp.unlink(missing_ok=True)
        if png.exists() and png.stat().st_size > 1000:
            return "mmdc", None
    except Exception as e:
        print(f"   mmdc fail: {e}")
    return None, Exception("mmdc rendering failed")
