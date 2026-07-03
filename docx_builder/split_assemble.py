"""Split and assemble deliverable draft into/from section files."""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

DRAFT = Path("draft/draft.md")
SECTIONS = Path("draft/sections")


def slugify(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-").replace("--", "-")


def _get_anchor(line: str) -> str | None:
    m = re.search(r"\{#([^}]+)\}", line)
    return m.group(1) if m else None


def split(*, draft: str = str(DRAFT), out: str = str(SECTIONS)) -> None:
    """Split draft into section files in draft/sections/."""
    draft_path = Path(draft)
    if not draft_path.exists():
        print(f"not found: {draft_path}", file=sys.stderr)
        sys.exit(1)

    text = draft_path.read_text()
    lines = text.splitlines(True)

    h = hashlib.sha256(text.encode()).hexdigest()
    out_dir = Path(out)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / ".draft_hash").write_text(h)

    heading_indices = [i for i, line in enumerate(lines) if line.startswith("## ")]
    boundaries = [0] + heading_indices + [len(lines)]
    index = []

    for seq in range(len(heading_indices) + 1):
        start, end = boundaries[seq], boundaries[seq + 1]
        chunk = lines[start:end]

        if seq == 0:
            heading, anchor, fname = "Frontmatter", "frontmatter", "00_frontmatter.md"
        else:
            hline = chunk[0].rstrip("\n\r")
            heading = re.sub(r"^##\s*", "", hline).split("{")[0].strip()
            anchor = _get_anchor(hline) or slugify(heading)
            fname = f"{seq:02d}_{anchor}.md"

        (out_dir / fname).write_text("".join(chunk))
        index.append({"file": fname, "heading": heading, "anchor": anchor,
                       "line_start": start + 1, "line_end": end})

    (out_dir / "_index.json").write_text(json.dumps(index, indent=2) + "\n")
    print(f"split {len(index)} sections -> {out_dir}/ (.draft_hash={h[:8]})")


def assemble(*, out: str = str(DRAFT), sections: str = str(SECTIONS), check: bool = False) -> None:
    """Assemble section files back into single draft.

    With --check, verify SHA-256 hash matches stored .draft_hash.
    """
    out_dir = Path(sections)
    index_file = out_dir / "_index.json"

    if index_file.exists():
        files = [e["file"] for e in json.loads(index_file.read_text())]
    else:
        files = sorted(f.name for f in out_dir.glob("*.md") if not f.name.startswith("."))

    if not files:
        print("no section files found", file=sys.stderr)
        sys.exit(1)

    assembled = "".join((out_dir / f).read_text() for f in files)
    out_path = Path(out)
    out_path.write_text(assembled)

    if check:
        hash_file = out_dir / ".draft_hash"
        if not hash_file.exists():
            print("no .draft_hash found (run split first)", file=sys.stderr)
            sys.exit(1)
        stored = hash_file.read_text().strip()
        actual = hashlib.sha256(assembled.encode()).hexdigest()
        if stored == actual:
            print(f"check passed: {actual[:8]}")
        else:
            print(f"check failed: stored={stored[:8]} actual={actual[:8]}", file=sys.stderr)
            sys.exit(1)
    else:
        print(f"assembled {len(files)} sections -> {out_path}")
