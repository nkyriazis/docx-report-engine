"""Allow ``python -m docx_builder`` as a build entry point."""
import clize

from .cli import build, load_prep_module
from .split_assemble import split, assemble
from .anchors import features

def instrument_template(
    *,
    template: str = "documents/template.docx",
    output: str = "template_jinja.docx",
    prep: str = "template_prep.py",
) -> None:
    """Instrument a DOCX template with Jinja2 markers via the project's prep script."""
    load_prep_module(prep).instrument_template(template, output)

clize.run({
    "build": build,
    "split": split,
    "assemble": assemble,
    "instrument-template": instrument_template,
    "features": features,
})
