"""Pick the packing-list extractor that fits the PDF, instead of asking.

Two real, differently-laid-out packing-list templates exist - the
original custom one emailed as an attachment (pdf_parser) and one
rendered by Odoo's own report engine (odoo_pdf_parser). Callers used to
choose between them: the GUI put a dropdown in front of the user and the
CLI took an --odoo flag. Picking wrong doesn't produce a helpful message,
it produces the *other* parser's token-dump failure, which reads like the
PDF is broken rather than like the wrong extractor was chosen.

proforma_parser already settled this for pro-forma invoices - it reads
both of its layouts and picks by what it finds in the file - and this
does the same for packing lists, by the simplest signal available: which
extractor actually parses the file. That signal is exact here because
each template's item rows stream in an order the other parser cannot
read; verified against every sample PDF, each one is parsed by exactly
one of the two.

A misdetection that somehow parsed anyway would still be caught
downstream: generate_pl1.validate() reconciles the extracted totals
against the PDF's own Grand Total row, and both callers refuse to
generate when that fails.
"""

from odoo_pdf_parser import extract_pdf as extract_odoo
from pdf_parser import ExtractionError
from pdf_parser import extract_pdf as extract_email

TEMPLATES = (
    ("Email PDF", extract_email),
    ("Odoo PDF", extract_odoo),
)


def extract_pdf(pdf_path, template=None):
    """Return (header, items, grand_total, template_name).

    `template` forces one of the names in TEMPLATES, for a caller that
    already knows; left None, each is tried in turn and the first that
    parses wins.
    """
    if template is not None:
        chosen = dict(TEMPLATES).get(template)
        if chosen is None:
            raise ExtractionError(f"Unknown packing list template: {template!r}")
        return (*chosen(pdf_path), template)

    errors = []
    for name, extract in TEMPLATES:
        try:
            return (*extract(pdf_path), name)
        except ExtractionError as e:
            errors.append(f"  as {name}: {e}")

    raise ExtractionError(
        "This PDF doesn't match either known packing list template.\n"
        + "\n".join(errors)
    )
