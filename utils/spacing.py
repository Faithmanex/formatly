"""
Spacing removal utility for cleaning document paragraph spacing.
"""
from docx.shared import Pt
from docx.oxml.ns import qn


def remove_all_spacing(doc):
    """Remove all space_before and space_after from every paragraph in the document."""
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

        pPr = paragraph._p.get_or_add_pPr()
        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            pPr.remove(spacing)
