"""
Spacing removal utility for cleaning document paragraph spacing.
"""
from docx.shared import Pt
from docx.oxml.ns import qn


def remove_all_spacing(doc):
    """
    Remove all space_before and space_after from every paragraph in the document,
    including those within tables.
    """
    # 1. Process all body paragraphs
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

    # 2. Process all paragraphs inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
