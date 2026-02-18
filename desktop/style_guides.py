"""
Desktop Style Guides Configuration
----------------------------------
This module defines the formatting rules for the Desktop application.

MAINTENANCE WARNING:
    This file is currently a duplicate of the root `style_guides.py`.
    It exists here to keep the desktop application self-contained.
    Any changes made here must be mirrored in the root `style_guides.py`.
"""

from docx.shared import Pt, RGBColor, Inches

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_UNDERLINE
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from docx.enum.style import WD_STYLE_TYPE

STYLE_GUIDES = {
    "apa": {
        "inline_formatting": [],
        "reference_formatting": {
            "patterns": [
                {
                    "description": "Italicize journal titles and book titles in references",
                    "regex": r"\.\\s*([A-Z][A-Za-z0-9 &]+?[.?!])(?=\\s|$)",
                    "formatting": {"italic": True}
                },
                {
                    "description": "Format volume numbers in references",
                    "regex": r"(\\d+)\\s*\\(([^)]+)\\)",
                    "formatting": {"italic": False}
                }
            ]
        },
        "meta": {
            "title_page": True,
            "page_numbers": "header-right",
            "requires_abstract": True,
            "default_font": "Times New Roman",
            "citation_format": "({Author}, {Year})",
            "page_size": (Inches(8.5), Inches(11)),  # Letter size
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "all_caps": False, # APA 7 Title is Bold, Title Case, not All Caps
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "widow_control": True
                },
                "based_on": "Normal",
                "next_style": "Title Author"
            },
            "Title Byline": { #supports words like "by" on title pages
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Author"
            },
            "Title Author": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Department"
            },
            "Title Department": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Affiliation"
            },
            "Title Affiliation": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Course"
            },
            "Title Course": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Instructor"
            },
            "Title Instructor": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Date"
            },
            "Title Date": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Abstract"
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),  # No extra space before
                    "space_after": Pt(0.0),   # No extra space after
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 2": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 1 # Corrected: 1 for Heading 2
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 3": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 2 # Corrected: 2 for Heading 3
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },

            "Heading 4": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5), # Indented
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 3 # Corrected: 3 for Heading 4
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 5": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5), # Indented
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 4 # Corrected: 4 for Heading 5
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.0),  # Explicit left margin
                    "right_indent": Inches(0.0),  # Explicit right margin
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5),  # First line indent for paragraphs
                    "keep_together": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False # Normal paragraphs don't start on new pages
                },
                "base_style": None,
                "next_style": "Normal"
            },
            "Abstract": { # This style is for the *content* of the abstract, not the "Abstract" heading
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False, # APA abstract text is not italic
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0), # No first line indent for abstract
                    "keep_together": False,
                    "keep_with_next": True,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "page_break_before": False 
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Keywords": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5), # Indented like a paragraph
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False, # APA block quotes are not italic
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0), # No first line indent for block quotes
                    "keep_together": True,
                    "page_break_before": False # Block quotes don't start on new pages
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Caption": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "References": { # This style is for individual reference entries
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.5),
                    "first_line_indent": Inches(-0.5),  # Hanging indent for references
                    "keep_together": True,
                    "widow_control": True,
                    "orphan_control": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Figure": { # This style is for the "Figure X" label
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False, # Figures don't force new pages
                    "keep_together": True, # Keep figure label together
                    "keep_with_next": True, # Keep with the actual figure
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Figure Caption"
            },
            "Figure Caption": { # This style is for the caption text
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False, # Captions stay with their figures
                    "keep_together": True, # Keep caption text together
                    "keep_with_next": False, # Caption is last part of figure
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table": { # This style is for the "Table X" label
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False, # Tables don't force new pages by default
                    "keep_together": True, # Keep table label together
                    "keep_with_next": True, # Keep with the table
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Table Title"
            },
            "Table Title": { # This style is for the table title text
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False, # Table titles stay with their tables
                    "keep_together": True, # Keep title text together
                    "keep_with_next": True, # Keep with the table
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table Note": { # This style is for table notes
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT, # Often left-aligned
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE, # Often single-spaced
                    "page_break_before": False, # Table notes stay with their tables
                    "keep_together": True, # Keep note text together
                    "keep_with_next": False, # Note is last part of table
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Footnote Text": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10),
                    "bold": False,
                    "italic": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE, # APA footnotes are double-spaced
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "page_break_before": False, # Footnotes should not start on new pages
                    "keep_together": True, # Keep footnote text together
                    "keep_with_next": True, # Keep with the next footnote if any
                    "widow_control": True,
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal", # Footnote Text is usually a built-in style, but can be customized
                "next_style": "Normal"
            },

            "Appendix Title": {  # For "Appendix A", "Appendix B", etc. when no master "Appendices" exists
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,  # Heading 1 style (centered like Heading 1)
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,  # Each appendix starts on new page
                    "outline_level": 0  # Heading 1 level
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },

            "Appendices Title": {  # For the main "Appendices" heading (when it exists)
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,  # Heading 1 style (centered)
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,  # Appendices section starts on new page
                    "outline_level": 0  # Heading 1 level
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "List Bullet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Bullet"
            },
            "List Number": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Number"
            },
            "List Alphabet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Alphabet"
            },
            "List Item": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Item"
            },

            "Table Content": { # New style for content inside tables (0 indent)
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12), # Standard size
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0), # NO Indent
                    "first_line_indent": Inches(0), # NO Hanging indent
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE, # Single spacing is better for tables
                    "keep_together": False,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Table Content"
            },
            "Code Block": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Courier New",
                    "size": Pt(10),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(12.0), # Space before like block quote
                    "space_after": Pt(12.0), # Space after like block quote
                    "line_spacing": WD_LINE_SPACING.SINGLE, # Often single spaced for code
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            }
        }
    },
    
    "mla": {
        "inline_formatting": [],
        "reference_formatting": {
            "patterns": [
                {
                    "description": "Italicize book and journal titles in works cited",
                    "regex": r"\.\s*([A-Z][A-Za-z0-9 &]+?[.?!])(?=\s|$)",
                    "formatting": {"italic": True}
                },
                {
                    "description": "Format volume and issue numbers",
                    "regex": r"(\d+)\s*\(([^)]+)\)",
                    "formatting": {"italic": False}
                }
            ]
        },
        "meta": {
            "title_page": False,
            "page_numbers": "header-right",
            "requires_abstract": False,
            "default_font": "Times New Roman",
            "citation_format": "({Author} {Page})",
            "page_size": (Inches(8.5), Inches(11)),
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "all_caps": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "widow_control": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Title Byline": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Author"
            },
           "Title Author": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Department"
            },
            "Title Department": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Affiliation"
            },
            "Title Affiliation": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Course"
            },
            "Title Course": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Instructor"
            },
            "Title Instructor": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Date"
            },
            "Title Date": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Keywords": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                     "name": "Times New Roman",
                     "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.0)
                },
                "based_on": "Normal"
            },
            "Dedication": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                     "name": "Times New Roman",
                     "size": Pt(12),
                     "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(200.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": True
                },
                "based_on": "Normal"
            },
            "Epigraph": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(2.0),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(24.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                },
                "based_on": "Normal"
            },
            "Table Body": {
                 "type": WD_STYLE_TYPE.PARAGRAPH,
                 "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                 },
                 "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                 },
                 "based_on": "Normal"
            },
            "Table Note": { # MLA Table Note
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT, 
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE, 
                    "page_break_before": False, 
                    "keep_together": True, 
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 2": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 1
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 3": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 2
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 4": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 3
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 5": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 4
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5),
                    "keep_together": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "base_style": None,
                "next_style": "Normal"
            },
            "Abstract": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": False,
                    "keep_with_next": True,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendix Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendices Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "References": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.0),
                    "first_line_indent": Inches(-0.5),
                    "keep_together": True,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Figure": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Figure Caption"
            },
            "Figure Caption": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Table Title"
            },
            "Table Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table Note": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Footnote Text": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendix Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_with_next": True,
                    "page_break_before": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendices Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "List Bullet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Bullet"
            },
            "List Number": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Number"
            },
            "List Alphabet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Alphabet"
            },
            "List Item": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Item"
            },

        }
    },
    
    "chicago": {
        "inline_formatting": [],
        "reference_formatting": {
            "patterns": [
                {
                    "description": "Italicize journal titles and book titles in bibliography",
                    "regex": r"\.\s*([A-Z][A-Za-z0-9 &]+?[.?!])(?=\s|$)",
                    "formatting": {"italic": True}
                },
                {
                    "description": "Format volume numbers in bibliography",
                    "regex": r"(\d+)\s*\(([^)]+)\)",
                    "formatting": {"italic": False}
                }
            ]
        },
        "meta": {
            "title_page": True,
            "page_numbers": "header-right",
            "requires_abstract": False,
            "default_font": "Times New Roman",
            "citation_format": "({Author} {Year}, {Page})",
            "page_size": (Inches(8.5), Inches(11)),
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(18),
                    "bold": True,
                    "all_caps": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "widow_control": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Title Byline": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Author"
            },
           "Title Author": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Department"
            },
            "Title Department": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Affiliation"
            },
            "Title Affiliation": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Course"
            },
            "Title Course": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Instructor"
            },
            "Title Instructor": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Date"
            },
            "Title Date": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "page_break_after": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Keywords": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                     "name": "Times New Roman",
                     "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.0)
                },
                "based_on": "Normal"
            },
            "Dedication": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                     "name": "Times New Roman",
                     "size": Pt(12),
                     "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(200.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": True
                },
                "based_on": "Normal"
            },
            "Epigraph": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(2.0),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(24.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                },
                "based_on": "Normal"
            },
            "Table Body": {
                 "type": WD_STYLE_TYPE.PARAGRAPH,
                 "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                 },
                 "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                 },
                 "based_on": "Normal"
            },
            "Table Note": { 
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT, 
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE, 
                    "page_break_before": False, 
                    "keep_together": True, 
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 2": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 1
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 3": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 2
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 4": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 3
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 5": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 4
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5),
                    "keep_together": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "base_style": None,
                "next_style": "Normal"
            },
            "Abstract": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": False,
                    "keep_with_next": True,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "References": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "left_indent": Inches(0.5),
                    "first_line_indent": Inches(-0.5),
                    "keep_together": True,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Figure": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Figure Caption"
            },
            "Figure Caption": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Table Title"
            },
            "Table Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table Note": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Footnote Text": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendix Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_with_next": True,
                    "page_break_before": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendices Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "List Bullet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Bullet"
            },
            "List Number": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Number"
            },
            "List Alphabet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Alphabet"
            },
            "List Item": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Item"
            },

        }
    }
,
    "harvard": {
        "inline_formatting": [],
        "reference_formatting": {
            "patterns": [
                {
                    "description": "Italicize book and journal titles",
                    "regex": r"\.\s*([A-Z][A-Za-z0-9 &]+?[.?!])(?=\s|$)", 
                    "formatting": {"italic": True}
                }
            ]
        },
        "meta": {
            "title_page": True,
            "page_numbers": "header-right",
            "requires_abstract": False,
            "default_font": "Times New Roman",
            "citation_format": "({Author} {Year}, p. {Page})",
            "page_size": (Inches(8.5), Inches(11)),
            "orientation": WD_ORIENTATION.PORTRAIT,
            "default_tab_stops": Inches(0.5)
        },
        "margins": {
            "left": Inches(1),
            "right": Inches(1),
            "top": Inches(1),
            "bottom": Inches(1),
            "header": Inches(0.5),
            "footer": Inches(0.5),
            "gutter": Inches(0)
        },
        "styles": {
            "Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(18),
                    "bold": True,
                    "all_caps": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "widow_control": True
                },
                "based_on": "Normal",
                "next_style": "Title Author"
            },
            "Title Byline": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Author"
            },
            "Title Author": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Department"
            },
            "Title Department": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Affiliation"
            },
            "Title Affiliation": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Course"
            },
            "Title Course": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Instructor"
            },
            "Title Instructor": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True
                },
                "based_on": "Normal",
                "next_style": "Title Date"
            },
            "Title Date": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Abstract"
            },
            "Keywords": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0.5),
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE
                },
                "based_on": "Normal"
            },
            "Dedication": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": True
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(120),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": True
                },
                "based_on": "Normal"
            },
            "Epigraph": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(2.0),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(24.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                },
                "based_on": "Normal"
            },
            "Table Body": {
                 "type": WD_STYLE_TYPE.PARAGRAPH,
                 "font": {
                    "name": "Times New Roman",
                    "size": Pt(12)
                 },
                 "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "space_before": Pt(0),
                    "space_after": Pt(0),
                    "line_spacing": WD_LINE_SPACING.SINGLE
                 },
                 "based_on": "Normal"
            },
             "Table Note": { 
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10)
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT, 
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE, 
                    "page_break_before": False, 
                    "keep_together": True, 
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 1": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(24.0),
                    "space_after": Pt(24.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 2": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(24.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 1
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 3": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 2
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 4": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 3
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Heading 5": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": False,
                    "outline_level": 4
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Normal": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0.5),
                    "keep_together": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "base_style": None,
                "next_style": "Normal"
            },
            "Abstract": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": False,
                    "keep_with_next": True,
                    "left_indent": Inches(0.0),
                    "right_indent": Inches(0.0),
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Block Quote": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.5),
                    "right_indent": Inches(0.5),
                    "space_before": Pt(12.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "References": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "left_indent": Inches(0.5), # Hanging indent
                    "first_line_indent": Inches(-0.5),
                    "keep_together": True,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Figure": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Figure Caption"
            },
            "Figure Caption": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(0.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(12.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Table Title"
            },
            "Table Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(0.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Table Note": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "space_before": Pt(6.0),
                    "space_after": Pt(12.0),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "first_line_indent": Inches(0),
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Footnote Text": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(10),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "first_line_indent": Inches(0),
                    "line_spacing": WD_LINE_SPACING.SINGLE,
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "page_break_before": False,
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": True,
                    "left_indent": Inches(0),
                    "right_indent": Inches(0)
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendix Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_with_next": True,
                    "page_break_before": True
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "Appendices Title": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": True,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before": Pt(24.0),
                    "space_after": Pt(6.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "first_line_indent": Inches(0),
                    "keep_together": True,
                    "keep_with_next": True,
                    "page_break_before": True,
                    "outline_level": 0
                },
                "based_on": "Normal",
                "next_style": "Normal"
            },
            "List Bullet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Bullet"
            },
            "List Number": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Number"
            },
            "List Alphabet": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Alphabet"
            },
            "List Item": {
                "type": WD_STYLE_TYPE.PARAGRAPH,
                "font": {
                    "name": "Times New Roman",
                    "size": Pt(12),
                    "bold": False,
                    "italic": False,
                    "color": RGBColor(0, 0, 0),
                    "underline": WD_UNDERLINE.NONE
                },
                "paragraph": {
                    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    "left_indent": Inches(0.25),
                    "first_line_indent": Inches(-0.25),
                    "right_indent": Inches(0),
                    "space_before": Pt(0.0),
                    "space_after": Pt(0.0),
                    "line_spacing": WD_LINE_SPACING.DOUBLE,
                    "keep_together": True,
                    "keep_with_next": False,
                    "widow_control": True,
                    "orphan_control": True,
                    "page_break_before": False
                },
                "based_on": "Normal",
                "next_style": "List Item"
            },

        }
    }

}