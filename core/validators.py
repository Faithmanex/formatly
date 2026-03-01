"""
Shared Validation Logic
----------------------
This module contains reusable functions for validating document formatting
against style guide requirements. These functions are used by both the 
FormattingAnalyzer and potentially the AdvancedFormatter.
"""

import re
from typing import List, Dict, Tuple, Optional
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

def validate_margins(doc: Document, style_guide: Dict) -> Dict:
    """
    Validate document margins against style guide requirements.
    
    Returns:
        Dict with 'compliant' boolean and 'details' list of issues.
    """
    requirements = style_guide.get("margins", {})
    issues = []
    
    for section in doc.sections:
        # Check top margin
        expected_top = requirements.get("top", Inches(1))
        if abs(section.top_margin.inches - expected_top.inches) > 0.01:
            issues.append(f"Top margin is {section.top_margin.inches:.2f}\", expected {expected_top.inches:.2f}\"")
            
        # Check bottom margin
        expected_bottom = requirements.get("bottom", Inches(1))
        if abs(section.bottom_margin.inches - expected_bottom.inches) > 0.01:
            issues.append(f"Bottom margin is {section.bottom_margin.inches:.2f}\", expected {expected_bottom.inches:.2f}\"")
            
        # Check left margin
        expected_left = requirements.get("left", Inches(1))
        if abs(section.left_margin.inches - expected_left.inches) > 0.01:
            issues.append(f"Left margin is {section.left_margin.inches:.2f}\", expected {expected_left.inches:.2f}\"")
            
        # Check right margin
        expected_right = requirements.get("right", Inches(1))
        if abs(section.right_margin.inches - expected_right.inches) > 0.01:
            issues.append(f"Right margin is {section.right_margin.inches:.2f}\", expected {expected_right.inches:.2f}\"")
            
    return {
        "compliant": len(issues) == 0,
        "details": issues
    }

def validate_fonts(doc: Document, style_guide: Dict) -> Dict:
    """
    Validate font face and size compliance.
    """
    meta = style_guide.get("meta", {})
    expected_font = meta.get("default_font", "Times New Roman")
    issues = []
    
    # Sample a set of paragraphs to avoid excessive processing time
    # Focus on non-empty paragraphs
    sample_size = 50
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    sample = paragraphs[:sample_size]
    
    for p in sample:
        # Check paragraph style font
        if p.style.font.name and p.style.font.name != expected_font:
            issues.append(f"Paragraph uses style '{p.style.name}' with font '{p.style.font.name}', expected '{expected_font}'")
            break # Just report once
            
        # Check run-level font (overrides style)
        for run in p.runs:
            if run.font.name and run.font.name != expected_font:
                issues.append(f"Text '{run.text[:20]}...' uses font '{run.font.name}', expected '{expected_font}'")
                break
        if issues: break
        
    return {
        "compliant": len(issues) == 0,
        "details": issues
    }

def validate_spacing(doc: Document, style_guide: Dict) -> Dict:
    """
    Validate line spacing compliance.
    """
    issues = []
    # Most academic styles require double spacing (WD_LINE_SPACING.DOUBLE)
    # We check the 'Normal' style or specific block types
    
    # APA/MLA/Chicago usually double space
    paragraphs = [p for p in doc.paragraphs if p.text.strip() and p.style.name == 'Normal']
    sample = paragraphs[:20]
    
    for p in sample:
        if p.paragraph_format.line_spacing_rule != WD_LINE_SPACING.DOUBLE:
            # Check if it's explicitly set to 2.0 (Double)
            if p.paragraph_format.line_spacing != 2.0:
                issues.append(f"Paragraph line spacing is not double spaced")
                break
                
    return {
        "compliant": len(issues) == 0,
        "details": issues
    }
