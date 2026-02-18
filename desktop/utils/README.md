# Formatly V3 Utilities

This package contains reusable helper modules for document processing, analysis, and formatting.

## Modules

### AutoCorrector

The `AutoCorrector` module handles spelling and grammar corrections for academic documents. It leverages the existing `DocumentChecker` class and adds automated correction capabilities.

Key features:
- Automatic spelling correction using best suggestions
- Grammar correction with rule-based and AI-powered approaches
- Comprehensive correction with detailed reporting
- Correction explanations for user feedback

Usage:
```python
from utils import AutoCorrector

# Initialize
corrector = AutoCorrector()

# Get correction report
report = corrector.get_correction_report(paragraphs)

# Apply automatic corrections
corrected_paragraphs, correction_summary = corrector.correct_document(
    paragraphs,
    auto_fix_spelling=True,
    auto_fix_grammar=True,
    use_ai_correction=True
)

# Generate explanation of corrections
explanation = corrector.generate_correction_explanation(correction_summary)

# Cleanup
corrector.close()
```

### FormattingAnalyzer

The `FormattingAnalyzer` module performs in-depth style compliance checks for academic documents. It analyzes citations, headings, references, margins, spacing, and other formatting elements.

Key features:
- Style guide compliance scoring
- Detailed issue detection and reporting
- Citation and reference format checking
- Formatting recommendations
- Analysis of document structure, margins, and spacing

Usage:
```python
from utils import FormattingAnalyzer

# Initialize with style
analyzer = FormattingAnalyzer(style_name="apa")

# Analyze document
analysis_result = analyzer.analyze_document("path/to/document.docx")

# Generate human-readable report
report = analyzer.generate_report(analysis_result)
print(report)
```

## Integration with app.py

These modules have been integrated with the main application:

1. `AutoCorrector` is used in the `--fix-errors` mode to automatically correct spelling and grammar issues
2. `FormattingAnalyzer` is used in the `--report-only` mode to provide detailed formatting analysis without modifying the document
