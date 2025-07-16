"""
Advanced Text Classifier for Formatly V6
Enhanced text classification logic with better pattern recognition and machine learning techniques.
"""

import re
from typing import List, Dict, Tuple, Optional, Any
from collections import defaultdict
import json
import google.generativeai as genai


class DocumentTextClassifier:
    """
    Advanced text classifier for academic documents with enhanced detection capabilities.
    """
    
    def __init__(self, style_name: str = "apa"):
        self.style_name = style_name.lower()
        self.patterns = self._initialize_patterns()
        self.section_keywords = self._initialize_section_keywords()
        self.confidence_threshold = 0.7
        
    def _initialize_patterns(self) -> Dict[str, List[str]]:
        """Initialize regex patterns for different document elements."""
        return {
            # Title page patterns
            "title_indicators": [
                r"^title:\s*(.+)$",
                r"^(.+)\s*$",  # First substantial line often title
                r"^[A-Z][^.!?]*[^.!?]$"  # Capitalized without ending punctuation
            ],
            
            # Author patterns
            "author_indicators": [
                r"^author:\s*(.+)$",
                r"^by\s+(.+)$",
                r"^([A-Z][a-z]+\s+[A-Z][a-z]+)$",  # First Last name pattern
                r"^([A-Z]\.\s*[A-Z][a-z]+)$"  # Initial Last name pattern
            ],
            
            # Institution patterns
            "institution_indicators": [
                r"^institution:\s*(.+)$",
                r"^university\s+of\s+(.+)$",
                r"^(.+)\s+university$",
                r"^(.+)\s+college$",
                r"^(.+)\s+institute$"
            ],
            
            # Heading patterns
            "heading_patterns": [
                r"^[A-Z][A-Z\s]*[A-Z]$",  # ALL CAPS headings
                r"^[0-9]+\.\s*[A-Z]",  # Numbered headings
                r"^[A-Z][a-z]+(\s+[A-Z][a-z]+)*$",  # Title Case
                r"^(Introduction|Literature Review|Methodology|Results|Discussion|Conclusion)$",
                r"^Chapter\s+[0-9IVX]+",  # Chapter headings
                r"^Section\s+[0-9]+",  # Section headings
            ],
            
            # Citation patterns
            "citation_patterns": [
                r"\([A-Za-z]+,\s*[0-9]{4}\)",  # APA style (Author, Year)
                r"\([A-Za-z]+\s+[0-9]+\)",  # MLA style (Author Page)
                r"\([A-Za-z]+\s+[0-9]{4},\s*p\.\s*[0-9]+\)",  # Full citation with page
                r"\[[0-9]+\]",  # Numbered citations
                r"(?:see|cf\.|e\.g\.|i\.e\.)\s*\([^)]+\)"  # Referential citations
            ],
            
            # Reference patterns
            "reference_patterns": [
                r"^[A-Za-z]+,\s*[A-Z]\.\s*\([0-9]{4}\)",  # APA reference start
                r"^[A-Za-z]+,\s*[A-Z][a-z]+\.\s*[\"']",  # MLA reference start
                r"^[A-Za-z]+,\s*[A-Z][a-z]+,\s*and\s*[A-Z]",  # Multi-author reference
                r"^\[[0-9]+\]",  # Numbered reference
                r"^[A-Za-z]+,\s*[A-Z]\.\s*[A-Z]\.",  # Multiple initials
            ],
            
            # Figure and table patterns
            "figure_patterns": [
                r"^Figure\s+[0-9]+",
                r"^Fig\.\s*[0-9]+",
                r"^Image\s+[0-9]+",
                r"^Chart\s+[0-9]+"
            ],
            
            "table_patterns": [
                r"^Table\s+[0-9]+",
                r"^Tbl\.\s*[0-9]+",
                r"^\s*\|.*\|.*\|",  # Table row pattern
                r"^[-+|=\s]+$"  # Table separator pattern
            ],
            
            # Block quote patterns
            "block_quote_patterns": [
                r"^\s{4,}",  # Indented text
                r"^>\s*",  # Markdown-style quote
                r"^[\"'].*[\"']$",  # Quoted text
                r"^As\s+[A-Za-z]+\s+stated",  # Attribution patterns
                r"^According\s+to\s+[A-Za-z]+"
            ],
            
            # Special sections
            "special_sections": [
                r"^Abstract\s*$",
                r"^Keywords?\s*:?",
                r"^Acknowledgments?\s*$",
                r"^References?\s*$",
                r"^Bibliography\s*$",
                r"^Appendix\s*[A-Z]?\s*$",
                r"^Glossary\s*$",
                r"^Index\s*$",
                r"^Table\s+of\s+Contents\s*$",
                r"^List\s+of\s+(Figures?|Tables?)\s*$"
            ]
        }
    
    def _initialize_section_keywords(self) -> Dict[str, List[str]]:
        """Initialize keywords for different section types."""
        return {
            "introduction": [
                "introduction", "overview", "background", "purpose", "objective",
                "scope", "rationale", "motivation", "problem statement",
                "aims", "goals", "intended contribution", "context"
            ],
            
            "literature_review": [
                "literature review", "review of literature", "literature survey",
                "related work", "previous research", "research background",
                "theoretical framework", "conceptual framework", "state of the art",
                "prior studies", "scholarly context"
            ],
            
            "methodology": [
                "methodology", "methods", "approach", "procedure", "experimental design",
                "research design", "study design", "data collection", "data analysis",
                "materials and methods", "experimental setup", "implementation details",
                "framework", "protocol", "technique"
            ],
            
            "results": [
                "results", "findings", "outcomes", "data", "observations",
                "measurements", "statistics", "performance", "evaluation",
                "data presentation", "empirical results", "statistical results",
                "experimental results"
            ],
            
            "discussion": [
                "discussion", "interpretation", "implications", "significance",
                "comparison", "evaluation", "assessment", "critique",
                "discussion of results", "key insights", "reflections", "synthesis",
                "analysis"
            ],
            
            "conclusion": [
                "conclusion", "summary", "concluding remarks", "final thoughts",
                "recommendations", "future work", "limitations", "closing",
                "closing remarks", "research implications", "key takeaways",
                "summary of findings"
            ],
            
            "abstract": [
                "abstract", "summary", "overview", "synopsis", "brief",
                "executive summary"
            ],
            
            "acknowledgments": [
                "acknowledgments", "acknowledgements", "thanks", "gratitude",
                "appreciation", "recognition", "dedication"
            ]
        }

    
    def classify_paragraph(self, text: str, context: List[str] = None) -> Dict[str, Any]:
        """
        Classify a single paragraph with confidence scoring.
        
        Args:
            text: The paragraph text to classify
            context: Surrounding paragraphs for context
            
        Returns:
            Dictionary with classification results and confidence scores
        """
        text = text.strip()
        if not text:
            return {"type": "empty", "confidence": 1.0, "metadata": {}}
        
        # Initialize classification results
        classification = {
            "type": "body_paragraph",
            "confidence": 0.0,
            "metadata": {},
            "alternative_types": []
        }
        
        # Check for special sections first (highest priority)
        special_result = self._classify_special_section(text)
        if special_result["confidence"] > self.confidence_threshold:
            classification.update(special_result)
            return classification
        
        # Check for headings
        heading_result = self._classify_heading(text, context)
        if heading_result["confidence"] > self.confidence_threshold:
            classification.update(heading_result)
            return classification
        
        # Check for citations and references
        citation_result = self._classify_citation_reference(text)
        if citation_result["confidence"] > self.confidence_threshold:
            classification.update(citation_result)
            return classification
        
        # Check for figures and tables
        figure_table_result = self._classify_figure_table(text)
        if figure_table_result["confidence"] > self.confidence_threshold:
            classification.update(figure_table_result)
            return classification
        
        # Check for block quotes
        quote_result = self._classify_block_quote(text)
        if quote_result["confidence"] > self.confidence_threshold:
            classification.update(quote_result)
            return classification
        
        # Check for title page elements
        title_result = self._classify_title_page_element(text, context)
        if title_result["confidence"] > self.confidence_threshold:
            classification.update(title_result)
            return classification
        
        # Default to body paragraph with section classification
        section_result = self._classify_body_section(text, context)
        classification.update(section_result)
        
        return classification
    
    def _classify_special_section(self, text: str) -> Dict[str, Any]:
        """Classify special sections like Abstract, References, etc."""
        text_lower = text.lower().strip()
        
        special_mappings = {
            "abstract": ["abstract"],
            "references": ["references", "bibliography", "works cited"],
            "acknowledgments": ["acknowledgments", "acknowledgements"],
            "appendix": ["appendix"],
            "glossary": ["glossary"],
            "index": ["index"],
            "table_of_contents": ["table of contents", "contents"],
            "list_of_figures": ["list of figures"],
            "list_of_tables": ["list of tables"]
        }
        
        for section_type, keywords in special_mappings.items():
            for keyword in keywords:
                if text_lower == keyword or text_lower == keyword + ":":
                    return {
                        "type": section_type,
                        "confidence": 0.95,
                        "metadata": {"exact_match": True}
                    }
        
        return {"type": "unknown", "confidence": 0.0, "metadata": {}}
    
    def _classify_heading(self, text: str, context: List[str] = None) -> Dict[str, Any]:
        """Classify heading levels and types."""
        confidence = 0.0
        level = 1
        metadata = {}
        
        # Check for numbered headings
        numbered_match = re.match(r"^(\d+\.)+\s*(.+)$", text)
        if numbered_match:
            level_count = numbered_match.group(1).count('.')
            level = min(level_count, 5)
            confidence = 0.9
            metadata["numbered"] = True
            metadata["number"] = numbered_match.group(1)
            metadata["title"] = numbered_match.group(2)
        
        # Check for ALL CAPS headings
        elif text.isupper() and len(text.split()) <= 8:
            confidence = 0.8
            level = 1
            metadata["style"] = "all_caps"
        
        # Check for title case headings
        elif text.istitle() and not text.endswith('.') and len(text.split()) <= 10:
            confidence = 0.7
            level = 2
            metadata["style"] = "title_case"
        
        # Check for common section headings
        common_headings = {
            "introduction": 1,
            "literature review": 1,
            "methodology": 1,
            "results": 1,
            "discussion": 1,
            "conclusion": 1,
            "abstract": 1,
            "references": 1
        }
        
        text_lower = text.lower().strip()
        for heading_text, heading_level in common_headings.items():
            if text_lower == heading_text:
                confidence = 0.95
                level = heading_level
                metadata["common_section"] = True
                break
        
        if confidence > 0:
            return {
                "type": f"heading_{level}",
                "confidence": confidence,
                "metadata": metadata
            }
        
        return {"type": "unknown", "confidence": 0.0, "metadata": {}}
    
    def _classify_citation_reference(self, text: str) -> Dict[str, Any]:
        """Classify citations and references."""
        # Check for reference list entry
        reference_patterns = [
            r"^[A-Za-z]+,\s*[A-Z]\.\s*\([0-9]{4}\)",  # APA style
            r"^[A-Za-z]+,\s*[A-Z][a-z]+\.\s*[\"']",  # MLA style
            r"^\[[0-9]+\]"  # Numbered reference
        ]
        
        for pattern in reference_patterns:
            if re.match(pattern, text):
                return {
                    "type": "reference_entry",
                    "confidence": 0.85,
                    "metadata": {"pattern": pattern}
                }
        
        # Check for in-text citations
        citation_count = len(re.findall(r'\([A-Za-z]+,?\s*[0-9]{4}\)', text))
        if citation_count > 0:
            return {
                "type": "body_paragraph",
                "confidence": 0.6,
                "metadata": {"has_citations": True, "citation_count": citation_count}
            }
        
        return {"type": "unknown", "confidence": 0.0, "metadata": {}}
    
    def _classify_figure_table(self, text: str) -> Dict[str, Any]:
        """Classify figures and tables."""
        # Figure patterns
        figure_patterns = [
            r"^Figure\s+[0-9]+",
            r"^Fig\.\s*[0-9]+",
            r"^Image\s+[0-9]+",
            r"^Chart\s+[0-9]+"
        ]
        
        for pattern in figure_patterns:
            match = re.match(pattern, text, re.IGNORECASE)
            if match:
                return {
                    "type": "figure_caption",
                    "confidence": 0.9,
                    "metadata": {"pattern": pattern, "number": match.group()}
                }
        
        # Table patterns
        table_patterns = [
            r"^Table\s+[0-9]+",
            r"^Tbl\.\s*[0-9]+"
        ]
        
        for pattern in table_patterns:
            match = re.match(pattern, text, re.IGNORECASE)
            if match:
                return {
                    "type": "table_caption",
                    "confidence": 0.9,
                    "metadata": {"pattern": pattern, "number": match.group()}
                }
        
        # Check for table content (rows with separators)
        if re.match(r"^\s*\|.*\|.*\|", text) or re.match(r"^[-+|=\s]+$", text):
            return {
                "type": "table_content",
                "confidence": 0.8,
                "metadata": {"table_row": True}
            }
        
        return {"type": "unknown", "confidence": 0.0, "metadata": {}}
    
    def _classify_block_quote(self, text: str) -> Dict[str, Any]:
        """Classify block quotes."""
        confidence = 0.0
        metadata = {}
        
        # Check for indented text
        if text.startswith("    ") or text.startswith("\t"):
            confidence = 0.6
            metadata["indented"] = True
        
        # Check for quoted text
        if (text.startswith('"') and text.endswith('"')) or \
           (text.startswith("'") and text.endswith("'")):
            confidence = 0.7
            metadata["quoted"] = True
        
        # Check for attribution patterns
        attribution_patterns = [
            r"^As\s+[A-Za-z]+\s+stated",
            r"^According\s+to\s+[A-Za-z]+",
            r"^[A-Za-z]+\s+argues\s+that",
            r"^[A-Za-z]+\s+suggests\s+that"
        ]
        
        for pattern in attribution_patterns:
            if re.match(pattern, text):
                confidence = 0.8
                metadata["has_attribution"] = True
                break
        
        if confidence > 0:
            return {
                "type": "block_quote",
                "confidence": confidence,
                "metadata": metadata
            }
        
        return {"type": "unknown", "confidence": 0.0, "metadata": {}}
    
    def _classify_title_page_element(self, text: str, context: List[str] = None) -> Dict[str, Any]:
        """Classify title page elements."""
        if not context:
            context = []
        
        # Check if we're likely in title page context (first few paragraphs)
        position = len(context)
        if position > 10:  # Unlikely to be title page after 10 paragraphs
            return {"type": "unknown", "confidence": 0.0, "metadata": {}}
        
        # Title patterns
        if position <= 3:  # First few lines are likely title
            if len(text.split()) <= 15 and not text.endswith('.'):
                return {
                    "type": "title",
                    "confidence": 0.8,
                    "metadata": {"position": position}
                }
        
        # Author patterns
        author_patterns = [
            r"^by\s+(.+)$",
            r"^([A-Z][a-z]+\s+[A-Z][a-z]+)$",
            r"^([A-Z]\.\s*[A-Z][a-z]+)$"
        ]
        
        for pattern in author_patterns:
            if re.match(pattern, text):
                return {
                    "type": "author",
                    "confidence": 0.85,
                    "metadata": {"pattern": pattern}
                }
        
        # Institution patterns
        institution_keywords = ["university", "college", "institute", "school", "department"]
        text_lower = text.lower()
        for keyword in institution_keywords:
            if keyword in text_lower:
                return {
                    "type": "institution",
                    "confidence": 0.7,
                    "metadata": {"keyword": keyword}
                }
        
        return {"type": "unknown", "confidence": 0.0, "metadata": {}}
    
    def _classify_body_section(self, text: str, context: List[str] = None) -> Dict[str, Any]:
        """Classify body paragraph sections."""
        text_lower = text.lower()
        
        # Check for section-specific keywords
        section_scores = {}
        
        for section_type, keywords in self.section_keywords.items():
            score = 0
            for keyword in keywords:
                if keyword in text_lower:
                    score += 1
            
            if score > 0:
                section_scores[section_type] = score / len(keywords)
        
        if section_scores:
            best_section = max(section_scores, key=section_scores.get)
            confidence = min(section_scores[best_section] * 2, 0.8)  # Cap at 0.8
            
            return {
                "type": "body_paragraph",
                "confidence": confidence,
                "metadata": {
                    "section_type": best_section,
                    "section_confidence": confidence,
                    "keyword_matches": section_scores
                }
            }
        
        return {
            "type": "body_paragraph",
            "confidence": 0.5,
            "metadata": {"section_type": "general"}
        }
    
    def classify_document(self, paragraphs: List[str]) -> Dict[str, Any]:
        """
        Classify an entire document's structure.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Structured document classification
        """
        classifications = []
        
        # Classify each paragraph with context
        for i, paragraph in enumerate(paragraphs):
            context = paragraphs[max(0, i-2):i]  # Previous 2 paragraphs for context
            classification = self.classify_paragraph(paragraph, context)
            classification["index"] = i
            classification["text"] = paragraph
            classifications.append(classification)
        
        # Build structured document
        return self._build_document_structure(classifications)
    
    def _build_document_structure(self, classifications: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Build structured document from classifications."""
        structure = {
            "title_page": {},
            "front_matter": {},
            "abstract": {"text": "", "keywords": []},
            "headings": {
                "heading_1": [],
                "heading_2": [],
                "heading_3": [],
                "heading_4": [],
                "heading_5": []
            },
            "body_content": {
                "introduction": [],
                "methodology": [],
                "results": [],
                "discussion": [],
                "conclusion": [],
                "body_paragraphs": []
            },
            "special_elements": {
                "block_quotes": [],
                "equations": [],
                "code_blocks": [],
                "lists": []
            },
            "citations_and_references": {
                "in_text_citations": [],
                "footnotes": [],
                "endnotes": [],
                "references": []
            },
            "tables_and_figures": {
                "tables": [],
                "figures": [],
                "charts": [],
                "images": []
            },
            "appendices": {
                "appendix_sections": [],
                "supplementary_materials": []
            },
            "back_matter": {
                "bibliography": [],
                "index": [],
                "author_bio": ""
            }
        }
        
        current_section = None
        current_heading = None
        
        for classification in classifications:
            cls_type = classification["type"]
            text = classification["text"]
            metadata = classification.get("metadata", {})
            
            # Handle different classification types
            if cls_type == "title":
                structure["title_page"]["title"] = text
            elif cls_type == "author":
                structure["title_page"]["author"] = text
            elif cls_type == "institution":
                structure["title_page"]["institution"] = text
            elif cls_type == "abstract":
                current_section = "abstract"
            elif cls_type.startswith("heading_"):
                level = cls_type.split("_")[1]
                heading_entry = {"text": text, "content": []}
                structure["headings"][f"heading_{level}"].append(heading_entry)
                current_heading = heading_entry
                # Update current section based on heading
                text_lower = text.lower()
                if any(keyword in text_lower for keyword in ["introduction", "intro"]):
                    current_section = "introduction"
                elif any(keyword in text_lower for keyword in ["method", "approach"]):
                    current_section = "methodology"
                elif "result" in text_lower:
                    current_section = "results"
                elif "discussion" in text_lower:
                    current_section = "discussion"
                elif "conclusion" in text_lower:
                    current_section = "conclusion"
            elif cls_type == "body_paragraph":
                section_type = metadata.get("section_type", "body_paragraphs")
                if current_section and current_section in structure["body_content"]:
                    structure["body_content"][current_section].append(text)
                elif section_type in structure["body_content"]:
                    structure["body_content"][section_type].append(text)
                else:
                    structure["body_content"]["body_paragraphs"].append(text)
                
                # Add to current heading content if available
                if current_heading:
                    current_heading["content"].append(text)
            elif cls_type == "block_quote":
                quote_entry = {"quote": text, "source": "", "page": ""}
                structure["special_elements"]["block_quotes"].append(quote_entry)
            elif cls_type == "reference_entry":
                ref_entry = {"entry": text, "type": "unknown"}
                structure["citations_and_references"]["references"].append(ref_entry)
            elif cls_type == "figure_caption":
                figure_entry = {"caption": text, "content": "", "source": ""}
                structure["tables_and_figures"]["figures"].append(figure_entry)
            elif cls_type == "table_caption":
                table_entry = {"title": text, "content": "", "notes": []}
                structure["tables_and_figures"]["tables"].append(table_entry)
            elif cls_type == "references":
                current_section = "references"
            elif cls_type == "acknowledgments":
                structure["front_matter"]["acknowledgments"] = text
            
            # Handle abstract content
            if current_section == "abstract" and cls_type == "body_paragraph":
                if structure["abstract"]["text"]:
                    structure["abstract"]["text"] += "\n" + text
                else:
                    structure["abstract"]["text"] = text
        
        return structure
    
    def enhance_with_ai(self, paragraphs: List[str], api_key: str) -> Dict[str, Any]:
        """
        Enhance classification using AI model.
        
        Args:
            paragraphs: List of paragraph texts
            api_key: Gemini API key
            
        Returns:
            Enhanced document structure
        """
        try:
            # Get base classification
            base_structure = self.classify_document(paragraphs)
            
            # Use AI to refine and enhance
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-2.0-flash-lite")
            
            # Create enhanced prompt
            prompt = self._create_enhancement_prompt(paragraphs, base_structure)
            response = model.generate_content(prompt)
            
            # Parse AI response and merge with base structure
            try:
                ai_structure = json.loads(response.text.strip())
                return self._merge_structures(base_structure, ai_structure)
            except json.JSONDecodeError:
                print("Warning: AI response could not be parsed, using base classification")
                return base_structure
                
        except Exception as e:
            print(f"Warning: AI enhancement failed: {e}")
            return self.classify_document(paragraphs)
    
    def _create_enhancement_prompt(self, paragraphs: List[str], base_structure: Dict[str, Any]) -> str:
        """Create prompt for AI enhancement."""
        return f"""
        You are an expert academic document analyzer. Review the following document paragraphs and the base classification, then provide an enhanced JSON structure.

        Base Classification:
        {json.dumps(base_structure, indent=2)}

        Document Paragraphs:
        {chr(10).join(f"{i+1}. {para}" for i, para in enumerate(paragraphs))}

        Instructions:
        1. Refine the classification of each paragraph
        2. Improve section detection and grouping
        3. Better identify citations, references, figures, and tables
        4. Enhance metadata for each element
        5. Return only valid JSON matching the base structure format

        Enhanced JSON Structure:
        """
    
    def _merge_structures(self, base: Dict[str, Any], enhanced: Dict[str, Any]) -> Dict[str, Any]:
        """Merge base and AI-enhanced structures."""
        # For now, prefer AI structure but fall back to base for missing elements
        merged = base.copy()
        
        for key, value in enhanced.items():
            if key in merged and value:
                if isinstance(value, dict):
                    merged[key].update(value)
                elif isinstance(value, list) and value:
                    merged[key] = value
                else:
                    merged[key] = value
        
        return merged
