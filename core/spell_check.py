"""
Spelling Checking Module
------------------------
This module provides spelling correction capabilities for Formatly.
It utilizes the `pyspellchecker` library to identify and suggest corrections
for misspelled words in document paragraphs.

Note:
    Grammar checking features are currently disabled or stubbed out,
    pending integration with a reliable grammar checking backend (e.g., LanguageTool or AI).
"""

import re
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from spellchecker import SpellChecker
# import language_tool_python
import google.generativeai as genai
import json
import os

# Configuration (Expected to be loaded by app entry point)
AI_API_KEY = os.getenv("GEMINI_API_KEY")
AI_MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-2.0-flash-exp")

@dataclass
class SpellError:
    """Represents a spelling error."""
    word: str
    position: int
    suggestions: List[str]
    paragraph_index: int

@dataclass
class GrammarError:
    """Represents a grammar error."""
    message: str
    position: int
    length: int
    suggestions: List[str]
    paragraph_index: int
    rule_id: str

class DocumentChecker:
    """Main class for checking spelling and grammar in documents."""

    def suggest_spelling_corrections(self, paragraphs: List[str]) -> Dict[str, str]:
        """
        Suggest spelling corrections for a list of paragraphs.
        print(f"[DEBUG] DocumentChecker.suggest_spelling_corrections: Number of paragraphs: {len(paragraphs)}")

        Args:
            paragraphs: List of paragraph texts

        Returns:
            A dictionary where keys are misspelled words and values are best suggestion
        """
        suggestions = {}
        for paragraph in paragraphs:
            print(f"[DEBUG] DocumentChecker.suggest_spelling_corrections: Checking paragraph: {paragraph[:30]}...")
            words_with_positions = self._extract_words_with_positions(paragraph)
            for word, _ in words_with_positions:
                if word.lower() not in self.spell_checker:
                    try:
                        # Get the best suggestion for the first candidate
                        suggestion = self.spell_checker.correction(word)
                        suggestions[word] = suggestion
                    except Exception:
                        continue
        return suggestions

    def suggest_grammar_corrections(self, paragraphs: List[str]) -> List[Tuple[int, int, int, str]]:
        """
        Suggest grammar corrections for a list of paragraphs.
        print(f"[DEBUG] DocumentChecker.suggest_grammar_corrections: Number of paragraphs: {len(paragraphs)}")

        Args:
            paragraphs: List of paragraph texts

        Returns:
            List of tuples indicating grammatical corrections
        """
        patches = []
        if not self.grammar_checker:
            print(f"[DEBUG] DocumentChecker.suggest_grammar_corrections: Grammar checker not configured")
            return patches
        for para_idx, paragraph in enumerate(paragraphs):
            print(f"[DEBUG] DocumentChecker.suggest_grammar_corrections: Checking paragraph {para_idx + 1}")
            matches = self.grammar_checker.check(paragraph)
            for match in matches:
                patches.append((
                    para_idx,
                    match.offset,
                    match.errorLength,
                    match.replacements[0] if match.replacements else ''
                ))
        return patches

    def apply_corrections(self, paragraphs: List[str]) -> Tuple[List[str], List[Tuple[str, str]]]:
        """
        Apply spelling and grammar corrections to paragraphs.
        print(f"[DEBUG] DocumentChecker.apply_corrections: Number of paragraphs: {len(paragraphs)}")

        Args:
            paragraphs: List of paragraph texts

        Returns:
            A tuple containing corrected paragraphs and a mapping of before → after
        """
        before_after_map = []
        # Correct spelling
        spelling_suggestions = self.suggest_spelling_corrections(paragraphs)
        paragraphs = [self.apply_spelling_corrections([p], spelling_suggestions)[0] for p in paragraphs]

        # Correct grammar
        grammar_patches = self.suggest_grammar_corrections(paragraphs)
        for para_idx, start, length, replacement in grammar_patches:
            paragraph = paragraphs[para_idx]
            before = paragraph[start:start+length]
            paragraphs[para_idx] = paragraph[:start] + replacement + paragraph[start+length:]
            before_after_map.append((before, replacement))

        return paragraphs, before_after_map

    def __init__(self, language: str = "en-US", english_variation: str = "american"):
        """
        Initialize the document checker.

        Args:
            language: Language code for checking (default: en-US)
            english_variation: Variation of English for spelling (american, british, australian, canadian)
        """
        self.language = language

        # Use the specified English variation for the spell checker
        # Note: pyspellchecker doesn't natively support different English variations
        # We'll use the standard spell checker but record the English variation for reference
        self.spell_checker = SpellChecker(language="en", distance=2)

        self.english_variation = english_variation.lower()
        print(f"Using {english_variation.capitalize()} English for spell checking")

        self.grammar_checker = None
        self._initialize_grammar_checker()

    def _initialize_grammar_checker(self):
        """Grammar checker initialization has been disabled."""
        self.grammar_checker = None

    def check_spelling(self, paragraphs: List[str]) -> List[SpellError]:
        """
        Check spelling in a list of paragraphs.

        Args:
            paragraphs: List of paragraph texts

        Returns:
            List of spelling errors found
        """
        errors = []

        for para_idx, paragraph in enumerate(paragraphs):
            # Extract words and their positions
            words_with_positions = self._extract_words_with_positions(paragraph)

            for word, position in words_with_positions:
                # Skip if word is in dictionary
                if word.lower() in self.spell_checker:
                    continue

                # Skip if word contains numbers or special characters
                if re.search(r'[0-9]', word) or len(word) < 2:
                    continue

                # Get suggestions (with timeout protection)
                try:
                    # Limit suggestion generation for performance
                    if len(word) > 15:  # Skip very long words
                        suggestions = []
                    else:
                        candidates = self.spell_checker.candidates(word)
                        suggestions = list(candidates)[:3] if candidates else []  # Limit to 3 for speed
                except (Exception, KeyboardInterrupt):
                    suggestions = []

                errors.append(SpellError(
                    word=word,
                    position=position,
                    suggestions=suggestions,
                    paragraph_index=para_idx
                ))

        return errors

    def check_grammar(self, paragraphs: List[str]) -> List[GrammarError]:
        """
        Grammar checking has been disabled.

        Args:
            paragraphs: List of paragraph texts

        Returns:
            Empty list as grammar checking is disabled
        """
        print("Note: Grammar checking has been disabled")
        return []

    def _check_grammar_java(self, paragraphs: List[str]) -> List[GrammarError]:
        """Check grammar using Java-based LanguageTool."""
        errors = []

        for para_idx, paragraph in enumerate(paragraphs):
            if not paragraph or not paragraph.strip():
                continue

            try:
                matches = self.grammar_checker.check(paragraph)

                for match in matches:
                    errors.append(GrammarError(
                        message=match.message,
                        position=match.offset,
                        length=match.errorLength,
                        suggestions=match.replacements[:3],  # Limit to 3 suggestions
                        paragraph_index=para_idx,
                        rule_id=match.ruleId
                    ))
            except Exception as e:
                print(f"Warning: Grammar check failed for paragraph {para_idx}: {e}")
                continue

        return errors

    def _check_grammar_ai(self, paragraphs: List[str]) -> List[GrammarError]:
        """Check grammar using AI (Gemini) as fallback."""
        if not AI_API_KEY:
            return []

        try:
            # Configure Gemini AI
            genai.configure(api_key=AI_API_KEY)
            model = genai.GenerativeModel(AI_MODEL_NAME)

            # Combine paragraphs for analysis (limit to reasonable size)
            text_sample = "\n\n".join(paragraphs[:10])  # Limit to first 10 paragraphs

            prompt = f"""
            You are a professional grammar checker for academic writing. Analyze the following text and identify grammar, punctuation, and style errors.

            Return your analysis as a JSON array of error objects with this format:
            [{{
                "message": "Brief description of the error",
                "suggestion": "Corrected version or suggestion",
                "paragraph_number": 0,
                "severity": "low|medium|high"
            }}]

            Text to analyze:

            {text_sample}

            Return only the JSON array, no additional text or markdown.
            """

            response = model.generate_content(prompt)
            return self._parse_ai_grammar_response(response.text, paragraphs)

        except Exception as e:
            print(f"Warning: AI grammar check failed: {e}")
            return []

    def _parse_ai_grammar_response(self, response_text: str, paragraphs: List[str]) -> List[GrammarError]:
        """Parse AI grammar response into GrammarError objects."""
        try:
            # Clean up response text
            text = response_text.strip()
            if text.startswith("```json"):
                text = text[7:]
            elif text.startswith("```"):
                text = text[3:]
            if text.endswith("```"):
                text = text[:-3]
            text = text.strip()

            # Parse JSON
            ai_errors = json.loads(text)

            if not isinstance(ai_errors, list):
                return []

            grammar_errors = []
            for idx, error in enumerate(ai_errors[:10]):  # Limit to 10 errors
                if isinstance(error, dict):
                    grammar_errors.append(GrammarError(
                        message=error.get("message", "Grammar error detected"),
                        position=0,  # AI doesn't provide exact position
                        length=1,
                        suggestions=[error.get("suggestion", "")] if error.get("suggestion") else [],
                        paragraph_index=min(error.get("paragraph_number", 0), len(paragraphs) - 1),
                        rule_id=f"AI_GRAMMAR_{idx}"
                    ))

            return grammar_errors

        except (json.JSONDecodeError, KeyError, ValueError) as e:
            print(f"Warning: Could not parse AI grammar response: {e}")
            return []

    def _extract_words_with_positions(self, text: str) -> List[Tuple[str, int]]:
        """
        Extract words and their positions from text.

        Args:
            text: Input text

        Returns:
            List of tuples (word, position)
        """
        words_with_positions = []

        # Find all words (letters only)
        for match in re.finditer(r'\b[a-zA-Z]+\b', text):
            word = match.group()
            position = match.start()
            words_with_positions.append((word, position))

        return words_with_positions

    def get_correction_report(self, paragraphs: List[str]) -> Dict:
        """
        Generate a comprehensive correction report.

        Args:
            paragraphs: List of paragraph texts

        Returns:
            Dictionary containing spelling and grammar errors
        """
        spelling_errors = self.check_spelling(paragraphs)
        grammar_errors = self.check_grammar(paragraphs)

        return {
            "spelling_errors": spelling_errors,
            "grammar_errors": grammar_errors,
            "total_spelling_errors": len(spelling_errors),
            "total_grammar_errors": len(grammar_errors),
            "paragraphs_checked": len(paragraphs),
            "overall_score": self._calculate_score(paragraphs, spelling_errors, grammar_errors)
        }


    def _calculate_score(self, paragraphs: List[str], spelling_errors: List[SpellError],
                        grammar_errors: List[GrammarError]) -> Dict:
        """
        Calculate an overall quality score for the document.

        Args:
            paragraphs: List of paragraph texts
            spelling_errors: List of spelling errors
            grammar_errors: List of grammar errors

        Returns:
            Dictionary with scoring information
        """
        total_words = sum(len(para.split()) for para in paragraphs)
        total_errors = len(spelling_errors) + len(grammar_errors)

        if total_words == 0:
            return {"score": 100, "grade": "A+", "total_words": 0}

        error_rate = (total_errors / total_words) * 100

        # Calculate score (100 - error percentage)
        score = max(0, 100 - (error_rate * 10))  # Amplify error impact

        # Assign grade
        if score >= 95:
            grade = "A+"
        elif score >= 90:
            grade = "A"
        elif score >= 85:
            grade = "B+"
        elif score >= 80:
            grade = "B"
        elif score >= 75:
            grade = "C+"
        elif score >= 70:
            grade = "C"
        elif score >= 65:
            grade = "D+"
        elif score >= 60:
            grade = "D"
        else:
            grade = "F"

        return {
            "score": round(score, 1),
            "grade": grade,
            "total_words": total_words,
            "error_rate": round(error_rate, 2)
        }

    def apply_spelling_corrections(self, paragraphs: List[str],
                                 corrections: Dict[str, str]) -> List[str]:
        """
        Apply spelling corrections to paragraphs.

        Args:
            paragraphs: List of paragraph texts
            corrections: Dictionary of {misspelled_word: correction}

        Returns:
            List of corrected paragraphs
        """
        corrected_paragraphs = []

        for paragraph in paragraphs:
            corrected_paragraph = paragraph

            # Apply corrections (case-sensitive)
            for misspelled, correction in corrections.items():
                # Use word boundaries to avoid partial matches
                pattern = r'\b' + re.escape(misspelled) + r'\b'
                corrected_paragraph = re.sub(pattern, correction, corrected_paragraph)

            corrected_paragraphs.append(corrected_paragraph)

        return corrected_paragraphs

    def close(self):
        """Clean up resources."""
        if self.grammar_checker:
            self.grammar_checker.close()

def format_error_report(report: Dict) -> str:
    """
    Format the error report for display.

    Args:
        report: Report dictionary from get_correction_report

    Returns:
        Formatted string report
    """
    lines = []
    lines.append("=" * 50)
    lines.append("  FORMATLY V3 - SPELLING REPORT")
    lines.append("=" * 50)
    lines.append("")

    # Overall score
    score_info = report["overall_score"]
    lines.append(f"📊 OVERALL SCORE: {score_info['score']}% ({score_info['grade']})")
    lines.append(f"📝 Total Words: {score_info['total_words']}")
    lines.append(f"⚠️  Error Rate: {score_info['error_rate']}%")
    lines.append("")

    # Spelling errors
    lines.append(f"🔤 SPELLING ERRORS: {report['total_spelling_errors']}")
    if report['spelling_errors']:
        lines.append("─" * 30)
        for error in report['spelling_errors'][:10]:  # Show first 10
            suggestions = ", ".join(error.suggestions[:3])
            lines.append(f"  '{error.word}' → Suggestions: {suggestions}")

        if len(report['spelling_errors']) > 10:
            lines.append(f"  ... and {len(report['spelling_errors']) - 10} more")
    lines.append("")

    # Grammar errors
    lines.append(f"📝 GRAMMAR ERRORS: {report['total_grammar_errors']}")
    if report['grammar_errors']:
        lines.append("─" * 30)
        for error in report['grammar_errors'][:10]:  # Show first 10
            suggestions = ", ".join(error.suggestions[:2])
            lines.append(f"  {error.message}")
            if suggestions:
                lines.append(f"     → Suggestions: {suggestions}")

        if len(report['grammar_errors']) > 10:
            lines.append(f"  ... and {len(report['grammar_errors']) - 10} more")

    lines.append("")
    lines.append("=" * 50)

    return "\n".join(lines)
