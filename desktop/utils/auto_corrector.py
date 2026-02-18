"""
AutoCorrector Module for Formatly V3
Handles spelling and grammar correction for academic documents.
Leverages DocumentChecker and implements advanced grammar fix logic.
"""

import re
from typing import List, Dict, Tuple, Optional
import google.generativeai as genai
import os
from dotenv import load_dotenv
from spell_check import DocumentChecker, SpellError, GrammarError, format_error_report

# Load environment variables for AI grammar correction
load_dotenv()
AI_API_KEY = os.getenv("GEMINI_API_KEY")
AI_MODEL_NAME = os.getenv("GEMINI_MODEL")

class AutoCorrector:
    """
    Handles automatic correction of spelling and grammar errors in documents.
    Extends DocumentChecker functionality with automatic correction capabilities.
    """
    
    def __init__(self, language: str = "en-US", english_variation: str = "american"):
        """
        Initialize the auto corrector.
        
        Args:
            language: Language code for checking (default: en-US)
            english_variation: Variation of English for spelling (american, british, australian, canadian)
        """
        self.language = language
        self.english_variation = english_variation
        # Keep document_checker for backward compatibility
        self.document_checker = DocumentChecker(language=language, english_variation=english_variation)
    
    def _initialize_ai_model(self):
        """Initialize the AI model for advanced grammar correction."""
        try:
            genai.configure(api_key=AI_API_KEY)
            self.ai_model = genai.GenerativeModel(AI_MODEL_NAME)
        except Exception as e:
            print(f"Warning: Could not initialize AI model: {e}")
            self.ai_model = None
    
    def get_correction_report(self, paragraphs: List[str]) -> Dict:
        """
        Generate a comprehensive correction report.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Dictionary containing spelling and grammar errors
        """
        return self.document_checker.get_correction_report(paragraphs)
    
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
        return self.document_checker.apply_spelling_corrections(paragraphs, corrections)
    
    def auto_correct_spelling(self, paragraphs: List[str]) -> Tuple[List[str], Dict]:
        """
        Automatically correct spelling errors using the most likely suggestion.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Tuple of (corrected_paragraphs, corrections_applied)
        """
        # Get spelling errors
        spelling_errors = self.document_checker.check_spelling(paragraphs)
        
        # Create corrections dictionary with the top suggestion for each error
        corrections = {}
        for error in spelling_errors:
            if error.suggestions:
                corrections[error.word] = error.suggestions[0]  # Use the top suggestion
        
        # Apply corrections
        corrected_paragraphs = self.apply_spelling_corrections(paragraphs, corrections)
        
        return corrected_paragraphs, corrections
    
    def auto_correct_grammar(self, paragraphs: List[str]) -> Tuple[List[str], Dict]:
        """
        Grammar correction has been disabled.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Tuple of (paragraphs, empty_corrections)
        """
        print("Note: Grammar correction has been disabled")
        return paragraphs.copy(), {}
    
    def apply_ai_grammar_correction(self, paragraphs: List[str]) -> Tuple[List[str], Dict]:
        """
        AI-based grammar correction has been disabled.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Tuple of (paragraphs, empty_corrections)
        """
        print("Note: AI-based grammar correction has been disabled")
        return paragraphs.copy(), {}
    
    def _parse_ai_correction_response(self, response_text: str) -> Dict:
        """Parse AI correction response into structured format."""
        import json
        
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
            return json.loads(text)
            
        except (json.JSONDecodeError, Exception) as e:
            print(f"Warning: Could not parse AI correction response: {e}")
            return {}
    
    def correct_document(self, paragraphs: List[str], 
                         auto_fix_spelling: bool = True,
                         auto_fix_grammar: bool = False,  # Changed default to False
                         use_ai_correction: bool = True,
                         chunk_size: int = 5) -> Tuple[List[str], Dict]:
        """
        Apply comprehensive correction to a document.
        Uses Gemini AI for correction if available, otherwise falls back to basic spell checker.
        
        Args:
            paragraphs: List of paragraph texts
            auto_fix_spelling: Whether to auto-fix spelling errors
            auto_fix_grammar: Whether to fix grammar (now via Gemini AI)
            use_ai_correction: Whether to use Gemini AI for correction
            
        Returns:
            Tuple of (corrected_paragraphs, correction_summary)
        """
        corrected_paragraphs = paragraphs.copy()
        correction_summary = {
            "spelling": {},
            "grammar": {},
            "capitalization": {},
            "punctuation": {},
            "ai_grammar": {},  # Add missing ai_grammar key
            "total_corrections": 0
        }
        
        # AI correction has been removed
        if use_ai_correction:
            print("Note: AI correction has been disabled")
        
        # Fall back to standard spell checking if Gemini is not available or failed
        if auto_fix_spelling:
            corrected_paragraphs, spelling_corrections = self.auto_correct_spelling(corrected_paragraphs)
            correction_summary["spelling"] = spelling_corrections
            correction_summary["total_corrections"] += len(spelling_corrections)
        
        # Grammar correction via old method is disabled
        if auto_fix_grammar and not self.gemini_corrector:
            print("Note: Grammar correction requires Gemini AI which is not available.")
        
        return corrected_paragraphs, correction_summary
    
    def generate_correction_explanation(self, correction_summary: Dict) -> str:
        """
        Generate a human-readable explanation of corrections applied.
        
        Args:
            correction_summary: Dictionary of corrections from correct_document()
            
        Returns:
            Formatted string explaining corrections
        """
        lines = []
        lines.append("=== CORRECTION SUMMARY ===")
        lines.append(f"Total corrections applied: {correction_summary.get('total_corrections', 0)}")
        
        # Use safe dictionary access with .get() to avoid KeyError
        if correction_summary.get("spelling", {}):
            lines.append("\nSPELLING CORRECTIONS:")
            for original, corrected in correction_summary["spelling"].items():
                lines.append(f"  '{original}' → '{corrected}'")
        
        if correction_summary.get("grammar", {}):
            lines.append("\nGRAMMAR CORRECTIONS:")
            for original, corrected in correction_summary["grammar"].items():
                lines.append(f"  '{original}' → '{corrected}'")
        
        if correction_summary.get("ai_grammar", {}):
            lines.append("\nADVANCED AI CORRECTIONS:")
            for original, corrected in correction_summary["ai_grammar"].items():
                lines.append(f"  '{original}' → '{corrected}'")
                
        # Add punctuation and capitalization to report if available
        if correction_summary.get("punctuation", {}):
            lines.append("\nPUNCTUATION CORRECTIONS:")
            for original, corrected in correction_summary["punctuation"].items():
                lines.append(f"  '{original}' → '{corrected}'")
                
        if correction_summary.get("capitalization", {}):
            lines.append("\nCAPITALIZATION CORRECTIONS:")
            for original, corrected in correction_summary["capitalization"].items():
                lines.append(f"  '{original}' → '{corrected}'")
        
        return "\n".join(lines)
    
    def apply_all(self, paragraphs: List[str], chunk_size: int = 3) -> Tuple[List[str], Dict]:
        """
        Apply comprehensive text correction to paragraphs and return corrected text.
        Uses Gemini AI for spelling, grammar, punctuation, and capitalization if available.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Tuple of (corrected_paragraphs, correction_summary)
        """
        print("Using basic spell checker...")
        return self.correct_document(
            paragraphs,
            auto_fix_spelling=True,
            auto_fix_grammar=False,
            use_ai_correction=False,
            chunk_size=chunk_size
        )
        

    
    def close(self):
        """Clean up resources."""
        self.document_checker.close()
