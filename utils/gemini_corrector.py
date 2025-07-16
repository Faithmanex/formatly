"""
GeminiCorrector Module for Formatly V3
Uses Google's Gemini AI to provide advanced spelling, punctuation, and capitalization corrections.
"""

import google.generativeai as genai
import json
import os
import re
import time
from dotenv import load_dotenv
from typing import List, Dict, Tuple, Optional
from .rate_limit_manager import RateLimitManager, DailyQuotaExceededException
from .dynamic_chunk_calculator import DynamicChunkCalculator

# Load environment variables
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = os.getenv("GEMINI_MODEL")

class GeminiCorrector:
    """
    Uses Google's Gemini AI to correct spelling, punctuation, and capitalization in text.
    """
    
    def __init__(self, api_key: str = API_KEY, model_name: str = MODEL_NAME):
        """
        Initialize the Gemini corrector.
        
        Args:
            api_key: API key for Gemini (default: from environment)
            model_name: Model name to use (default: from environment)
        """
        self.api_key = api_key
        self.model_name = model_name or "gemini-2.0-flash"  # Default fallback
        self.model = None
        
        if not self.api_key:
            raise ValueError("GEMINI_API_KEY is required for GeminiCorrector")
        
        # Initialize rate limit manager
        self.rate_limit_manager = RateLimitManager(self.model_name)
        
        # Initialize dynamic chunk calculator
        self.chunk_calculator = DynamicChunkCalculator(self.rate_limit_manager)
        
        self._initialize_model()
    
    def _initialize_model(self):
        """Initialize the Gemini model."""
        try:
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel(self.model_name)
        except Exception as e:
            raise ValueError(f"Failed to initialize Gemini model: {e}")
    
    def correct_text(self, text: str) -> Tuple[str, Dict]:
        """
        Correct spelling, punctuation, and capitalization in text.
        
        Args:
            text: Text to correct
            
        Returns:
            Tuple of (corrected_text, correction_details)
        """
        prompt = self._create_correction_prompt(text)
        
        def _make_request():
            """Internal function to make the actual API request."""
            response = self.model.generate_content(prompt)
            return self._parse_correction_response(response.text, text)
        
        try:
            # Use rate limit manager to handle the request with automatic retries
            return self.rate_limit_manager.execute_with_rate_limit(_make_request)
        except DailyQuotaExceededException:
            # Re-raise daily quota exceptions to stop processing immediately
            raise
        except Exception as e:
            print(f"Error during Gemini correction: {e}")
            return text, {"error": str(e), "corrections": 0}
    
    def _create_correction_prompt(self, text: str) -> str:
        """Create a prompt for text correction."""
        return f"""
        You are a text correction assistant focusing ONLY on spelling, punctuation, and capitalization.
        
        Correct the following text ONLY for:
        1. Spelling errors
        2. Punctuation errors
        3. Capitalization errors
        
        DO NOT change the meaning, style, or sentence structure.
        DO NOT add or remove content except for fixing the error types listed above.
        
        IMPORTANT: Do all three types of corrections (spelling, punctuation, and capitalization) in a single pass.
        
        Return ONLY a JSON object with the following structure:
        {{
          "corrected_text": "The full corrected text",
          "corrections": [
            {{
              "original": "original text with error",
              "corrected": "corrected text",
              "type": "spelling|punctuation|capitalization"
            }}
          ]
        }}
        
        TEXT TO CORRECT:
        {text}
        
        Return valid JSON only - no additional text or markdown formatting.
        """
    
    def _parse_correction_response(self, response_text: str, original_text: str) -> Tuple[str, Dict]:
        """
        Parse the correction response from Gemini.
        
        Args:
            response_text: Response from Gemini
            original_text: Original text (fallback if parsing fails)
            
        Returns:
            Tuple of (corrected_text, correction_details)
        """
        try:
            # Clean up response text (remove markdown code blocks if present)
            text = response_text.strip()
            if text.startswith("```json"):
                text = text[7:]
            elif text.startswith("```"):
                text = text[3:]
            
            if text.endswith("```"):
                text = text[:-3]
            
            text = text.strip()
            
            # Parse JSON
            result = json.loads(text)
            
            # Validate the response format
            if "corrected_text" not in result:
                return original_text, {"error": "Invalid response format", "corrections": 0}
            
            # Return the corrected text and details
            return result["corrected_text"], {
                "corrections": len(result.get("corrections", [])),
                "details": result.get("corrections", [])
            }
            
        except json.JSONDecodeError:
            return original_text, {"error": "Failed to parse JSON response", "corrections": 0}
        except Exception as e:
            return original_text, {"error": str(e), "corrections": 0}
    
    def correct_paragraphs(self, paragraphs: List[str], chunk_size: Optional[int] = None) -> Tuple[List[str], Dict]:
        """
        Correct spelling, punctuation, and capitalization in a list of paragraphs.
        Processes paragraphs in chunks to avoid API limits with dynamic chunk sizing.

        Args:
            paragraphs: List of paragraph texts
            chunk_size: Number of paragraphs to process in each API request (optional, will be calculated dynamically)

        Returns:
            Tuple of (corrected_paragraphs, correction_summary)
        """
        corrected_paragraphs = [None] * len(paragraphs)  # Pre-allocate list with correct size
        all_corrections = []
        total_corrections = 0

        # Process non-empty paragraphs
        non_empty_indices = [i for i, p in enumerate(paragraphs) if p.strip()]
        total_paragraphs = len(non_empty_indices)

        if total_paragraphs == 0:
            return paragraphs, {"total_corrections": 0, "corrections_by_type": {}, "details": []}

        # Calculate optimal chunk size
        optimal_chunk_size, chunk_analysis = self.chunk_calculator.calculate_optimal_chunk_size(
            paragraphs, 
            user_chunk_size=chunk_size,
            max_chunk_size=10,
            min_chunk_size=1
        )

        print(f"Processing {total_paragraphs} paragraphs for spelling, punctuation, and capitalization...")
        print(f"Using {chunk_analysis['method']} chunk size: {optimal_chunk_size} paragraphs per API request")
        print(f"Reason: {chunk_analysis['reason']}")
        
        # Show processing estimate
        estimate = self.chunk_calculator.get_processing_estimate(paragraphs, optimal_chunk_size)
        print(f"Estimated processing time: {estimate['estimated_time_minutes']} minutes ({estimate['estimated_requests']} requests)")
        
        chunk_size = optimal_chunk_size

        # First handle empty paragraphs
        for i, paragraph in enumerate(paragraphs):
            if not paragraph.strip():
                corrected_paragraphs[i] = paragraph

        # Process non-empty paragraphs in chunks
        for start_idx in range(0, len(non_empty_indices), chunk_size):
            chunk_indices = non_empty_indices[start_idx:start_idx + chunk_size]
            chunk_paragraphs = [paragraphs[i] for i in chunk_indices]
            
            print(f"  Processing chunk {start_idx//chunk_size + 1} ({len(chunk_indices)} paragraphs)...")
            
            # Process each paragraph in the chunk individually to maintain separation
            for i, para_idx in enumerate(chunk_indices):
                try:
                    # Apply correction to individual paragraph
                    corrected_text, details = self.correct_text(paragraphs[para_idx])
                    corrected_paragraphs[para_idx] = corrected_text
                    
                    # Track corrections
                    if "details" in details and details["details"]:
                        for correction in details["details"]:
                            correction["paragraph_index"] = para_idx
                            all_corrections.append(correction)
                        total_corrections += details["corrections"]
                        
                except DailyQuotaExceededException:
                    # Re-raise daily quota exceptions to stop processing immediately
                    raise
                except Exception as e:
                    print(f"Error processing paragraph {para_idx}: {e}")
                    # Keep original on error
                    corrected_paragraphs[para_idx] = paragraphs[para_idx]
            
            # Show progress
            processed = min(start_idx + chunk_size, len(non_empty_indices))
            print(f"  Processed {processed}/{len(non_empty_indices)} paragraphs")

        # Create summary
        correction_summary = {
            "total_corrections": total_corrections,
            "corrections_by_type": self._summarize_corrections_by_type(all_corrections),
            "details": all_corrections
        }

        return corrected_paragraphs, correction_summary
    
    def _summarize_corrections_by_type(self, corrections: List[Dict]) -> Dict:
        """Summarize corrections by type."""
        summary = {
            "spelling": 0,
            "punctuation": 0,
            "capitalization": 0,
            "other": 0
        }
        
        for correction in corrections:
            correction_type = correction.get("type", "other")
            if correction_type in summary:
                summary[correction_type] += 1
            else:
                summary["other"] += 1
        
        return summary
    
    def generate_correction_report(self, correction_summary: Dict) -> str:
        """
        Generate a human-readable correction report.
        
        Args:
            correction_summary: Correction summary from correct_paragraphs
            
        Returns:
            Formatted string report
        """
        lines = []
        lines.append("=" * 50)
        lines.append("  FORMATLY V3 - GEMINI AI CORRECTION REPORT")
        lines.append("=" * 50)
        lines.append("")
        
        # Summary - safer access with default value of 0
        lines.append(f"📊 TOTAL CORRECTIONS: {correction_summary.get('total_corrections', 0)}")
        
        # Breakdown by type
        by_type = correction_summary.get("corrections_by_type", {})
        lines.append(f"🔤 Spelling: {by_type.get('spelling', 0)}")
        lines.append(f"🔣 Punctuation: {by_type.get('punctuation', 0)}")
        lines.append(f"🔠 Capitalization: {by_type.get('capitalization', 0)}")
        lines.append(f"⚙️ Other: {by_type.get('other', 0)}")
        lines.append("")
        
        # Details - safer access with default value of 0
        if correction_summary.get("total_corrections", 0) > 0:
            lines.append("CORRECTION DETAILS:")
            lines.append("-" * 30)
            
            # Get the first 10 corrections to show as examples
            for i, correction in enumerate(correction_summary.get("details", [])[:10]):
                original = correction.get("original", "")
                corrected = correction.get("corrected", "")
                corr_type = correction.get("type", "unknown").title()
                para_idx = correction.get("paragraph_index", "?")
                
                lines.append(f"{i+1}. [{corr_type}] Paragraph {para_idx+1}")
                lines.append(f"   Original: \"{original}\"")
                lines.append(f"   Corrected: \"{corrected}\"")
                lines.append("")
            
            # If there are more corrections, add a note
            details = correction_summary.get("details", [])
            if len(details) > 10:
                lines.append(f"... and {len(details) - 10} more corrections")
        
        lines.append("=" * 50)
        return "\n".join(lines)
