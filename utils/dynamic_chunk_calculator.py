"""
Dynamic Chunk Size Calculator for Formatly V6
Calculates optimal chunk sizes based on document characteristics, rate limits, and performance considerations.
"""

import re
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from .rate_limit_manager import RateLimitManager


@dataclass
class DocumentMetrics:
    """Metrics about a document for chunk size calculation."""
    total_paragraphs: int
    avg_paragraph_length: float
    max_paragraph_length: int
    total_word_count: int
    avg_words_per_paragraph: float
    complexity_score: float
    estimated_processing_time: float


class DynamicChunkCalculator:
    """
    Calculates optimal chunk sizes based on document characteristics and API constraints.
    """
    
    # Base chunk sizes for different document types
    BASE_CHUNK_SIZES = {
        "small": 5,      # < 10 paragraphs
        "medium": 3,     # 10-50 paragraphs  
        "large": 2,      # 50-200 paragraphs
        "xlarge": 1,     # > 200 paragraphs
    }
    
    # Complexity multipliers
    COMPLEXITY_MULTIPLIERS = {
        "simple": 1.5,    # Simple text, larger chunks
        "moderate": 1.0,  # Normal complexity
        "complex": 0.7,   # Complex text, smaller chunks
        "very_complex": 0.5  # Very complex, smallest chunks
    }
    
    def __init__(self, rate_limit_manager: RateLimitManager):
        """
        Initialize the dynamic chunk calculator.
        
        Args:
            rate_limit_manager: Rate limit manager for API constraints
        """
        self.rate_limit_manager = rate_limit_manager
        self.max_tokens_per_request = self._get_max_tokens_per_request()
        self.requests_per_minute = self._get_requests_per_minute()
    
    def _get_max_tokens_per_request(self) -> int:
        """Get maximum tokens per request based on model."""
        model_limits = {
            "gemini-2.0-flash": 1000000,
            "gemini-2.0-flash-lite": 1000000,
            "gemini-2.5-flash": 250000,
            "gemini-2.5-flash-lite": 250000,
            "gemini-2.5-pro": 250000,
            "gemini-1.5-flash": 250000,
            "gemini-1.5-pro": 32000,
        }
        
        model_name = self.rate_limit_manager.model_name
        return model_limits.get(model_name, 100000)  # Conservative default
    
    def _get_requests_per_minute(self) -> int:
        """Get requests per minute limit from rate limit manager."""
        return self.rate_limit_manager.rate_limits.get("rpm", 10)
    
    def analyze_document(self, paragraphs: List[str]) -> DocumentMetrics:
        """
        Analyze document characteristics for chunk size calculation.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            DocumentMetrics with analysis results
        """
        if not paragraphs:
            return DocumentMetrics(0, 0, 0, 0, 0, 0, 0)
        
        # Basic metrics
        total_paragraphs = len(paragraphs)
        paragraph_lengths = [len(p) for p in paragraphs]
        avg_paragraph_length = sum(paragraph_lengths) / total_paragraphs
        max_paragraph_length = max(paragraph_lengths)
        
        # Word count metrics
        word_counts = [len(p.split()) for p in paragraphs]
        total_word_count = sum(word_counts)
        avg_words_per_paragraph = total_word_count / total_paragraphs
        
        # Complexity analysis
        complexity_score = self._calculate_complexity_score(paragraphs)
        
        # Estimated processing time (rough estimate)
        estimated_processing_time = self._estimate_processing_time(paragraphs)
        
        return DocumentMetrics(
            total_paragraphs=total_paragraphs,
            avg_paragraph_length=avg_paragraph_length,
            max_paragraph_length=max_paragraph_length,
            total_word_count=total_word_count,
            avg_words_per_paragraph=avg_words_per_paragraph,
            complexity_score=complexity_score,
            estimated_processing_time=estimated_processing_time
        )
    
    def _calculate_complexity_score(self, paragraphs: List[str]) -> float:
        """
        Calculate document complexity score based on various factors.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Complexity score (0-1, where 1 is most complex)
        """
        if not paragraphs:
            return 0.0
        
        complexity_factors = []
        
        for paragraph in paragraphs:
            # Average sentence length
            sentences = re.split(r'[.!?]+', paragraph)
            sentences = [s.strip() for s in sentences if s.strip()]
            if sentences:
                avg_sentence_length = sum(len(s.split()) for s in sentences) / len(sentences)
                complexity_factors.append(min(avg_sentence_length / 25, 1.0))  # Normalize to 0-1
            
            # Special characters and formatting
            special_chars = len(re.findall(r'[(){}\[\]"\'`~@#$%^&*+=|\\:;<>,/]', paragraph))
            complexity_factors.append(min(special_chars / 50, 1.0))  # Normalize to 0-1
            
            # Academic/technical terms (rough heuristic)
            academic_patterns = [
                r'\b\w+tion\b',  # -tion endings
                r'\b\w+ment\b',  # -ment endings
                r'\b\w+ence\b',  # -ence endings
                r'\b\w+ance\b',  # -ance endings
                r'\b\w{10,}\b',  # Long words
                r'\([^)]+\)',    # Parenthetical content
                r'\b\d+\.\d+\b', # Numbers with decimals
            ]
            
            academic_count = sum(len(re.findall(pattern, paragraph)) for pattern in academic_patterns)
            complexity_factors.append(min(academic_count / 20, 1.0))  # Normalize to 0-1
        
        return sum(complexity_factors) / len(complexity_factors) if complexity_factors else 0.0
    
    def _estimate_processing_time(self, paragraphs: List[str]) -> float:
        """
        Estimate processing time based on document characteristics.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            Estimated processing time in seconds
        """
        # Base processing time per paragraph (rough estimate)
        base_time_per_paragraph = 0.5  # seconds
        
        # Adjust for paragraph length
        total_chars = sum(len(p) for p in paragraphs)
        char_multiplier = 1 + (total_chars / 100000)  # Increase time for longer documents
        
        # Adjust for rate limits
        rate_limit_multiplier = 1 + (60 / self.requests_per_minute)  # Account for rate limiting
        
        return len(paragraphs) * base_time_per_paragraph * char_multiplier * rate_limit_multiplier
    
    def _classify_document_size(self, paragraph_count: int) -> str:
        """Classify document size based on paragraph count."""
        if paragraph_count < 10:
            return "small"
        elif paragraph_count < 50:
            return "medium"
        elif paragraph_count < 200:
            return "large"
        else:
            return "xlarge"
    
    def _classify_complexity(self, complexity_score: float) -> str:
        """Classify document complexity."""
        if complexity_score < 0.3:
            return "simple"
        elif complexity_score < 0.5:
            return "moderate"
        elif complexity_score < 0.7:
            return "complex"
        else:
            return "very_complex"
    
    def calculate_optimal_chunk_size(self, paragraphs: List[str], 
                                   user_chunk_size: Optional[int] = None,
                                   max_chunk_size: int = 10,
                                   min_chunk_size: int = 1) -> Tuple[int, Dict]:
        """
        Calculate optimal chunk size based on document characteristics.
        
        Args:
            paragraphs: List of paragraph texts
            user_chunk_size: User-specified chunk size (overrides calculation)
            max_chunk_size: Maximum allowed chunk size
            min_chunk_size: Minimum allowed chunk size
            
        Returns:
            Tuple of (optimal_chunk_size, analysis_report)
        """
        if user_chunk_size is not None:
            # User override - validate and use
            chunk_size = max(min_chunk_size, min(user_chunk_size, max_chunk_size))
            return chunk_size, {
                "method": "user_override",
                "original_user_size": user_chunk_size,
                "final_size": chunk_size,
                "reason": "User-specified chunk size"
            }
        
        # Analyze document
        metrics = self.analyze_document(paragraphs)
        
        if metrics.total_paragraphs == 0:
            return min_chunk_size, {
                "method": "default",
                "final_size": min_chunk_size,
                "reason": "Empty document"
            }
        
        # Base chunk size from document size
        doc_size_class = self._classify_document_size(metrics.total_paragraphs)
        base_chunk_size = self.BASE_CHUNK_SIZES[doc_size_class]
        
        # Complexity adjustment
        complexity_class = self._classify_complexity(metrics.complexity_score)
        complexity_multiplier = self.COMPLEXITY_MULTIPLIERS[complexity_class]
        
        # Rate limit adjustment
        rate_limit_factor = self._calculate_rate_limit_factor(metrics)
        
        # Token limit adjustment
        token_limit_factor = self._calculate_token_limit_factor(metrics)
        
        # Calculate final chunk size
        adjusted_chunk_size = base_chunk_size * complexity_multiplier * rate_limit_factor * token_limit_factor
        
        # Apply bounds
        final_chunk_size = max(min_chunk_size, min(int(round(adjusted_chunk_size)), max_chunk_size))
        
        # Create analysis report
        analysis_report = {
            "method": "dynamic_calculation",
            "document_metrics": {
                "total_paragraphs": metrics.total_paragraphs,
                "avg_paragraph_length": round(metrics.avg_paragraph_length, 1),
                "avg_words_per_paragraph": round(metrics.avg_words_per_paragraph, 1),
                "complexity_score": round(metrics.complexity_score, 3),
                "estimated_processing_time": round(metrics.estimated_processing_time, 1)
            },
            "calculation_factors": {
                "document_size_class": doc_size_class,
                "complexity_class": complexity_class,
                "base_chunk_size": base_chunk_size,
                "complexity_multiplier": complexity_multiplier,
                "rate_limit_factor": rate_limit_factor,
                "token_limit_factor": token_limit_factor,
                "raw_calculated_size": round(adjusted_chunk_size, 2)
            },
            "final_size": final_chunk_size,
            "reason": f"Optimized for {doc_size_class} {complexity_class} document with {metrics.total_paragraphs} paragraphs"
        }
        
        return final_chunk_size, analysis_report
    
    def _calculate_rate_limit_factor(self, metrics: DocumentMetrics) -> float:
        """Calculate adjustment factor based on rate limits."""
        # If we have low rate limits, use smaller chunks to avoid hitting limits
        if self.requests_per_minute <= 5:
            return 0.5  # Use smaller chunks
        elif self.requests_per_minute <= 15:
            return 0.8  # Slightly smaller chunks
        else:
            return 1.0  # Normal chunks
    
    def _calculate_token_limit_factor(self, metrics: DocumentMetrics) -> float:
        """Calculate adjustment factor based on token limits."""
        # Rough estimate: 1 token ≈ 4 characters
        estimated_tokens_per_paragraph = metrics.avg_paragraph_length / 4
        
        # If paragraphs are very long, use smaller chunks
        if estimated_tokens_per_paragraph > 1000:
            return 0.5
        elif estimated_tokens_per_paragraph > 500:
            return 0.7
        else:
            return 1.0
    
    def get_processing_estimate(self, paragraphs: List[str], chunk_size: int) -> Dict:
        """
        Get estimated processing time and request count.
        
        Args:
            paragraphs: List of paragraph texts
            chunk_size: Chunk size to use
            
        Returns:
            Dictionary with processing estimates
        """
        metrics = self.analyze_document(paragraphs)
        
        # Calculate number of requests needed
        non_empty_paragraphs = len([p for p in paragraphs if p.strip()])
        estimated_requests = (non_empty_paragraphs + chunk_size - 1) // chunk_size  # Ceiling division
        
        # Estimate time based on rate limits
        time_per_request = 2.0  # Average time per request including processing
        rate_limit_delay = max(0, (60 / self.requests_per_minute) - time_per_request)
        
        total_time = estimated_requests * (time_per_request + rate_limit_delay)
        
        return {
            "estimated_requests": estimated_requests,
            "estimated_time_seconds": round(total_time, 1),
            "estimated_time_minutes": round(total_time / 60, 1),
            "requests_per_minute_limit": self.requests_per_minute,
            "chunk_size": chunk_size,
            "paragraphs_to_process": non_empty_paragraphs
        }
    
    def generate_chunk_report(self, paragraphs: List[str], chunk_size: int, 
                            analysis_report: Dict) -> str:
        """
        Generate a human-readable report about chunk size selection.
        
        Args:
            paragraphs: List of paragraph texts
            chunk_size: Selected chunk size
            analysis_report: Analysis report from calculate_optimal_chunk_size
            
        Returns:
            Formatted report string
        """
        estimate = self.get_processing_estimate(paragraphs, chunk_size)
        
        lines = []
        lines.append("=" * 50)
        lines.append("  DYNAMIC CHUNK SIZE ANALYSIS")
        lines.append("=" * 50)
        lines.append("")
        
        # Document overview
        metrics = analysis_report.get("document_metrics", {})
        lines.append("📄 DOCUMENT OVERVIEW")
        lines.append(f"  Total paragraphs: {metrics.get('total_paragraphs', 0)}")
        lines.append(f"  Average paragraph length: {metrics.get('avg_paragraph_length', 0)} chars")
        lines.append(f"  Average words per paragraph: {metrics.get('avg_words_per_paragraph', 0)}")
        lines.append(f"  Complexity score: {metrics.get('complexity_score', 0):.3f}")
        lines.append("")
        
        # Chunk size decision
        lines.append("⚙️ CHUNK SIZE DECISION")
        lines.append(f"  Selected chunk size: {chunk_size}")
        lines.append(f"  Method: {analysis_report.get('method', 'unknown')}")
        lines.append(f"  Reason: {analysis_report.get('reason', 'No reason provided')}")
        lines.append("")
        
        # Processing estimates
        lines.append("⏱️ PROCESSING ESTIMATES")
        lines.append(f"  Estimated requests: {estimate['estimated_requests']}")
        lines.append(f"  Estimated time: {estimate['estimated_time_minutes']} minutes")
        lines.append(f"  Rate limit: {estimate['requests_per_minute_limit']} requests/minute")
        lines.append("")
        
        # Calculation details (if available)
        if analysis_report.get("method") == "dynamic_calculation":
            calc_factors = analysis_report.get("calculation_factors", {})
            lines.append("🔬 CALCULATION DETAILS")
            lines.append(f"  Document size class: {calc_factors.get('document_size_class', 'unknown')}")
            lines.append(f"  Complexity class: {calc_factors.get('complexity_class', 'unknown')}")
            lines.append(f"  Base chunk size: {calc_factors.get('base_chunk_size', 'unknown')}")
            lines.append(f"  Complexity multiplier: {calc_factors.get('complexity_multiplier', 'unknown')}")
            lines.append(f"  Rate limit factor: {calc_factors.get('rate_limit_factor', 'unknown')}")
            lines.append(f"  Token limit factor: {calc_factors.get('token_limit_factor', 'unknown')}")
            lines.append("")
        
        lines.append("=" * 50)
        return "\n".join(lines)
