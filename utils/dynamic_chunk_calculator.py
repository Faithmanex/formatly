"""
Dynamic Chunk Calculator for Formatly
Calculates optimal chunks based on token limits and rate limits.
"""

from typing import List
from .rate_limit_manager import RateLimitManager
from .input_token_counter import InputTokenCounter
from config import MODEL_QUOTAS


class DynamicChunkCalculator:
    """
    A thin wrapper around InputTokenCounter that adds rate limit awareness 
    and model-specific token limits.
    """
    
    def __init__(self, rate_limit_manager: RateLimitManager):
        """
        Initialize the dynamic chunk calculator.
        
        Args:
            rate_limit_manager: Rate limit manager for API constraints
        """
        self.rate_limit_manager = rate_limit_manager
        self.max_tokens_per_request = self._get_max_tokens_per_request()
        self.token_counter = InputTokenCounter()
    
    def _get_max_tokens_per_request(self) -> int:
        """Get maximum tokens per request based on model."""
        model_name = self.rate_limit_manager.model_name
        return MODEL_QUOTAS.get(model_name, {}).get("max_tokens", 50000)
    
    def split_doc_into_chunks(self, text: str) -> List[str]:
        """
        Split text into chunks based on model's token limit.
        
        Args:
            text: The text to split
            
        Returns:
            List of text chunks each under the model's token limit
        """
        print(f"\nDynamicChunkCalculator: Starting text chunking with max {self.max_tokens_per_request} tokens per chunk")
        chunks_needed, chunks = self.token_counter.estimate_chunks_needed(text, self.max_tokens_per_request)
        print(f"DynamicChunkCalculator: Created {len(chunks)} chunks")
        return chunks