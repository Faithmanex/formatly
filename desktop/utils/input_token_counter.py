"""
Input Token Counter for Formatly
Uses the official Google Gemini SDK to count tokens in input documents.
"""

import google.generativeai as genai
from typing import List, Dict, Tuple
from pathlib import Path
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-2.5-pro")

class InputTokenCounter:
    """Counts tokens in input documents using official Gemini SDK."""
    
    def __init__(self, api_key: str = API_KEY, model_name: str = MODEL_NAME):
        self.model_name = model_name
        self._initialize_client(api_key)
    
    def _initialize_client(self, api_key: str):
        """Initialize the Gemini client."""
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable is not set")
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(self.model_name)
    
    def count_tokens(self, text: str) -> int:
        """
        Count tokens in the given text using the official Gemini countTokens method.
        
        Args:
            text: The text to count tokens in
            
        Returns:
            Number of tokens in the text
        """
        # Use the official countTokens method from Gemini SDK
        return self.model.count_tokens(text).total_tokens
    
    def estimate_chunks_needed(self, text: str, chunk_size: int = 50000) -> Tuple[int, List[str]]:
        """
        Estimate how many chunks are needed for the given text.
        
        Args:
            text: The text to analyze
            chunk_size: Maximum tokens per chunk (default 50000)
            
        Returns:
            Tuple of (number of chunks needed, list of chunks)
        """
        total_tokens = self.count_tokens(text)
        print(f"\nInput text analysis:")
        print(f"Total tokens in text: {total_tokens}")
        print(f"Maximum chunk size: {chunk_size}")
        
        chunks_needed = (total_tokens + chunk_size - 1) // chunk_size  # Ceiling division
        
        if chunks_needed > 1:
            # Split text into roughly equal chunks by characters first
            chars_per_chunk = len(text) // chunks_needed
            chunks = []
            
            start = 0
            for i in range(chunks_needed):
                if i == chunks_needed - 1:
                    # Last chunk gets the remainder
                    chunk_text = text[start:]
                else:
                    # Find the next paragraph break after the target position
                    target = start + chars_per_chunk
                    # Look for next paragraph break (preferring \n\n, then \n)
                    next_break = text.find('\n\n', target)
                    if next_break == -1:
                        next_break = text.find('\n', target)
                    if next_break == -1:
                        next_break = target  # No breaks found, split at target
                    
                    chunk_text = text[start:next_break]
                    start = next_break + 1
                
                # Verify chunk token count and adjust if needed
                chunk_tokens = self.count_tokens(chunk_text)
                if chunk_tokens > chunk_size:
                    # If chunk is too large, reduce it by finding a closer break point
                    reduction_ratio = chunk_size / chunk_tokens
                    target = int(len(chunk_text) * reduction_ratio)
                    next_break = chunk_text.find('\n\n', target)
                    if next_break == -1:
                        next_break = chunk_text.find('\n', target)
                    if next_break == -1:
                        next_break = target
                    chunk_text = chunk_text[:next_break]
                    start = start - (len(chunk_text) - next_break)  # Adjust start for next iteration
                
                chunk_token_count = self.count_tokens(chunk_text)
                print(f"Chunk {len(chunks) + 1}: {chunk_token_count} tokens")
                chunks.append(chunk_text)
            
            print(f"\nChunking summary:")
            print(f"Total input tokens: {total_tokens}")
            print(f"Number of chunks created: {chunks_needed}")
            print(f"Average tokens per chunk: {total_tokens / chunks_needed:.0f}")
            print("-" * 50)
            return chunks_needed, chunks
        else:
            print(f"\nChunking summary:")
            print(f"Total input tokens: {total_tokens}")
            print(f"Text fits in a single chunk (under {chunk_size} tokens)")
            print("-" * 50)
            return 1, [text]

# Example usage when run directly
if __name__ == "__main__":
    counter = InputTokenCounter()
    
    print("\n=== Testing with small text (should fit in one chunk) ===")
    small_text = "This is a small test text that should fit in a single chunk." * 10
    _, _ = counter.estimate_chunks_needed(small_text, chunk_size=50000)
    
    print("\n=== Testing with large text (should require multiple chunks) ===")
    large_text = "qwertyuiop" * 4991  # Will generate a larger text
    _, chunks = counter.estimate_chunks_needed(large_text, chunk_size=50000)
    
    print("\nValidation of chunk sizes:")
    print("-" * 50)
    for i, chunk in enumerate(chunks, 1):
        chunk_tokens = counter.count_tokens(chunk)
        print(f"Chunk {i}: {chunk_tokens} tokens, {len(chunk)} characters")
