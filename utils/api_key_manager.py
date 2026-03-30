"""
API Key Manager for handling multiple API keys and rotation.
"""
import os
import logging
from typing import List, Optional

# Configure logging
# logging.basicConfig(level=logging.INFO) # Removed to avoid overriding global config
logger = logging.getLogger(__name__)

class APIKeyManager:
    """Manages multiple API keys with rotation and failure tracking."""
    
    def __init__(self):
        """Initialize with API keys from environment variables."""
        self.keys = []
        self.current_key_index = 0
        self.failed_keys = set()
        self._load_api_keys()
    
    def _load_api_keys(self):
        """Load API keys from environment variables only.
        
        Supports multiple API keys through numbered environment variables:
        - GEMINI_API_KEY (primary)
        - GEMINI_API_KEY_1, GEMINI_API_KEY_2, etc. (additional keys)
        
        All keys starting with GEMINI_API_KEY will be loaded and used for rotation.
        """
        self.keys = []
        
        # Load all GEMINI_API_KEY* environment variables
        for key, value in os.environ.items():
            if key.startswith('GEMINI_API_KEY') and value.strip():
                if value not in self.keys:  # Avoid duplicates
                    self.keys.append(value)
        
        # Remove any empty or None keys
        self.keys = [key for key in self.keys if key and key.strip()]
        
        if not self.keys:
            logger.error(
                "No valid API keys found. Please set GEMINI_API_KEY in your .env file.\n"
                "For multiple keys, use: GEMINI_API_KEY, GEMINI_API_KEY_1, GEMINI_API_KEY_2, etc."
            )
        else:
            logger.info(f"Successfully loaded {len(self.keys)} API key(s)")
    
    def get_next_key(self) -> Optional[str]:
        """Get the next available API key.
        
        Returns:
            str: The next available API key, or None if no keys are available.
        """
        if not self.keys:
            return None
            
        # Try to find a key that hasn't failed
        start_index = self.current_key_index
        while True:
            key = self.keys[self.current_key_index]
            self.current_key_index = (self.current_key_index + 1) % len(self.keys)
            
            if key not in self.failed_keys:
                return key
                
            # If we've gone through all keys and they've all failed
            if self.current_key_index == start_index:
                logger.error("All API keys have been marked as failed.")
                return None
    
    def mark_key_failed(self, key: str):
        """Mark an API key as failed.
        
        Args:
            key: The API key that failed.
        """
        self.failed_keys.add(key)
        logger.warning("An API key has been marked as failed.")
    
    def get_available_key_count(self) -> int:
        """Get the number of available API keys."""
        return len(self.keys) - len(self.failed_keys)

# Global instance
api_key_manager = APIKeyManager()
