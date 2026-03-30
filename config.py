"""
Configuration module for Formatly V3.
Handles secure API key management and application settings.
"""

import os
from pathlib import Path
from dotenv import load_dotenv
from typing import Optional, Dict, Any

# Load environment variables
load_dotenv()

# Model quotas and limitations
MODEL_QUOTAS = {
    "gemini-2.0-flash": {
        "rpm": 2000,
        "tpm": 4000000,
        "rpd": 1000000,  # Unlimited
        "max_tokens": 120000
    },
    "gemini-2.0-flash-lite": {
        "rpm": 4000,
        "tpm": 4000000,
        "rpd": 1000000,  # Unlimited
        "max_tokens": 120000
    },
    "gemini-2.5-flash": {
        "rpm": 1000,
        "tpm": 1000000,
        "rpd": 10000,
        "max_tokens": 120000
    },
    "gemini-2.5-flash-lite": {
        "rpm": 4000,
        "tpm": 4000000,
        "rpd": 1000000,  # Unlimited
        "max_tokens": 120000
    },
    "gemini-2.5-pro": {
        "rpm": 150,
        "tpm": 2000000,
        "rpd": 1000,
        "max_tokens": 120000
    },
    "gemini-3-flash": {
        "rpm": 1000,
        "tpm": 2000000,
        "rpd": 10000,
        "max_tokens": 120000
    },
    "gemini-3.1-flash-lite": {
        "rpm": 4000,
        "tpm": 4000000,
        "rpd": 150000,
        "max_tokens": 120000
    },
    "gemini-3.1-pro": {
        "rpm": 25,
        "tpm": 2000000,
        "rpd": 250,
        "max_tokens": 120000
    },
    "llama-3.1-8b-instant": {
        "rpm": 5,
        "rpd": 100,
        "max_tokens": 120000,
        "tpm": 250000
    },
    "deepseek-ai/DeepSeek-V3.2:novita": {
        "rpm": 5,
        "rpd": 100,
        "max_tokens": 120000,
        "tpm": 250000
    }
}

class Config:
    """Configuration class for Formatly V3."""
    
    def __init__(self):
        self.api_key = self._get_api_key()
        self.model_name = os.getenv("GEMINI_MODEL")
        self.max_retries = int(os.getenv("MAX_RETRIES", "3"))
        self.timeout = int(os.getenv("TIMEOUT", "30"))
        self.log_level = os.getenv("LOG_LEVEL", "INFO")
        self.convertapi_api_key = os.getenv("CONVERTAPI_API_KEY", "")
        
        # Gemini API settings
        self.temperature = float(os.getenv("GEMINI_TEMPERATURE", "1.0"))
        self.top_p = float(os.getenv("GEMINI_TOP_P", "0.95"))
        self.top_k = int(os.getenv("GEMINI_TOP_K", "40"))
        self.max_output_tokens = int(os.getenv("GEMINI_MAX_OUTPUT_TOKENS", "1024"))
        self.stop_sequences = os.getenv("GEMINI_STOP_SEQUENCES", "\n\n").split(',')
        
    def _get_api_key(self) -> Optional[str]:
        """Securely retrieve API key from environment variables."""
        return os.getenv("GEMINI_API_KEY")
    
    def is_api_key_configured(self) -> bool:
        """Check if API key is configured."""
        return self.api_key is not None and len(self.api_key.strip()) > 0
    
    def get_safe_config(self) -> Dict[str, Any]:
        """Get configuration without sensitive data for logging."""
        return {
            "model_name": self.model_name,
            "max_retries": self.max_retries,
            "timeout": self.timeout,
            "log_level": self.log_level,
            "api_key_configured": self.is_api_key_configured(),
            "convertapi_configured": bool(self.convertapi_api_key)
        }

# Global configuration instance
config = Config()

# Supported file formats
SUPPORTED_FORMATS = [".docx"]

# Supported citation styles
SUPPORTED_STYLES = ["apa", "mla", "chicago", "turabian"]

# Default settings
DEFAULT_STYLE = "apa"
DEFAULT_OUTPUT_SUFFIX = "formatted"

# Application metadata
APP_NAME = "Formatly V1.0"
APP_VERSION = "1.0.0"
APP_DESCRIPTION = "AI-powered academic document formatter"
