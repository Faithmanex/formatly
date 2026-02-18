"""
AI Client Wrappers
------------------
This module provides a unified interface for interacting with different AI providers
(Google Gemini, Hugging Face, Groq). It abstracts away the specific API details,
handling authentication, rate limiting, and response parsing.

Classes:
    - AIClient: Abstract base class defining the interface.
    - HuggingFaceClient: Implementation for Hugging Face Inference API (via OpenAI SDK compatibility).
    - GeminiClient: Implementation for Google's Gemini API.
    - GroqClient: Implementation for Groq API (specifically for chat features).
"""

from abc import ABC, abstractmethod
import os
import json
import traceback
from openai import OpenAI
import google.generativeai as genai
from utils.rate_limit_manager import RateLimitManager

class AIClient(ABC):
    """
    Abstract base class for AI clients.

    All specific AI provider implementations must inherit from this class and
    implement the `detect_structure` method and `model_name` property.
    """
    
    @abstractmethod
    def detect_structure(self, system_prompt: str, user_prompt: str) -> str:
        """
        Sends prompts to the AI model and returns the response text.
        Implementations should handle streaming if desired, but must return full text.

        Args:
            system_prompt (str): The system-level instruction (context, role).
            user_prompt (str): The user's input (document content).

        Returns:
            str: The raw text response from the AI.
        """
        pass
    
    @property
    @abstractmethod
    def model_name(self):
        pass

class HuggingFaceClient(AIClient):
    """
    Client for interacting with Hugging Face's API.
    Uses the OpenAI SDK as a wrapper for the HF inference endpoints.
    """
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv("HF_API_KEY")
        if not self.api_key:
            raise ValueError("HF_API_KEY not provided")
            
        self._model_name = "deepseek-ai/DeepSeek-V3.2:novita"
        self.rate_limit_manager = RateLimitManager(self._model_name)
        self.client = OpenAI(
            base_url="https://router.huggingface.co/v1",
            api_key=self.api_key,
            timeout=120.0,
        )

    @property
    def model_name(self):
        return self._model_name

    def detect_structure(self, system_prompt: str, user_prompt: str) -> tuple[str, object]:
        """
        Executes request to Hugging Face via OpenAI SDK.
        Returns: (full_text_response, usage_stats)
        """
        def _make_request(api_key=None, **kwargs):
            print("\n🤖 AI Response (Hugging Face)...\n")
            stream = self.client.chat.completions.create(
                model=self.model_name,
                max_tokens=65536,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                stream=True,
                stream_options={"include_usage": True},
            )
            
            full_text = ""
            usage_stats = None
            
            for chunk in stream:
                if chunk.choices and chunk.choices[0].delta.content:
                    content = chunk.choices[0].delta.content
                    print(content, end='', flush=True)
                    full_text += content
                if hasattr(chunk, 'usage') and chunk.usage:
                    usage_stats = chunk.usage
            
            print("\n")
            return full_text, usage_stats

        return self.rate_limit_manager.execute_with_rate_limit(_make_request)

class GeminiClient(AIClient):
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")
        if not self.api_key:
            raise ValueError("GEMINI_API_KEY not provided")
            
        self._model_name = os.getenv("GEMINI_MODEL", "")
        self.rate_limit_manager = RateLimitManager(self._model_name)
        genai.configure(api_key=self.api_key)

    @property
    def model_name(self):
        return self._model_name

    def detect_structure(self, system_prompt: str, user_prompt: str) -> tuple[str, object]:
        """
        Executes request to Google Gemini.
        Returns: (full_text_response, None) - Gemini usage stats handling might differ
        """
        def _make_request(api_key=None, **kwargs):
            model = genai.GenerativeModel(
                model_name=self.model_name,
                system_instruction=system_prompt
            )
            
            # Define safety settings to prevent false positives on academic content
            safety_settings = [
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
            ]

            print("\n🤖 AI Response (Gemini)...\n")
            response = model.generate_content(
                user_prompt,
                request_options={'timeout': 600},
                stream=True,
                safety_settings=safety_settings
            )
            
            full_text = ""
            for chunk in response:
                try:
                    if chunk.text:
                        print(chunk.text, end='', flush=True)
                        full_text += chunk.text
                except ValueError:
                    # Provide visual feedback that a chunk was blocked/skipped
                    print("⚠️", end='', flush=True)
            
            print("\n")
            # Gemini typically doesn't stream usage in the same chunk object way as OpenAI
            # We might need to access response.usage_metadata if available at the end
            # For compatibility with the formatter expected signature, we return None for usage if not easily available
            return full_text, None

        return self.rate_limit_manager.execute_with_rate_limit(_make_request)

class GroqClient:
    """Client specifically for the AI Chat feature, not for document formatting/structure detection."""
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv("GROQ_API_KEY")
        if not self.api_key:
            raise ValueError("GROQ_API_KEY not provided")
            
        self._model_name = "llama3-70b-8192" # Default to a good Groq model
        self.rate_limit_manager = RateLimitManager(self._model_name)
        self.client = OpenAI(
            base_url="https://api.groq.com/openai/v1",
            api_key=self.api_key,
            timeout=30.0,
        )

    @property
    def model_name(self):
        return self._model_name

    def generate_chat_response(self, system_prompt: str, user_prompt: str) -> tuple[str, object]:
        """
        Executes chat request to Groq via OpenAI SDK.
        Returns: (full_text_response, usage_stats)
        """
        def _make_request(api_key=None, **kwargs):
            print("\n🤖 AI Response (Groq Chat)...")
            stream = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                stream=True,
                stream_options={"include_usage": True}
            )
            
            full_text = ""
            usage_stats = None
            
            for chunk in stream:
                if chunk.choices and chunk.choices[0].delta.content:
                    content = chunk.choices[0].delta.content
                    print(content, end='', flush=True)
                    full_text += content
                if hasattr(chunk, 'usage') and chunk.usage:
                    usage_stats = chunk.usage
            
            print("\n")
            return full_text, usage_stats

        return self.rate_limit_manager.execute_with_rate_limit(_make_request)
