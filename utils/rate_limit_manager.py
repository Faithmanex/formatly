"""
Rate Limit Manager
------------------
Handles rate limiting, quota detection, and retry logic for AI API calls.
Designed primarily for Google's Gemini API but adaptable for others.

Features:
    - Token bucket-style rate limiting (RPM, TPM, RPD).
    - Automatic detection of quota limits from API error messages.
    - Exponential backoff and retry logic.
    - Integration with APIKeyManager for key rotation (if configured).
"""

import time
import os
import re
import sys
from typing import Dict, Optional, Tuple
from datetime import datetime, timedelta
import logging
from .api_key_manager import api_key_manager
from core.config import MODEL_QUOTAS

# Configure logging
# Configure logging
# logging.basicConfig(level=logging.INFO) # Removed to allow entry-point configuration
logger = logging.getLogger(__name__)

class DailyQuotaExceededException(Exception):
    """
    Custom exception raised when daily quota limits are exceeded.
    This exception terminates the entire application immediately.
    """

    def __init__(self, message: str, quota_type: str, quota_value: int, model: str = None):
        super().__init__(message)
        self.quota_type = quota_type
        self.quota_value = quota_value
        self.model = model

class RateLimitManager:
    """
    Manages rate limits for Gemini API requests with dynamic quota detection.
    Handles retries with exponential backoff and user notifications.
    Integrates with APIKeyManager for key rotation on rate limit failures.
    """

    def __init__(self, model_name: str = "gemini-2.0-flash"):
        """
        Initialize the rate limit manager.

        Args:
            model_name: The Gemini model being used
        """
        self.model_name = model_name
        self.request_times = []
        self.daily_requests = 0
        self.daily_reset_time = datetime.now() + timedelta(days=1)
        self.current_api_key = None

        # Get rate limits for the model from MODEL_QUOTAS with fallback
        # Get rate limits for the model from MODEL_QUOTAS
        default_limits = {"rpm": 5, "tpm": 100000, "rpd": 100, "max_tokens": 50000}

        if os.getenv("DEFAULT_BACKEND") == "huggingface":
            self.rate_limits = MODEL_QUOTAS.get(model_name, default_limits)
        else:
            # For other backends (Gemini), require explicit quotas
            self.rate_limits = MODEL_QUOTAS.get(model_name)
            if not self.rate_limits:
                 raise ValueError(f"Rate limits for model '{model_name}' not defined in config.MODEL_QUOTAS")

        logger.info(f"Rate limit manager initialized for {model_name}")
        logger.info(f"Limits: {self.rate_limits['rpm']} RPM, {self.rate_limits.get('tpm', 100000)} TPM, {self.rate_limits['rpd']} RPD")

    def extract_rate_limit_info(self, error_message: str) -> Dict:
        """
        Extract rate limit information from error message.

        Args:
            error_message: The error message from the API

        Returns:
            Dictionary with quota information
        """
        rate_info = {}

        # Extract quota metric and value
        quota_metric_match = re.search(r'quota_metric: "([^"]*)"', error_message)
        if quota_metric_match:
            rate_info["quota_metric"] = quota_metric_match.group(1)

        quota_value_match = re.search(r'quota_value: (\d+)', error_message)
        if quota_value_match:
            rate_info["quota_value"] = int(quota_value_match.group(1))

        # Extract quota ID for more detailed limit type detection
        quota_id_match = re.search(r'quota_id: "([^"]*)"', error_message)
        if quota_id_match:
            rate_info["quota_id"] = quota_id_match.group(1)

        # Extract model information
        model_match = re.search(r'key: "model"\s*value: "([^"]*)"', error_message)
        if model_match:
            rate_info["model"] = model_match.group(1)

        # Extract location information
        location_match = re.search(r'key: "location"\s*value: "([^"]*)"', error_message)
        if location_match:
            rate_info["location"] = location_match.group(1)

        # Extract retry delay
        retry_delay_match = re.search(r'retry_delay {\s*seconds: (\d+)', error_message)
        if retry_delay_match:
            rate_info["retry_delay"] = int(retry_delay_match.group(1))

        return rate_info

    def detect_limit_type(self, error_message: str) -> str:
        """
        Detect the type of rate limit that was violated.

        Args:
            error_message: The error message from the API

        Returns:
            String indicating the limit type: 'rpm', 'rpd_model', 'rpd_project', 'tpm', or 'unknown'
        """
        rate_info = self.extract_rate_limit_info(error_message)

        quota_metric = rate_info.get("quota_metric", "")
        quota_id = rate_info.get("quota_id", "")

        # Parse quota_id for specific patterns (e.g., "GenerateRequestsPerDayPerProjectPerModel-FreeTier")
        if quota_id:
            quota_id_lower = quota_id.lower()

            # Check for requests per minute patterns
            if "requestsperminute" in quota_id_lower or "rpm" in quota_id_lower:
                return "rpm"

            # Check for tokens per minute patterns
            elif "tokensperminute" in quota_id_lower or "tpm" in quota_id_lower:
                return "tpm"

            # Check for requests per day patterns
            elif "requestsperday" in quota_id_lower or "rpd" in quota_id_lower:
                # Determine if it's model-specific or project-wide
                if "permodel" in quota_id_lower:
                    return "rpd_model"
                elif "perproject" in quota_id_lower:
                    return "rpd_project"
                else:
                    # Default to model-specific if unclear
                    return "rpd_model"

        # Fallback to quota_metric if quota_id parsing fails
        if quota_metric:
            if "RequestsPerMinute" in quota_metric:
                return "rpm"
            elif "TokensPerMinute" in quota_metric:
                return "tpm"
            elif "RequestsPerDay" in quota_metric:
                # Check if it's model-specific or project-wide
                if rate_info.get("model") or "model" in quota_metric.lower():
                    return "rpd_model"
                else:
                    return "rpd_project"

        # Final fallback detection based on error message content
        error_lower = error_message.lower()
        if "requests per minute" in error_lower:
            return "rpm"
        elif "requests per day" in error_lower:
            if "model" in error_lower:
                return "rpd_model"
            else:
                return "rpd_project"
        elif "tokens per minute" in error_lower:
            return "tpm"

        return "unknown"

    def update_rate_limits_from_error(self, error_message: str):
        """
        Update rate limits based on error message information.

        Args:
            error_message: The error message from the API
        """
        rate_info = self.extract_rate_limit_info(error_message)

        if "quota_value" in rate_info:
            quota_value = rate_info["quota_value"]

            # Update RPM based on quota type
            if "RequestsPerMinute" in rate_info.get("quota_metric", ""):
                self.rate_limits["rpm"] = quota_value
                logger.info(f"Updated RPM limit to {quota_value} based on error response")

            # Update RPD if daily limit is mentioned
            elif "RequestsPerDay" in rate_info.get("quota_metric", ""):
                self.rate_limits["rpd"] = quota_value
                logger.info(f"Updated RPD limit to {quota_value} based on error response")

    def check_rate_limit(self) -> Tuple[bool, Optional[int]]:
        """
        Check if we're within rate limits and return wait time if needed.

        Returns:
            Tuple of (can_proceed, wait_seconds)
        """
        current_time = time.time()

        # Reset daily counter if needed
        if datetime.now() > self.daily_reset_time:
            self.daily_requests = 0
            self.daily_reset_time = datetime.now() + timedelta(days=1)

        # Check daily limit
        if self.daily_requests >= self.rate_limits["rpd"]:
            reset_in = (self.daily_reset_time - datetime.now()).total_seconds()
            return False, int(reset_in)

        # Clean old requests (older than 1 minute)
        minute_ago = current_time - 60
        self.request_times = [t for t in self.request_times if t > minute_ago]

        # Check per-minute limit
        if len(self.request_times) >= self.rate_limits["rpm"]:
            # Calculate wait time until oldest request is > 1 minute old
            oldest_request = min(self.request_times)
            wait_time = int(61 - (current_time - oldest_request))
            return False, wait_time

        return True, None

    def record_request(self):
        """Record a successful request."""
        self.request_times.append(time.time())
        self.daily_requests += 1

    def handle_rate_limit_error(self, error_message: str) -> Tuple[bool, int]:
        """
        Handle a rate limit error with intelligent quota type detection.

        Args:
            error_message: The error message from the API

        Returns:
            Tuple of (should_retry, wait_seconds)
        """
        # Update rate limits based on error
        self.update_rate_limits_from_error(error_message)

        # Extract rate limit info and detect type
        rate_info = self.extract_rate_limit_info(error_message)
        limit_type = self.detect_limit_type(error_message)

        # Handle different quota types
        if limit_type == "rpm":
            # Requests per minute - can retry after wait
            wait_time = rate_info.get("retry_delay", 60) + 2
            quota_value = rate_info.get("quota_value", self.rate_limits["rpm"])

            logger.warning(f"🔄 Requests per minute limit exceeded ({quota_value} RPM)")
            logger.info(f"   Waiting {wait_time} seconds before retrying...")

            return True, wait_time

        elif limit_type in ["rpd_model", "rpd_project"]:
            # Requests per day - should not retry
            quota_value = rate_info.get("quota_value", self.rate_limits["rpd"])
            model = rate_info.get("model", self.model_name)

            if limit_type == "rpd_model":
                error_message = f"🛑 Daily request limit exceeded: {quota_value} requests per day for model '{model}'"
                logger.error(error_message)
                logger.info(f"   This is a model-specific daily limit.")
                logger.info(f"   Suggestions:")
                logger.info(f"   • Switch to a different model (e.g., gemini-2.0-flash-lite)")
                logger.info(f"   • Upgrade to a paid plan for higher limits")
                logger.info(f"   • Wait until tomorrow for the quota to reset")

                # Raise custom exception to terminate the application
                raise DailyQuotaExceededException(
                    error_message,
                    quota_type="rpd_model",
                    quota_value=quota_value,
                    model=model
                )
            else:
                error_message = f"🛑 Daily request limit exceeded: {quota_value} requests per day across all models"
                logger.error(error_message)
                logger.info(f"   This is a project-wide daily limit.")
                logger.info(f"   Suggestions:")
                logger.info(f"   • Upgrade to a paid plan for higher limits")
                logger.info(f"   • Create a new project with separate quotas")
                logger.info(f"   • Wait until tomorrow for the quota to reset")

                # Raise custom exception to terminate the application
                raise DailyQuotaExceededException(
                    error_message,
                    quota_type="rpd_project",
                    quota_value=quota_value
                )

        elif limit_type == "tpm":
            # Tokens per minute - can retry after wait
            wait_time = rate_info.get("retry_delay", 60) + 2
            quota_value = rate_info.get("quota_value", self.rate_limits["tpm"])

            logger.warning(f"🔄 Tokens per minute limit exceeded ({quota_value:,} TPM)")
            logger.info(f"   Waiting {wait_time} seconds before retrying...")

            return True, wait_time

        else:
            # Unknown limit type - use fallback behavior
            wait_time = rate_info.get("retry_delay", 60) + 2

            logger.warning(f"⚠️ Unknown rate limit type detected")
            logger.warning(f"   Error: {error_message[:200]}...")
            logger.info(f"   Attempting to retry after {wait_time} seconds...")

            return True, wait_time

    def wait_with_progress(self, wait_seconds: int):
        """
        Wait for the specified time while showing progress to the user.

        Args:
            wait_seconds: Number of seconds to wait
        """
        if wait_seconds <= 0:
            return

        logger.info(f"⏳ Rate limit reached. Waiting {wait_seconds} seconds before retrying...")

        # Show progress for waits longer than 10 seconds
        if wait_seconds > 10:
            intervals = min(wait_seconds // 5, 20)  # Show progress up to 20 times
            interval_time = wait_seconds / intervals

            for i in range(intervals):
                remaining = wait_seconds - (i * interval_time)
                # print(f"   ⏳ {remaining:.0f} seconds remaining...") # Reduce noise in logs
                time.sleep(interval_time)
        else:
            time.sleep(wait_seconds)

        logger.info("✅ Wait completed. Resuming processing...")

    def execute_with_rate_limit(self, func, *args, **kwargs):
        """
        Execute a function with rate limit checking and retry logic.
        Integrates with APIKeyManager for key rotation on rate limit failures.

        Args:
            func: The function to execute
            *args: Arguments for the function
            **kwargs: Keyword arguments for the function

        Returns:
            The result of the function call
        """
        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            # Get API key if we don't have one or previous one failed
            if not self.current_api_key:
                self.current_api_key = api_key_manager.get_next_key()
                if not self.current_api_key:
                    raise Exception("No available API keys")
                kwargs["api_key"] = self.current_api_key

            # Check rate limit before making request
            can_proceed, wait_time = self.check_rate_limit()

            if not can_proceed:
                self.wait_with_progress(wait_time)
                continue

            try:
                # Execute the function with current API key
                result = func(*args, **kwargs)

                # Record successful request
                self.record_request()

                return result

            except Exception as e:
                error_str = str(e)

                # Check if this is a rate limit error
                # Check if this is a rate limit error or model overload (503)
                if ("429" in error_str and "exceeded your current quota" in error_str) or \
                   ("503" in error_str and "overloaded" in error_str):

                    retry_count += 1

                    is_overload = "503" in error_str

                    if not is_overload:
                        # Only mark key failed for rate limits, not server overloads
                        api_key_manager.mark_key_failed(self.current_api_key)
                        self.current_api_key = None

                    if is_overload:
                        # For overload, use standard backoff
                        should_retry = True
                        wait_time = 5 * (2 ** (retry_count - 1)) # 5, 10, 20...
                        logger.warning(f"⚠️ Service Overloaded (503). Retrying in {wait_time}s...")
                    else:
                        should_retry, wait_time = self.handle_rate_limit_error(error_str)

                    if not should_retry:
                        # Daily limit reached - don't retry
                        logger.error(f"❌ Daily quota limit reached. Stopping execution.")
                        raise

                    if retry_count < max_retries:
                         if not is_overload and api_key_manager.get_available_key_count() > 0:
                            logger.info(f"🔄 Rate limit exceeded (attempt {retry_count}/{max_retries})")
                            logger.info(f"🔑 Switching to next available API key...")
                         elif is_overload:
                            logger.info(f"🔄 Retrying request (attempt {retry_count}/{max_retries})")

                         self.wait_with_progress(wait_time)
                         continue
                    else:
                        logger.error(f"❌ Maximum retries ({max_retries}) exceeded or no more API keys available")
                        raise

                # If it's not a rate limit error, re-raise immediately
                raise

        # This should not be reached, but just in case
        raise Exception("Maximum retries exceeded")

    def get_status_info(self) -> Dict:
        """
        Get current status information about rate limits.

        Returns:
            Dictionary with current status
        """
        current_time = time.time()
        minute_ago = current_time - 60
        recent_requests = len([t for t in self.request_times if t > minute_ago])

        return {
            "model": self.model_name,
            "rate_limits": self.rate_limits,
            "requests_last_minute": recent_requests,
            "requests_today": self.daily_requests,
            "daily_reset_time": self.daily_reset_time.strftime("%Y-%m-%d %H:%M:%S")
        }
