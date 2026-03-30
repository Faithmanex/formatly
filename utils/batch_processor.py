"""
BatchProcessor Module for Formatly V3
Implements asynchronous document processing using Google's Gemini API Batch Mode.
Provides significant cost savings (50% off standard pricing) and handles large documents efficiently.
"""

import os
import json
import time
from typing import List, Dict, Tuple, Optional, Any
import google.generativeai as genai
from pathlib import Path
import uuid
import tempfile
import logging

logger = logging.getLogger(__name__)

class BatchProcessor:
    """
    Handles asynchronous batch processing of documents using Gemini API Batch Mode.
    Provides methods for creating, monitoring, and retrieving batch jobs.
    """
    
    def __init__(self, api_key: str = None, model_name: str = None):
        """
        Initialize the batch processor with API credentials.
        
        Args:
            api_key: Gemini API key (default: from environment)
            model_name: Gemini model to use (default: from environment)
        """
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")
        self.model_name = model_name or os.getenv("GEMINI_MODEL")
        self.client = None
        self.temp_dir = None
        
        if not self.api_key:
            raise ValueError("GEMINI_API_KEY is required for BatchProcessor")
        
        self._initialize_client()
        self.temp_dir = tempfile.TemporaryDirectory()
    
    def _initialize_client(self):
        """Initialize the Gemini client."""
        try:
            genai.configure(api_key=self.api_key)
            self.client = genai.Client()
            logger.info(f"Initialized Gemini client for Batch processing with model: {self.model_name}")
        except Exception as e:
            raise ValueError(f"Failed to initialize Gemini client: {e}")
    
    def create_batch_job_from_paragraphs(self, 
                                         paragraphs: List[str], 
                                         batch_name: Optional[str] = None,
                                         system_instruction: Optional[str] = None) -> str:
        """
        Create a batch job from a list of paragraphs.
        
        Args:
            paragraphs: List of paragraph texts to process
            batch_name: Optional name for the batch job
            system_instruction: Optional system instruction for all requests
            
        Returns:
            Batch job ID
        """
        if not batch_name:
            batch_name = f"formatly-batch-{uuid.uuid4().hex[:8]}"
        
        # Create a temporary JSONL file
        temp_file_path = os.path.join(self.temp_dir.name, f"{batch_name}.jsonl")
        
        # Prepare batch requests
        with open(temp_file_path, "w") as f:
            for i, paragraph in enumerate(paragraphs):
                if not paragraph.strip():
                    continue  # Skip empty paragraphs
                
                request = {
                    "contents": [{
                        "parts": [{"text": paragraph}],
                        "role": "user"
                    }]
                }
                
                # Add system instruction if provided
                if system_instruction:
                    request["system_instructions"] = {
                        "parts": [{"text": system_instruction}]
                    }
                
                # Add to JSONL file
                entry = {
                    "key": f"para-{i}",
                    "request": request
                }
                f.write(json.dumps(entry) + "\n")
        
        logger.info(f"Created batch request file with {len(paragraphs)} paragraphs")
        
        # Upload the file to the File API
        uploaded_file = self.client.files.upload(
            file=temp_file_path,
            config={"display_name": batch_name, "mime_type": "application/jsonl"}
        )
        logger.info(f"Uploaded batch request file: {uploaded_file.name}")
        
        # Create batch job
        batch_job = self.client.batches.create(
            model=self.model_name,
            src=uploaded_file.name,
            config={
                "display_name": batch_name,
            },
        )
        
        logger.info(f"Created batch job: {batch_job.name}")
        return batch_job.name
    
    def create_batch_job_for_formatting(self,
                                        document_text: str,
                                        style_name: str,
                                        batch_name: Optional[str] = None) -> str:
        """
        Create a specialized batch job for document formatting.
        
        Args:
            document_text: The full document text to format
            style_name: Citation style to apply (apa, mla, chicago)
            batch_name: Optional name for the batch job
            
        Returns:
            Batch job ID
        """
        if not batch_name:
            batch_name = f"formatly-{style_name}-{uuid.uuid4().hex[:8]}"
        
        # Format-specific system instruction
        system_instruction = f"""
        You are an academic formatting assistant specializing in {style_name.upper()} style.
        Your task is to correct and format text according to {style_name.upper()} guidelines.
        Fix spelling, punctuation, capitalization, and apply proper {style_name.upper()} formatting.
        Maintain the original meaning and content of the text.
        """
        
        # Split document into reasonably-sized chunks (3000 chars)
        chunks = self._split_document_into_chunks(document_text, 3000)
        
        return self.create_batch_job_from_paragraphs(
            paragraphs=chunks,
            batch_name=batch_name,
            system_instruction=system_instruction
        )
    
    def _split_document_into_chunks(self, text: str, max_chunk_size: int = 3000) -> List[str]:
        """
        Split a document into chunks of reasonable size for the API.
        
        Args:
            text: The full document text
            max_chunk_size: Maximum character size per chunk
            
        Returns:
            List of text chunks
        """
        # Split by paragraphs first
        paragraphs = text.split('\n\n')
        chunks = []
        current_chunk = ""
        
        for para in paragraphs:
            if len(current_chunk) + len(para) > max_chunk_size and current_chunk:
                chunks.append(current_chunk)
                current_chunk = para
            else:
                current_chunk += "\n\n" + para if current_chunk else para
        
        if current_chunk:
            chunks.append(current_chunk)
        
        logger.info(f"Split document into {len(chunks)} chunks")
        return chunks
    
    def check_batch_job_status(self, job_name: str) -> Dict:
        """
        Check the status of a batch job.
        
        Args:
            job_name: Batch job ID to check
            
        Returns:
            Dictionary with job status information
        """
        batch_job = self.client.batches.get(name=job_name)
        
        status = {
            "name": batch_job.name,
            "state": batch_job.state.name,
            "create_time": batch_job.create_time,
            "update_time": batch_job.update_time,
            "completed": batch_job.state.name in [
                'JOB_STATE_SUCCEEDED',
                'JOB_STATE_FAILED',
                'JOB_STATE_CANCELLED'
            ]
        }
        
        if batch_job.state.name == 'JOB_STATE_FAILED' and hasattr(batch_job, 'error'):
            status["error"] = batch_job.error
        
        logger.info(f"Job {job_name} status: {status['state']}")
        return status
    
    def wait_for_batch_job(self, job_name: str, 
                          polling_interval: int = 30,
                          max_wait_time: int = 3600) -> Dict:
        """
        Wait for a batch job to complete, with periodic status updates.
        
        Args:
            job_name: Batch job ID to wait for
            polling_interval: Seconds between status checks
            max_wait_time: Maximum seconds to wait before timing out
            
        Returns:
            Final job status
        """
        start_time = time.time()
        elapsed = 0
        
        logger.info(f"Waiting for batch job {job_name} to complete...")
        
        while elapsed < max_wait_time:
            status = self.check_batch_job_status(job_name)
            
            if status["completed"]:
                logger.info(f"Job completed with state: {status['state']}")
                return status
            
            logger.info(f"Job in progress... State: {status['state']}, Elapsed time: {elapsed:.0f}s")
            time.sleep(polling_interval)
            elapsed = time.time() - start_time
        
        logger.warning(f"Maximum wait time ({max_wait_time}s) exceeded")
        return {"state": "TIMEOUT", "completed": False, "name": job_name}
    
    def get_batch_results(self, job_name: str) -> List[Dict]:
        """
        Retrieve and parse the results of a completed batch job.
        
        Args:
            job_name: Batch job ID to retrieve results for
            
        Returns:
            List of processed results with original keys
        """
        batch_job = self.client.batches.get(name=job_name)
        
        if batch_job.state.name != 'JOB_STATE_SUCCEEDED':
            raise ValueError(f"Cannot retrieve results. Job state is: {batch_job.state.name}")
        
        results = []
        
        # Results are in a file
        if batch_job.dest and batch_job.dest.file_name:
            result_file_name = batch_job.dest.file_name
            logger.info(f"Downloading results from file: {result_file_name}")
            
            # Download and process the JSONL results file
            file_content = self.client.files.download(file=result_file_name)
            content_str = file_content.decode('utf-8')
            
            for line in content_str.splitlines():
                if not line.strip():
                    continue
                
                result_obj = json.loads(line)
                key = result_obj.get("key", "")
                
                if "response" in result_obj:
                    # Extract text from the response
                    response = result_obj["response"]
                    if "candidates" in response and len(response["candidates"]) > 0:
                        candidate = response["candidates"][0]
                        if "content" in candidate and "parts" in candidate["content"]:
                            parts = candidate["content"]["parts"]
                            text = "".join([part.get("text", "") for part in parts])
                            results.append({
                                "key": key,
                                "success": True,
                                "text": text
                            })
                            continue
                
                # Handle error cases or unexpected formats
                results.append({
                    "key": key,
                    "success": False,
                    "error": result_obj.get("error", "Unknown error"),
                    "raw": result_obj
                })
        
        # Sort results by key (which contains the original paragraph order)
        results.sort(key=lambda x: int(x["key"].split("-")[1]) if "-" in x["key"] else 0)
        
        logger.info(f"Retrieved {len(results)} results from batch job")
        return results
    
    def reconstruct_document(self, batch_results: List[Dict]) -> str:
        """
        Reconstruct a complete document from batch results.
        
        Args:
            batch_results: Results from get_batch_results()
            
        Returns:
            Reconstructed document text
        """
        # Extract successfully processed text chunks in order
        text_chunks = []
        for result in batch_results:
            if result.get("success", False) and "text" in result:
                text_chunks.append(result["text"])
        
        # Join with paragraph breaks
        document = "\n\n".join(text_chunks)
        return document
    
    def process_document_batch(self, 
                              document_text: str,
                              style_name: str,
                              wait_for_completion: bool = True,
                              max_wait_time: int = 3600) -> Dict:
        """
        Complete end-to-end batch processing of a document.
        
        Args:
            document_text: Full document text to process
            style_name: Citation style to apply
            wait_for_completion: Whether to wait for job completion
            max_wait_time: Maximum seconds to wait
            
        Returns:
            Dictionary with job information and results if completed
        """
        # Create batch job
        job_name = self.create_batch_job_for_formatting(
            document_text=document_text,
            style_name=style_name
        )
        
        result = {"job_name": job_name, "completed": False}
        
        # Wait for completion if requested
        if wait_for_completion:
            status = self.wait_for_batch_job(
                job_name=job_name,
                max_wait_time=max_wait_time
            )
            
            result.update(status)
            
            # Retrieve and process results if job succeeded
            if status["state"] == "JOB_STATE_SUCCEEDED":
                batch_results = self.get_batch_results(job_name)
                processed_document = self.reconstruct_document(batch_results)
                
                result.update({
                    "completed": True,
                    "document": processed_document,
                    "result_count": len(batch_results)
                })
        
        return result
    
    def cancel_batch_job(self, job_name: str) -> bool:
        """
        Cancel an ongoing batch job.
        
        Args:
            job_name: Batch job ID to cancel
            
        Returns:
            True if cancellation was successful
        """
        try:
            self.client.batches.cancel(name=job_name)
            logger.info(f"Cancelled batch job: {job_name}")
            return True
        except Exception as e:
            logger.error(f"Error cancelling job {job_name}: {e}")
            return False
    
    def cleanup(self):
        """Clean up temporary resources."""
        if self.temp_dir:
            self.temp_dir.cleanup()
