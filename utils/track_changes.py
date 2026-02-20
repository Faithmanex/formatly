"""
Track Changes Utility (ConvertAPI Version)
------------------------------------------
Uses ConvertAPI to perform a document comparison and generate a tracked changes file.
This version is platform-independent and works on Linux/Railway.

Requirements:
    - convertapi library
    - CONVERTAPI_API_KEY environment variable
"""

import convertapi
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class TrackChanges:
    """
    Uses ConvertAPI to compare two documents and save the result with tracked changes.
    """
    def __init__(self, input_file: str, output_file: str):
        """
        Initialize with paths to the original and formatted documents.
        
        Args:
            input_file: Path to the original (unformatted) document.
            output_file: Path to the formatted document.
        """
        self.input_file = input_file
        self.output_file = output_file
        
        # Set API credentials
        # Set API credentials
        api_key = os.getenv("CONVERTAPI_API_KEY")
        if not api_key:
            logger.warning("CONVERTAPI_API_KEY environment variable is not set. Track changes might fail.")
            # Don't raise immediately, let it fail when called if needed, or check elsewhere
        
        if api_key:
             convertapi.api_credentials = api_key

    def compare_docs(self, save_dir: str = None) -> str:
        """
        Compare the documents using ConvertAPI.
        
        Args:
            save_dir: Directory to save the resulting tracked changes file. 
                     Defaults to the same directory as the output_file.
                     
        Returns:
            The path to the saved comparison file.
        """
        try:
            input_path = Path(self.input_file)
            output_path = Path(self.output_file)
            
            if not save_dir:
                save_dir = str(output_path.parent)
            
            comparison_filename = f"{output_path.stem}_Tracked.docx"
            comparison_path = os.path.join(save_dir, comparison_filename)
            
            logger.info(f"Comparing {input_path.name} and {output_path.name} via ConvertAPI...")
            
            # Perform comparison
            result = convertapi.convert('compare', {
                'File': str(input_path),
                'CompareFile': str(output_path),
                'CompareLevel': 'Character',
                'RevisionAuthor': 'Formatly'
            }, from_format='docx')
            
            # Save the result
            result.save_files(save_dir)
            
            # Clean up/Rename saved file to consistent name if needed
            saved_files = result.list_files()
            if saved_files:
                api_saved_name = saved_files[0].filename
                api_saved_path = os.path.join(save_dir, api_saved_name)
                
                if api_saved_path != comparison_path:
                    if os.path.exists(comparison_path):
                        os.remove(comparison_path)
                    os.rename(api_saved_path, comparison_path)
            
            logger.info(f"Comparison saved to: {comparison_path}")
            
            # Auto-open on Windows desktop only
            if os.name == 'nt' and not os.getenv("RAILWAY_ENVIRONMENT"):
                try:
                    os.startfile(comparison_path)
                except Exception:
                    pass
                    
            return comparison_path

        except Exception as e:
            logger.error(f"Error comparing documents via ConvertAPI: {e}")
            raise

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('original', help='Original file')
    parser.add_argument('formatted', help='Formatted file')
    args = parser.parse_args()
    
    tracker = TrackChanges(args.original, args.formatted)
    tracker.compare_docs()