"""
Track Changes Utility
---------------------
Uses Microsoft Word (via COM automation/pywin32) to generate a "Track Changes"
comparison between the original input document and the formatted output.

Requirements:
    - Windows OS
    - Microsoft Word installed
    - `pywin32` library
"""

import win32com.client
import os
from pathlib import Path
import time

class TrackChanges:
    """
    Automates Word to compare two documents and save the result with tracked changes.
    """
    def __init__(self, input_file: Path, output_file: Path):
        self.input_file = input_file
        self.output_file = output_file

    def compare_docs(self):
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False

        try:
            # Open original document
            original_doc = word.Documents.Open(str(self.input_file))
            
            # Open formatted document
            formatted_doc = word.Documents.Open(str(self.output_file))

            # Create a new name for the comparison document
            output_dir = Path(self.output_file).parent
            output_name = Path(self.output_file).stem
            comparison_path = str(output_dir / f"{output_name}_Tracked.docx")

            # Compare documents
            comparison = word.CompareDocuments(original_doc, formatted_doc)

            if comparison:
                # Save the comparison as tracked file with a new name
                comparison.ActiveWindow.View.Type = 3  # Print Layout view
                comparison.SaveAs(comparison_path)
                print(f"Comparison saved to: {comparison_path}")
                
                # Close the comparison document
                comparison.Close(SaveChanges=False)
                
                # Open the comparison document
                os.startfile(comparison_path)
            else:
                print("No differences found between documents.")

        except Exception as e:
            print(f"Error comparing documents: {e}")
            raise
        finally:
            # Close all documents
            for doc in word.Documents:
                doc.Close(SaveChanges=False)
            word.Quit()
            time.sleep(3)


if __name__ == "__main__":
    # Example usage:
    # Place two sample .docx files in the repository `formatted/` folder
    # named `example_original.docx` and `example_formatted.docx` (or update paths below).
    # repo_root = Path(__file__).resolve().parent.parent
    # examples_input_dir = repo_root / "documents"
    # examples_output_dir = repo_root / "formatted"
    input_sample = "C:/Users/DELL XPS 9360/Desktop/School/B.Sc Project/Main/Hard/A Secure Web-Based Event Ticketing System Using QR Code Technology.docx"
    output_sample = "C:/Users/DELL XPS 9360/Documents/GitHub/formatly/formatted/A Secure Web-Based Event Ticketing System Using QR Code Technology_formatted_apa.docx"

    # if not input_sample.exists() or not output_sample.exists():
    #     print(
    #         f"Example files not found: {input_sample} or {output_sample}\n"
    #         "Please place sample .docx files at these locations or update the paths in this block."
    #     )
    # else:
    tracker = TrackChanges(input_sample, output_sample)
    tracker.compare_docs()