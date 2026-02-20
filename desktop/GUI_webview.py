import sys
import os
import json
import webview
import threading
import time
from pathlib import Path
import inspect
from dotenv import load_dotenv

# Import core logic
import core.formatter
from core.formatter import AdvancedFormatter
from core.api_clients import HuggingFaceClient, GeminiClient
from core.style_guides import STYLE_GUIDES
from utils.track_changes import TrackChanges


load_dotenv()

DATA_FILE = "user_data.json"

def validate_style(style_name: str) -> str:
    return style_name if style_name.lower() in STYLE_GUIDES else "apa"

def determine_output_path(input_path: Path, output_arg: str, style_name: str) -> Path:
    if output_arg:
        return Path(output_arg)
    formatting_dir = Path("formatted")
    formatting_dir.mkdir(exist_ok=True)
    return formatting_dir / f"{input_path.stem}_formatted_{style_name}{input_path.suffix}"

class Api:
    def __init__(self):
        self._window = None
        self._is_cancelled = False
        self._data = self._load_data()
        self._active_processing_count = 0

    def set_window(self, window):
        self._window = window

    def minimize_window(self):
        """Minimize the window"""
        if self._window:
            self._window.minimize()

    def maximize_window(self):
        """Maximize the window"""
        if self._window:
            self._window.maximize()

    def restore_window(self):
        """Restore the window"""
        if self._window:
            self._window.restore()
    
    def close_window(self):
        """Close the window"""
        if self._window:
            self._window.destroy()

    def browse_file(self):
        """Open file dialog and return selected file path"""
        file_types = ('Word Documents (*.docx)', )
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG, 
            allow_multiple=False, 
            file_types=file_types
        )
        return result[0] if result else None

    def start_formatting(self, options):
        """Start formatting in a separate thread"""
        self._is_cancelled = False
        self._active_processing_count += 1
        thread = threading.Thread(target=self._run_formatting, args=(options,))
        thread.daemon = True
        thread.start()
        return {"success": True}

    def cancel_formatting(self):
        """Cancel the current formatting operation"""
        self._is_cancelled = True
        return {"success": True}

    def open_file(self, path):
        """Open file in default application"""
        try:
            if os.name == 'nt' and os.path.exists(path):
                os.startfile(path)
                return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_history(self):
        """Get formatting history"""
        return self._data.get("history", [])

    def get_notifications(self):
        """Get notifications"""
        return self._data.get("notifications", [])

    def mark_notification_read(self, notification_id):
        """Mark a notification as read"""
        notifications = self._data.get("notifications", [])
        for n in notifications:
            if n.get("id") == notification_id:
                n["read"] = True
        self._save_data()
        return {"success": True}

    def clear_notifications(self):
        """Clear all notifications"""
        self._data["notifications"] = []
        self._save_data()
        return {"success": True}

    def get_preferences(self):
        """Get user preferences"""
        return self._data.get("settings", {})

    def save_preferences(self, prefs):
        """Save user preferences"""
        self._data["settings"] = prefs
        self._save_data()
        return {"success": True}

    def _load_data(self):
        """Load user data from JSON file"""
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    return json.load(f)
            except:
                pass
        return {"history": [], "settings": {}, "notifications": []}

    def _save_data(self):
        """Save user data to JSON file"""
        try:
            with open(DATA_FILE, "w") as f:
                json.dump(self._data, f, indent=2)
        except:
            pass

    def _add_to_history(self, filename, status, path, style="apa", size_bytes=0):
        """Add entry to history"""
        # Format size
        size_str = "-"
        if size_bytes > 0:
            if size_bytes < 1024:
                size_str = f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                size_str = f"{size_bytes/1024:.1f} KB"
            else:
                size_str = f"{size_bytes/(1024*1024):.1f} MB"

        entry = {
            "id": int(time.time() * 1000),
            "filename": filename,
            "status": status,
            "date": time.strftime("%Y-%m-%d %H:%M"),
            "path": str(path),
            "style": style.upper() if style else "-",
            "size": size_str
        }
        self._data.setdefault("history", []).insert(0, entry)
        # Keep last 50 history items
        self._data["history"] = self._data["history"][:50]
        self._save_data()

    def _add_notification(self, title, message, type="info"):
        """Add a notification"""
        notification = {
            "id": int(time.time() * 1000),
            "title": title,
            "message": message,
            "type": type,
            "date": time.strftime("%H:%M"),
            "read": False
        }
        self._data.setdefault("notifications", []).insert(0, notification)
        # Keep last 20 notifications
        self._data["notifications"] = self._data["notifications"][:20]
        self._save_data()
        
        # Trigger UI update if window exists
        if self._window:
            try:
                self._window.evaluate_js("if(window.loadNotifications) window.loadNotifications();")
            except:
                pass

    def _run_formatting(self, options):
        """Run the formatting process"""
        try:
            input_path = Path(options['input'])
            
            # Validate file exists
            if not input_path.exists():
                self._finish(False, "Selected file does not exist", None)
                return
            
            # Get file size
            try:
                file_size = input_path.stat().st_size
            except:
                file_size = 0
                
            style_name = validate_style(options['style'])
            output_path = determine_output_path(input_path, None, style_name)
            backend = options['backend']

            self._update_progress("Initializing backend...", 10)

            # Initialize AI client
            client = None
            if backend == 'gemini':
                if not os.getenv("GEMINI_API_KEY"):
                    raise ValueError("GEMINI_API_KEY is not set.")
                client = GeminiClient()
            else:
                if not os.getenv("HF_API_KEY"):
                    raise ValueError("HF_API_KEY is not set.")
                client = HuggingFaceClient()

            if self._is_cancelled:
                self._finish(False, "cancelled", None)
                return

            self._update_progress("Analyzing and formatting document...", 40)
            english_variant = options.get('english_variant', 'us')
            formatter = AdvancedFormatter(style_name, ai_client=client, english_variant=english_variant)

            # Set MLA options if applicable
            if style_name == 'mla':
                setattr(formatter, 'mla_heading', options.get('mla_heading', True))
                setattr(formatter, 'mla_title_page', options.get('mla_title_page', False))

            # Execute formatting and capture the actual output path (in case of retry/rename)
            final_output_path_str = formatter.format_document(str(input_path), str(output_path))
            # Update output_path to the actual file that was saved
            if final_output_path_str:
                output_path = Path(final_output_path_str)

            if self._is_cancelled:
                self._finish(False, "cancelled", None)
                return

            # Track changes if requested (using the actual output path)
            if options.get('track_changes'):
                self._update_progress("Tracking changes...", 80)
                tracker = TrackChanges(str(input_path.absolute()), str(output_path.absolute()))
                tracker.compare_docs()

            self._update_progress("Success!", 100)
            
            # Add to history with size and style
            self._add_to_history(input_path.name, "Completed", output_path, style=style_name, size_bytes=file_size)
            
            # Add notification
            self._add_notification(
                "Formatting Complete", 
                f"{input_path.name} was successfully formatted in {style_name.upper()} style.", 
                "success"
            )
            
            self._finish(True, "Document formatted successfully!", str(output_path))

        except Exception as e:
            self._finish(False, str(e), None)
            # Add failure notification
            self._add_notification("Formatting Failed", str(e), "error")
        finally:
            self._active_processing_count = max(0, self._active_processing_count - 1)

    def get_stats(self):
        """Get dashboard statistics"""
        history = self._data.get("history", [])
        
        # Calculate stats
        total_formatted = sum(1 for item in history if item.get("status") == "Completed")
        unique_documents = len(set(item.get("filename") for item in history))
        
        # Simple quota usage based on history count for now
        # Assuming a plan limit of 100
        plan_usage_count = len(history) 
        
        return {
            "documents": unique_documents,
            "processing": self._active_processing_count,
            "formatted": total_formatted,
            "plan_usage": f"{plan_usage_count}/100"
        }

    def get_user_info(self):
        """Get user profile info"""
        return {
            "name": "Faithman Founder",
            "plan": "Pro",
            "avatar": None
        }

    def _update_progress(self, message, progress):
        """Update progress in UI"""
        if self._window:
            try:
                self._window.evaluate_js(f"window.updateProgress('{message}', {progress})")
            except:
                pass

    def _finish(self, success, message, output_path):
        """Send completion message to UI"""
        if self._window:
            try:
                msg = message.replace("'", "\\'").replace('"', '\\"')
                path = str(output_path).replace("\\", "\\\\") if output_path else ""
                self._window.evaluate_js(
                    f"window.formattingFinished({'true' if success else 'false'}, '{msg}', '{path}')"
                )
            except:
                pass

def main():
    """Main application entry point"""
    api = Api()
    
    # Path to index.html
    current_dir = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(current_dir, 'interface', 'index.html')
    
    # Verify HTML file exists
    if not os.path.exists(html_path):
        print(f"Error: Interface file not found at {html_path}")
        sys.exit(1)
    
    window = webview.create_window(
        'Formatly',
        html_path,
        js_api=api,
        width=1200,
        height=800,
        min_size=(1000, 700),
        background_color='#050505',
        frameless=True,
        easy_drag=False
    )
    api.set_window(window)
    webview_data_dir = os.path.join(current_dir, '.webview_data')
    os.makedirs(webview_data_dir, exist_ok=True)
    webview.start(storage_path=webview_data_dir)

if __name__ == '__main__':
    main()